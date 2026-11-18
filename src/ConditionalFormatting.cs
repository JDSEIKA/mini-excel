using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;

namespace CalcRapido
{
    /// <summary>
    /// Regla de formato condicional leída de un .xlsx. Se aplica solo a la
    /// presentación; el archivo original se conserva sin cambios.
    /// </summary>
    public sealed class ConditionalRule
    {
        public string Formula;
        public CellStyle Style;
        public bool StopIfTrue;
        public int Priority;
        public bool IsDataBar;
        public Color DataBarColor = Color.Empty;
        public string Kind = "expression";
        public string Operator;
        public string TextValue;
        public string TimePeriod;
        public int Rank;
        public bool Percent;
        public bool Bottom;
        public string IconSetName;
        public readonly List<double> Thresholds = new List<double>();
        public readonly List<double> ThresholdPercentiles = new List<double>();
        public readonly List<string> ThresholdTypes = new List<string>();
        public readonly List<Color> ScaleColors = new List<Color>();
        public Color ScaleFill = Color.Empty;
        public string IconGlyph;
        public Color IconColor = Color.Empty;

        public ConditionalRule Clone()
        {
            ConditionalRule r = new ConditionalRule
            {
                Formula = Formula,
                Style = Style == null ? null : Style.Clone(),
                StopIfTrue = StopIfTrue,
                Priority = Priority,
                IsDataBar = IsDataBar,
                DataBarColor = DataBarColor,
                Kind = Kind,
                Operator = Operator,
                TextValue = TextValue,
                TimePeriod = TimePeriod,
                Rank = Rank,
                Percent = Percent,
                Bottom = Bottom,
                IconSetName = IconSetName,
                ScaleFill = ScaleFill,
                IconGlyph = IconGlyph,
                IconColor = IconColor
            };
            r.Thresholds.AddRange(Thresholds);
            r.ThresholdPercentiles.AddRange(ThresholdPercentiles);
            r.ThresholdTypes.AddRange(ThresholdTypes);
            r.ScaleColors.AddRange(ScaleColors);
            return r;
        }
    }

    public sealed class ConditionalFormat
    {
        public int Row1;
        public int Col1;
        public int Row2;
        public int Col2;
        public readonly List<ConditionalRule> Rules = new List<ConditionalRule>();

        public ConditionalFormat Clone()
        {
            ConditionalFormat c = new ConditionalFormat
            {
                Row1 = Row1, Col1 = Col1, Row2 = Row2, Col2 = Col2
            };
            foreach (ConditionalRule r in Rules) c.Rules.Add(r.Clone());
            return c;
        }
    }

    internal static class ConditionalEvaluator
    {
        private const int MaxAstCache = 2048;
        private static readonly Dictionary<string, FormulaNode> Cache =
            new Dictionary<string, FormulaNode>(StringComparer.Ordinal);

        private static FormulaNode ParseCached(string formula)
        {
            FormulaNode node;
            if (Cache.TryGetValue(formula, out node)) return node;
            try { node = FormulaEngine.ParseInternal(formula); }
            catch { node = null; }
            if (Cache.Count >= MaxAstCache) Cache.Clear();
            Cache[formula] = node;
            return node;
        }

        public static void Reset()
        {
            Cache.Clear();
            StatsCache.Clear();
            ResultCache.Clear();
        }
        private sealed class RangeStats
        {
            public int Version;
            public double Min = double.MaxValue;
            public double Max = double.MinValue;
            public List<double> Values;
        }

        private static readonly Dictionary<ConditionalFormat, RangeStats> StatsCache =
            new Dictionary<ConditionalFormat, RangeStats>();

        private sealed class CachedResult
        {
            public int Version;
            public ConditionalRule Rule;
        }

        private static readonly Dictionary<string, CachedResult> ResultCache =
            new Dictionary<string, CachedResult>(StringComparer.Ordinal);

        private static RangeStats Stats(Sheet sheet, ConditionalFormat cf)
        {
            RangeStats stats;
            if (StatsCache.TryGetValue(cf, out stats) && stats.Version == sheet.DataVersion)
                return stats;
            stats = new RangeStats { Version = sheet.DataVersion, Values = new List<double>() };
            for (int r = cf.Row1; r <= cf.Row2; r++)
                for (int c = cf.Col1; c <= cf.Col2; c++)
                {
                    Cell cell = sheet.GetCell(r, c);
                    if (cell == null || cell.Value.Kind != ValueKind.Number) continue;
                    double v = cell.Value.Number;
                    stats.Values.Add(v);
                    if (v < stats.Min) stats.Min = v;
                    if (v > stats.Max) stats.Max = v;
                }
            stats.Values.Sort();
            if (StatsCache.Count > 512) StatsCache.Clear();
            StatsCache[cf] = stats;
            return stats;
        }

        public static ConditionalRule Find(Sheet sheet, int row, int col)
        {
            if (sheet == null || sheet.ConditionalFormats.Count == 0) return null;
            string resultKey = sheet.GetHashCode().ToString() + "|" + row + "|" + col;
            CachedResult cached;
            if (ResultCache.TryGetValue(resultKey, out cached) &&
                cached.Version == sheet.DataVersion) return cached.Rule;
            ConditionalRule best = null;
            ConditionalRule icon = null;
            bool stopped = false;
            foreach (ConditionalFormat cf in sheet.ConditionalFormats)
            {
                if (row < cf.Row1 || row > cf.Row2 || col < cf.Col1 || col > cf.Col2) continue;
                foreach (ConditionalRule rule in cf.Rules)
                {
                    if (stopped) break;
                    if (rule.IsDataBar) continue;
                    ConditionalRule evaluated = rule;
                    bool match = false;
                    string kind = rule.Kind ?? "expression";
                    if (kind == "colorScale")
                    {
                        Color scale;
                        if (TryColorScale(cf, rule, sheet, row, col, out scale))
                        {
                            evaluated = rule.Clone();
                            evaluated.ScaleFill = scale;
                            match = true;
                        }
                    }
                    else if (kind == "iconSet")
                    {
                        string glyph;
                        Color color;
                        if (TryIconSet(cf, rule, sheet, row, col, out glyph, out color))
                        {
                            evaluated = rule.Clone();
                            evaluated.IconGlyph = glyph;
                            evaluated.IconColor = color;
                            match = true;
                        }
                    }
                    else if (kind == "top10")
                    {
                        match = TryTop10(cf, rule, sheet, row, col);
                    }
                    else if (kind == "cellIs")
                    {
                        match = TryCellIs(cf, rule, sheet, row, col);
                    }
                    else if (kind == "containsText" || kind == "notContainsText" ||
                             kind == "beginsWith" || kind == "endsWith")
                    {
                        match = TryContainsText(rule, sheet, row, col);
                    }
                    else if (kind == "timePeriod")
                    {
                        match = TryTimePeriod(rule, sheet, row, col);
                    }
                    else if (kind == "duplicateValues" || kind == "uniqueValues")
                    {
                        match = TryDuplicate(cf, rule, sheet, row, col);
                    }
                    else if (!string.IsNullOrEmpty(rule.Formula))
                    {
                        string shifted = FormulaEngine.TranslateFormula(
                            "=" + rule.Formula, row - cf.Row1, col - cf.Col1).TrimStart('=');
                        match = Matches(sheet, row, col, shifted);
                    }
                    if (!match) continue;
                    if (evaluated.IconGlyph != null)
                    {
                        if (icon == null || rule.Priority < icon.Priority) icon = evaluated;
                        if (evaluated.ScaleFill.IsEmpty &&
                            (evaluated.Style == null || evaluated.Style.FillColor.IsEmpty))
                        {
                            if (rule.StopIfTrue) { stopped = true; break; }
                            continue;
                        }
                    }
                    bool ruleFill = !evaluated.ScaleFill.IsEmpty ||
                        (evaluated.Style != null && !evaluated.Style.FillColor.IsEmpty);
                    bool bestFill = best != null &&
                        (!best.ScaleFill.IsEmpty ||
                         (best.Style != null && !best.Style.FillColor.IsEmpty));
                    if (best == null || (ruleFill && !bestFill) ||
                        (ruleFill == bestFill && rule.Priority < best.Priority))
                        best = evaluated;
                    if (rule.StopIfTrue) { stopped = true; break; }
                }
                if (stopped) break;
            }
            if (best != null && string.IsNullOrEmpty(best.IconGlyph) && icon != null)
            {
                best = best.Clone();
                best.IconGlyph = icon.IconGlyph;
                best.IconColor = icon.IconColor;
            }
            if (ResultCache.Count > 200000) ResultCache.Clear();
            ResultCache[resultKey] = new CachedResult { Version = sheet.DataVersion, Rule = best };
            return best;
        }

        private static bool TryColorScale(ConditionalFormat cf, ConditionalRule rule, Sheet sheet,
                                          int row, int col, out Color color)
        {
            color = Color.Empty;
            double value;
            if (!NumberAt(sheet, row, col, out value)) return false;
            if (rule.ScaleColors.Count < 2) return false;
            RangeStats stats = Stats(sheet, cf);
            double min = stats.Min, max = stats.Max;
            if (stats.Values.Count == 0) { min = value; max = value; }
            List<double> points = ThresholdValues(cf, rule, min, max, value);
            if (points.Count < 2) { color = rule.ScaleColors[0]; return true; }
            int segment = 0;
            for (int i = 1; i < points.Count; i++)
                if (value >= points[i]) segment = i;
            int a = Math.Min(segment, rule.ScaleColors.Count - 1);
            int b = Math.Min(a + 1, rule.ScaleColors.Count - 1);
            if (a == b) { color = rule.ScaleColors[a]; return true; }
            double span = points[Math.Min(b, points.Count - 1)] - points[a];
            double f = span <= 0 ? 0 : (value - points[a]) / span;
            f = Math.Max(0, Math.Min(1, f));
            color = Lerp(rule.ScaleColors[a], rule.ScaleColors[b], f);
            return true;
        }

        private static List<double> ThresholdValues(ConditionalFormat cf, ConditionalRule rule,
                                                    double min, double max, double cellValue)
        {
            List<double> points = new List<double>();
            for (int i = 0; i < rule.Thresholds.Count; i++)
            {
                string type = i < rule.ThresholdTypes.Count ? rule.ThresholdTypes[i] : "num";
                if (type == "min") points.Add(min);
                else if (type == "max") points.Add(max);
                else if (type == "percent")
                    points.Add(min + (max - min) * rule.Thresholds[i] / 100.0);
                else if (type == "percentile")
                    points.Add(Percentile(cf, rule, sheet: null, percentile: rule.Thresholds[i],
                        min: min, max: max));
                else points.Add(rule.Thresholds[i]);
            }
            return points;
        }

        private static double Percentile(ConditionalFormat cf, ConditionalRule rule, Sheet sheet,
                                         double percentile, double min, double max)
        {
            return min + (max - min) * Math.Max(0, Math.Min(1, percentile / 100.0));
        }

        private static Color Lerp(Color a, Color b, double f)
        {
            return Color.FromArgb(
                (int)Math.Round(a.R + (b.R - a.R) * f),
                (int)Math.Round(a.G + (b.G - a.G) * f),
                (int)Math.Round(a.B + (b.B - a.B) * f));
        }

        private static bool TryIconSet(ConditionalFormat cf, ConditionalRule rule, Sheet sheet,
                                       int row, int col, out string glyph, out Color color)
        {
            glyph = null; color = Color.Empty;
            double value;
            if (!NumberAt(sheet, row, col, out value)) return false;
            RangeStats stats = Stats(sheet, cf);
            double min = stats.Min, max = stats.Max;
            if (stats.Values.Count == 0) { min = value; max = value; }
            List<double> points = ThresholdValues(cf, rule, min, max, value);
            int index = 0;
            for (int i = 0; i < points.Count; i++) if (value >= points[i]) index = i;
            int count = Math.Max(3, Math.Min(5, points.Count + 1));
            index = Math.Min(index, count - 1);
            string name = rule.IconSetName ?? "3TrafficLights1";
            if (name.StartsWith("3", StringComparison.Ordinal) || name.StartsWith("4", StringComparison.Ordinal) ||
                name.StartsWith("5", StringComparison.Ordinal))
            {
                string[] arrows = new string[] { "▼", "●", "▲" };
                string[] colors = new string[] { "#C00000", "#ED7D31", "#548235" };
                glyph = index < arrows.Length ? arrows[index] : "●";
                color = ParseHex(index < colors.Length ? colors[index] : "#548235");
            }
            else
            {
                glyph = index == 0 ? "●" : index == 1 ? "◐" : "○";
                color = index == 0 ? Color.FromArgb(192, 0, 0) :
                        index == 1 ? Color.FromArgb(237, 125, 49) : Color.FromArgb(84, 130, 53);
            }
            return true;
        }

        private static bool TryTop10(ConditionalFormat cf, ConditionalRule rule, Sheet sheet, int row, int col)
        {
            double value;
            if (!NumberAt(sheet, row, col, out value)) return false;
            RangeStats stats = Stats(sheet, cf);
            List<double> values = stats.Values;
            if (values.Count == 0) return false;
            int count = rule.Percent ? Math.Max(1, values.Count * rule.Rank / 100) : Math.Max(1, rule.Rank);
            count = Math.Min(count, values.Count);
            if (rule.Bottom) return value <= values[count - 1];
            return value >= values[values.Count - count];
        }

        private static bool TryCellIs(ConditionalFormat cf, ConditionalRule rule, Sheet sheet, int row, int col)
        {
            double value;
            if (!NumberAt(sheet, row, col, out value)) return false;
            double threshold;
            if (!TryThreshold(rule.Formula, sheet, row, col, out threshold)) return false;
            string op = rule.Operator ?? "equal";
            if (op == "greaterThan") return value > threshold;
            if (op == "greaterThanOrEqual") return value >= threshold;
            if (op == "lessThan") return value < threshold;
            if (op == "lessThanOrEqual") return value <= threshold;
            if (op == "notEqual") return Math.Abs(value - threshold) > 1e-9;
            return Math.Abs(value - threshold) <= 1e-9;
        }

        private static bool TryContainsText(ConditionalRule rule, Sheet sheet, int row, int col)
        {
            Cell cell = sheet.GetCell(row, col);
            string text = cell == null ? "" : cell.Value.ToDisplayString();
            string needle = rule.TextValue ?? "";
            string kind = rule.Kind ?? "";
            if (kind == "notContainsText") return text.IndexOf(needle, StringComparison.OrdinalIgnoreCase) < 0;
            if (kind == "beginsWith") return text.StartsWith(needle, StringComparison.OrdinalIgnoreCase);
            if (kind == "endsWith") return text.EndsWith(needle, StringComparison.OrdinalIgnoreCase);
            return text.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool TryTimePeriod(ConditionalRule rule, Sheet sheet, int row, int col)
        {
            double serial;
            if (!NumberAt(sheet, row, col, out serial)) return false;
            DateTime date;
            try { date = DateTime.FromOADate(serial).Date; }
            catch { return false; }
            DateTime today = DateTime.Today;
            string p = rule.TimePeriod ?? "";
            if (p == "today") return date == today;
            if (p == "yesterday") return date == today.AddDays(-1);
            if (p == "tomorrow") return date == today.AddDays(1);
            if (p == "thisWeek") return date >= today.AddDays(-(int)today.DayOfWeek + 1) && date <= today.AddDays(7 - (int)today.DayOfWeek);
            if (p == "lastWeek") return date >= today.AddDays(-(int)today.DayOfWeek - 6) && date <= today.AddDays(-(int)today.DayOfWeek);
            if (p == "nextWeek") return date >= today.AddDays(8 - (int)today.DayOfWeek) && date <= today.AddDays(14 - (int)today.DayOfWeek);
            if (p == "thisMonth") return date.Month == today.Month && date.Year == today.Year;
            if (p == "lastMonth")
            {
                DateTime m = today.AddMonths(-1);
                return date.Month == m.Month && date.Year == m.Year;
            }
            if (p == "nextMonth")
            {
                DateTime m = today.AddMonths(1);
                return date.Month == m.Month && date.Year == m.Year;
            }
            return false;
        }

        private static bool TryDuplicate(ConditionalFormat cf, ConditionalRule rule, Sheet sheet, int row, int col)
        {
            Cell cell = sheet.GetCell(row, col);
            if (cell == null || cell.Value.Kind == ValueKind.Empty) return false;
            string text = cell.Value.ToDisplayString();
            int count = 0;
            for (int r = cf.Row1; r <= cf.Row2; r++)
                for (int c = cf.Col1; c <= cf.Col2; c++)
                {
                    Cell other = sheet.GetCell(r, c);
                    if (other != null && string.Equals(other.Value.ToDisplayString(), text,
                        StringComparison.OrdinalIgnoreCase)) count++;
                }
            return rule.Kind == "uniqueValues" ? count == 1 : count > 1;
        }

        private static bool TrySimpleComparison(Sheet sheet, string formula, out bool result)
        {
            result = false;
            string f = formula.Trim();
            if (f.StartsWith("=", StringComparison.Ordinal)) f = f.Substring(1).Trim();
            string op = null;
            int at = -1;
            for (int i = 0; i < f.Length; i++)
            {
                char c = f[i];
                if (c == '<' || c == '>' || c == '=')
                {
                    if (c != '=' && i + 1 < f.Length && f[i + 1] == '=') { op = f.Substring(i, 2); at = i; break; }
                    if (c == '<' && i + 1 < f.Length && f[i + 1] == '>') { op = "<>"; at = i; break; }
                    op = c.ToString(); at = i; break;
                }
            }
            if (op == null || at <= 0) return false;
            string left = f.Substring(0, at).Trim().Replace("$", "");
            string right = f.Substring(at + op.Length).Trim();
            int cellRow, cellCol;
            bool absRow, absCol;
            if (!AddressUtil.TryParseCell(left, out cellRow, out cellCol, out absRow, out absCol)) return false;
            double target;
            if (!double.TryParse(right, NumberStyles.Float, CultureInfo.InvariantCulture, out target))
            {
                double local;
                if (!double.TryParse(right, NumberStyles.Float, CultureInfo.CurrentCulture, out local)) return false;
                target = local;
            }
            Cell cell = sheet.GetCell(cellRow, cellCol);
            if (cell == null || cell.Value.Kind != ValueKind.Number) return true;
            double value = cell.Value.Number;
            if (op == "<") result = value < target;
            else if (op == ">") result = value > target;
            else if (op == "<=") result = value <= target;
            else if (op == ">=") result = value >= target;
            else if (op == "<>") result = Math.Abs(value - target) > 1e-9;
            else result = Math.Abs(value - target) <= 1e-9;
            return true;
        }

        private static bool TryThreshold(string formula, Sheet sheet, int row, int col, out double value)
        {
            value = 0;
            if (string.IsNullOrEmpty(formula)) return false;
            string f = formula.Trim();
            if (f.StartsWith("=", StringComparison.Ordinal)) f = f.Substring(1);
            FormulaNode node = ParseCached(f);
            if (node == null) return false;
            try
            {
                EvalContext ctx = new EvalContext(sheet.Workbook);
                ctx.CurrentSheet = sheet; ctx.CurrentRow = row; ctx.CurrentCol = col;
                CellValue v = node.Eval(ctx);
                if (v.Kind == ValueKind.Number) { value = v.Number; return true; }
            }
            catch { }
            return false;
        }

        private static bool NumberAt(Sheet sheet, int row, int col, out double value)
        {
            value = 0;
            Cell cell = sheet.GetCell(row, col);
            if (cell == null || cell.Value.Kind != ValueKind.Number) return false;
            value = cell.Value.Number;
            return true;
        }

        private static Color ParseHex(string text)
        {
            try
            {
                return Color.FromArgb(
                    Convert.ToInt32(text.Substring(1, 2), 16),
                    Convert.ToInt32(text.Substring(3, 2), 16),
                    Convert.ToInt32(text.Substring(5, 2), 16));
            }
            catch { return Color.Gray; }
        }

        public static bool Matches(Sheet sheet, int row, int col, string formula)
        {
            if (sheet == null || string.IsNullOrEmpty(formula)) return false;
            bool simple;
            if (TrySimpleComparison(sheet, formula, out simple)) return simple;
            string translated = formula
                .Replace("task_start", "E" + (row + 1))
                .Replace("task_end", "F" + (row + 1))
                .Replace("task_progress", "D" + (row + 1));
            if (translated.StartsWith("=", StringComparison.Ordinal))
                translated = translated.Substring(1);
            FormulaNode node = ParseCached(translated);
            if (node == null) return false;
            try
            {
                EvalContext ctx = new EvalContext(sheet.Workbook);
                ctx.CurrentSheet = sheet;
                ctx.CurrentRow = row;
                ctx.CurrentCol = col;
                CellValue value = node.Eval(ctx);
                if (value.Kind == ValueKind.Bool) return value.Bool;
                bool b;
                CellValue err;
                return EvalUtil.TryBool(value, ctx, out b, out err) && b;
            }
            catch { return false; }
        }
    }
}
