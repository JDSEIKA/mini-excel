using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;

namespace CalcRapido
{
    public enum ChartKind
    {
        Column,
        Bar,
        Line,
        Area,
        Pie,
        Doughnut,
        Scatter,
        Radar,
        Combo,
        Unknown
    }

    public sealed class ChartSeries
    {
        public int Index;
        public ChartKind Kind = ChartKind.Unknown;
        public bool Secondary;
        public string Name;
        public string NameFormula;
        public string CategoriesFormula;
        public string ValuesFormula;
        public string XValuesFormula;
        public Color Color = Color.Empty;
        public readonly List<double> CachedValues = new List<double>();
        public readonly List<double> CachedXValues = new List<double>();
        public readonly List<string> CachedCategories = new List<string>();

        public ChartSeries Clone()
        {
            ChartSeries s = new ChartSeries
            {
                Index = Index,
                Kind = Kind,
                Secondary = Secondary,
                Name = Name, NameFormula = NameFormula,
                CategoriesFormula = CategoriesFormula,
                ValuesFormula = ValuesFormula,
                XValuesFormula = XValuesFormula,
                Color = Color
            };
            s.CachedValues.AddRange(CachedValues);
            s.CachedXValues.AddRange(CachedXValues);
            s.CachedCategories.AddRange(CachedCategories);
            return s;
        }
    }

    public sealed class SheetImage
    {
        public Image Image;
        public int FromRow;
        public int FromCol;
        public int ToRow = -1;
        public int ToCol = -1;
        public long FromColOff;
        public long FromRowOff;
        public long ToColOff;
        public long ToRowOff;
        public int WidthPx = 200;
        public int HeightPx = 150;
        public string SourcePart;
        public string SourceDrawingPart;
        public string SourceRelId;
        public bool Dirty;

        public SheetImage Clone()
        {
            return (SheetImage)MemberwiseClone();
        }
    }

    public sealed class SheetShape
    {
        public string Name;
        public string Text;
        public string Hyperlink;
        public string HyperlinkInternal;
        public int FromRow;
        public int FromCol;
        public int ToRow = -1;
        public int ToCol = -1;
        public long FromColOff;
        public long FromRowOff;
        public long ToColOff;
        public long ToRowOff;
        public Color FillColor = Color.FromArgb(68, 114, 196);
        public Color BorderColor = Color.FromArgb(47, 82, 143);
        public Color FontColor = Color.White;
        public float FontSize = 9f;

        public SheetShape Clone()
        {
            return (SheetShape)MemberwiseClone();
        }
    }

    public sealed class PivotCache
    {
        public readonly List<string> FieldNames = new List<string>();
        public readonly List<List<CellValue>> Rows = new List<List<CellValue>>();

        public PivotCache Clone()
        {
            PivotCache c = new PivotCache();
            c.FieldNames.AddRange(FieldNames);
            foreach (List<CellValue> row in Rows)
                c.Rows.Add(new List<CellValue>(row));
            return c;
        }
    }

    public sealed class PivotRange
    {
        public int Row1;
        public int Col1;
        public int Row2;
        public int Col2;
        public string StyleName;
        public bool ShowRowHeaders = true;
        public bool ShowColHeaders = true;
        public bool ShowRowStripes = true;
        public string RowCaption;
        public int RowField = -1;
        public int ColField = -1;
        public int FirstDataCol = 1;
        public readonly List<int> RowFields = new List<int>();
        public readonly List<int> DataFields = new List<int>();
        public readonly List<string> DataNames = new List<string>();
        public PivotCache Cache;

        public PivotRange Clone()
        {
            PivotRange p = (PivotRange)MemberwiseClone();
            p.RowFields.Clear(); p.RowFields.AddRange(RowFields);
            p.DataFields.Clear(); p.DataFields.AddRange(DataFields);
            p.DataNames.Clear(); p.DataNames.AddRange(DataNames);
            p.Cache = Cache == null ? null : Cache.Clone();
            return p;
        }
    }

    public sealed class SheetSlicer
    {
        public string Name;
        public string Caption;
        public string FieldName;
        public string SelectedValue;
        public int FromRow;
        public int FromCol;
        public int ToRow = -1;
        public int ToCol = -1;
        public long FromColOff;
        public long FromRowOff;
        public long ToColOff;
        public long ToRowOff;

        public SheetSlicer Clone()
        {
            return (SheetSlicer)MemberwiseClone();
        }
    }

    public sealed class SheetChart
    {
        public ChartKind Kind = ChartKind.Column;
        public string Title = "";
        public bool ShowLegend = true;
        public bool LegendTop;
        public bool ShowAxes = true;
        public bool Stacked;
        public bool PlotVisibleOnly = true;
        public bool OfPie;
        public string OfPieType = "pie";
        public string SplitType;
        public double SplitPos;
        public double SecondPieSize = 75;
        public readonly List<ChartSeries> Series = new List<ChartSeries>();
        public int FromRow;
        public int FromCol;
        public int ToRow = -1;
        public int ToCol = -1;
        public long FromColOff;
        public long FromRowOff;
        public long ToColOff;
        public long ToRowOff;
        public int WidthPx = 400;
        public int HeightPx = 260;
        public string SourceChartPart;
        public string SourceDrawingPart;
        public bool Imported;
        public bool Dirty;
        public string LastError;
        public Color[] ThemePalette;

        public SheetChart Clone()
        {
            SheetChart c = new SheetChart
            {
                Kind = Kind, Title = Title, ShowLegend = ShowLegend, LegendTop = LegendTop,
                ShowAxes = ShowAxes,
                Stacked = Stacked, PlotVisibleOnly = PlotVisibleOnly,
                OfPie = OfPie, OfPieType = OfPieType, SplitType = SplitType,
                SplitPos = SplitPos, SecondPieSize = SecondPieSize,
                FromRow = FromRow, FromCol = FromCol, ToRow = ToRow, ToCol = ToCol,
                FromColOff = FromColOff, FromRowOff = FromRowOff,
                ToColOff = ToColOff, ToRowOff = ToRowOff,
                WidthPx = WidthPx, HeightPx = HeightPx,
                SourceChartPart = SourceChartPart, SourceDrawingPart = SourceDrawingPart,
                Imported = Imported, Dirty = Dirty, LastError = LastError,
                ThemePalette = ThemePalette == null ? null : (Color[])ThemePalette.Clone()
            };
            foreach (ChartSeries s in Series) c.Series.Add(s.Clone());
            return c;
        }
    }

    public static class ChartData
    {
        public static Color[] ThemePalette;

        public static RangeRef ResolveRange(Workbook wb, Sheet context, string formula)
        {
            if (wb == null || string.IsNullOrEmpty(formula)) return null;
            return ResolveRange(wb, context, formula, 0);
        }

        private static RangeRef ResolveRange(Workbook wb, Sheet context, string formula, int depth)
        {
            if (depth > 8) return null;
            string f = formula.Trim();
            if (f.StartsWith("=", StringComparison.Ordinal)) f = f.Substring(1).Trim();
            if (f.Length == 0 || f.StartsWith("{", StringComparison.Ordinal)) return null;

            // Nombres definidos.
            DefinedName name = FindName(wb, f, context);
            if (name != null) return ResolveRange(wb, context, name.Formula, depth + 1);

            int bang = LastBang(f);
            Sheet target = context;
            string refText = f;
            if (bang >= 0)
            {
                string sheetText = f.Substring(0, bang).Trim();
                refText = f.Substring(bang + 1).Trim();
                if (sheetText.StartsWith("[", StringComparison.Ordinal))
                {
                    int close = sheetText.IndexOf(']');
                    if (close >= 0) sheetText = sheetText.Substring(close + 1);
                }
                if (sheetText.Length >= 2 && sheetText[0] == '\'' && sheetText[sheetText.Length - 1] == '\'')
                    sheetText = sheetText.Substring(1, sheetText.Length - 2).Replace("''", "'");
                foreach (Sheet s in wb.Sheets)
                    if (string.Equals(s.Name, sheetText, StringComparison.OrdinalIgnoreCase)) { target = s; break; }
                if (target == null) return null;
            }

            string[] parts = refText.Split(':');
            if (parts.Length > 2) return null;
            int r1, c1, r2, c2;
            bool ar, ac;
            if (!AddressUtil.TryParseCell(parts[0].Replace("$", ""), out r1, out c1, out ar, out ac))
                return null;
            if (parts.Length == 1) { r2 = r1; c2 = c1; }
            else if (!AddressUtil.TryParseCell(parts[1].Replace("$", ""), out r2, out c2, out ar, out ac))
                return null;
            if (r1 > r2) { int t = r1; r1 = r2; r2 = t; }
            if (c1 > c2) { int t = c1; c1 = c2; c2 = t; }
            return new RangeRef(target, r1, c1, r2, c2);
        }

        private static int LastBang(string s)
        {
            bool quote = false;
            int result = -1;
            for (int i = 0; i < s.Length; i++)
            {
                if (s[i] == '\'') quote = !quote;
                else if (s[i] == '!' && !quote) result = i;
            }
            return result;
        }

        private static DefinedName FindName(Workbook wb, string name, Sheet context)
        {
            foreach (DefinedName d in wb.DefinedNames)
            {
                if (!string.Equals(d.Name, name, StringComparison.OrdinalIgnoreCase)) continue;
                if (d.LocalSheetId < 0) return d;
                if (context != null && d.LocalSheetId == wb.Sheets.IndexOf(context)) return d;
            }
            return null;
        }

        public static string[] ReadTexts(Workbook wb, Sheet context, string formula,
                                         bool visibleOnly = false)
        {
            RangeRef r = ResolveRange(wb, context, formula);
            List<string> list = new List<string>();
            if (r == null) return list.ToArray();
            for (int row = r.Row1; row <= r.Row2; row++)
            {
                if (visibleOnly && r.Sheet.HiddenRows.Contains(row)) continue;
                for (int col = r.Col1; col <= r.Col2; col++)
                {
                    if (visibleOnly && r.Sheet.HiddenCols.Contains(col)) continue;
                    Cell c = r.Sheet.GetCell(row, col);
                    if (c == null || c.Value.Kind == ValueKind.Empty) list.Add("");
                    else list.Add(NumberFormatter.Format(c.Value, c.EffectiveStyle));
                }
            }
            return list.ToArray();
        }

        public static double[] ReadNumbers(Workbook wb, Sheet context, string formula,
                                           bool visibleOnly = false)
        {
            RangeRef r = ResolveRange(wb, context, formula);
            List<double> list = new List<double>();
            if (r == null) return list.ToArray();
            for (int row = r.Row1; row <= r.Row2; row++)
            {
                if (visibleOnly && r.Sheet.HiddenRows.Contains(row)) continue;
                for (int col = r.Col1; col <= r.Col2; col++)
                {
                    if (visibleOnly && r.Sheet.HiddenCols.Contains(col)) continue;
                    Cell c = r.Sheet.GetCell(row, col);
                    list.Add(c != null && c.Value.Kind == ValueKind.Number ? c.Value.Number : 0);
                }
            }
            return list.ToArray();
        }

        public static string SeriesName(Workbook wb, Sheet context, ChartSeries series, int index)
        {
            if (!string.IsNullOrEmpty(series.Name)) return series.Name;
            if (!string.IsNullOrEmpty(series.NameFormula))
            {
                RangeRef r = ResolveRange(wb, context, series.NameFormula);
                if (r != null && r.Row1 == r.Row2 && r.Col1 == r.Col2)
                {
                    Cell c = r.Sheet.GetCell(r.Row1, r.Col1);
                    if (c != null && c.Value.Kind != ValueKind.Empty)
                        return NumberFormatter.Format(c.Value, c.EffectiveStyle);
                }
            }
            return "Serie " + (index + 1);
        }
    }

    public static class ChartXml
    {
        public static SheetChart Parse(XDocument doc)
        {
            if (doc == null || doc.Root == null) return null;
            XElement chart = First(doc.Root, "chart");
            XElement plot = chart == null ? null : First(chart, "plotArea");
            if (plot == null) return null;
            SheetChart result = new SheetChart();
            string chartType = ChartTypeName(plot);
            result.Kind = KindForType(chartType);
            XElement typeEl = chartType == null ? null : First(plot, chartType);
            if (typeEl != null)
            {
                if (chartType == "ofPieChart")
                {
                    result.OfPie = true;
                    result.OfPieType = Attr(First(typeEl, "ofPieType"), "val") ?? "pie";
                    result.SplitType = Attr(First(typeEl, "splitType"), "val");
                    double split;
                    if (double.TryParse(Attr(First(typeEl, "splitPos"), "val"),
                        NumberStyles.Float, CultureInfo.InvariantCulture, out split))
                        result.SplitPos = split;
                    double size;
                    if (double.TryParse(Attr(First(typeEl, "secondPieSize"), "val"),
                        NumberStyles.Float, CultureInfo.InvariantCulture, out size))
                        result.SecondPieSize = Math.Max(20, Math.Min(150, size));
                }
                if (chartType == "barChart")
                {
                    XElement dir = First(typeEl, "barDir");
                    if (dir != null && string.Equals(Attr(dir, "val"), "bar",
                        StringComparison.OrdinalIgnoreCase)) result.Kind = ChartKind.Bar;
                }
                XElement grouping = First(typeEl, "grouping");
                string g = grouping == null ? "" : Attr(grouping, "val");
                result.Stacked = string.Equals(g, "stacked", StringComparison.OrdinalIgnoreCase) ||
                                 string.Equals(g, "percentStacked", StringComparison.OrdinalIgnoreCase);
            }
            XElement title = First(chart, "title");
            result.Title = TextOf(title);
            XElement legendEl = First(chart, "legend");
            result.ShowLegend = legendEl != null;
            XElement legendPos = legendEl == null ? null : First(legendEl, "legendPos");
            result.LegendTop = legendPos != null &&
                string.Equals(Attr(legendPos, "val"), "t", StringComparison.OrdinalIgnoreCase);
            result.ShowAxes = First(plot, "catAx") != null || First(plot, "valAx") != null;
            XElement plotVis = First(chart, "plotVisOnly");
            result.PlotVisibleOnly = plotVis == null || Attr(plotVis, "val") != "0";

            if (result.Kind == ChartKind.Combo)
            {
                foreach (XElement child in Children(plot, "barChart"))
                {
                    XElement dir = First(child, "barDir");
                    ChartKind kind = dir != null &&
                        string.Equals(Attr(dir, "val"), "bar", StringComparison.OrdinalIgnoreCase)
                        ? ChartKind.Bar : ChartKind.Column;
                    foreach (XElement ser in Children(child, "ser"))
                        result.Series.Add(ParseSeriesElement(ser, result, kind, result.Series.Count));
                }
                bool secondary = Children(plot, "valAx").Count() >= 2;
                foreach (XElement child in Children(plot, "lineChart"))
                    foreach (XElement ser in Children(child, "ser"))
                    {
                        ChartSeries line = ParseSeriesElement(ser, result, ChartKind.Line, result.Series.Count);
                        line.Secondary = secondary;
                        result.Series.Add(line);
                    }
            }
            else
            {
                foreach (XElement ser in Children(typeEl, "ser"))
                    result.Series.Add(ParseSeriesElement(ser, result, result.Kind, result.Series.Count));
            }
            result.Series.Sort(delegate(ChartSeries a, ChartSeries b)
            {
                return a.Index.CompareTo(b.Index);
            });
            return result;
        }

        private static ChartSeries ParseSeriesElement(XElement ser, SheetChart chart,
                                                     ChartKind kind, int fallbackIndex)
        {
            ChartSeries s = new ChartSeries();
            int seriesIndex;
            s.Index = int.TryParse(Attr(First(ser, "idx"), "val"), out seriesIndex)
                ? seriesIndex : fallbackIndex;
            s.Kind = kind;
            string seriesText = SeriesText(ser);
            s.Name = LooksLikeReference(seriesText) ? "" : seriesText;
            s.NameFormula = SeriesRef(ser, "tx");
            s.CategoriesFormula = SeriesRef(ser, "cat");
            if (kind == ChartKind.Scatter || chart.Kind == ChartKind.Scatter)
            {
                s.XValuesFormula = SeriesRef(ser, "xVal");
                s.ValuesFormula = SeriesRef(ser, "yVal");
            }
            else
            {
                s.ValuesFormula = SeriesRef(ser, "val");
            }
            s.Color = SeriesColor(ser);
            s.CachedValues.AddRange(CachedNumbers(ser, kind == ChartKind.Scatter ? "yVal" : "val"));
            if (kind == ChartKind.Scatter) s.CachedXValues.AddRange(CachedNumbers(ser, "xVal"));
            s.CachedCategories.AddRange(CachedTexts(ser, "cat"));
            return s;
        }

        private static readonly string[] ChartTypes = new string[]
        {
            "barChart", "lineChart", "areaChart", "pieChart", "ofPieChart", "doughnutChart",
            "scatterChart", "bubbleChart", "radarChart", "comboChart", "histogramChart",
            "waterfallChart", "funnelChart", "treemapChart", "sunburstChart", "surfaceChart",
            "stockChart", "boxWhiskerChart", "regionMapChart"
        };

        private static string ChartTypeName(XElement plot)
        {
            foreach (string name in ChartTypes)
                if (First(plot, name) != null) return name;
            return null;
        }

        private static ChartKind KindForType(string type)
        {
            if (string.IsNullOrEmpty(type)) return ChartKind.Unknown;
            if (type == "barChart") return ChartKind.Column;
            if (type == "lineChart" || type == "surfaceChart" || type == "stockChart" ||
                type == "boxWhiskerChart") return ChartKind.Line;
            if (type == "areaChart") return ChartKind.Area;
            if (type == "pieChart" || type == "ofPieChart") return ChartKind.Pie;
            if (type == "doughnutChart" || type == "sunburstChart") return ChartKind.Doughnut;
            if (type == "scatterChart" || type == "bubbleChart") return ChartKind.Scatter;
            if (type == "radarChart") return ChartKind.Radar;
            if (type == "comboChart") return ChartKind.Combo;
            return ChartKind.Column;
        }

        private static ChartKind ParseKind(XElement plot)
        {
            if (First(plot, "barChart") != null)
            {
                XElement dir = First(First(plot, "barChart"), "barDir");
                return dir != null && string.Equals(Attr(dir, "val"), "bar", StringComparison.OrdinalIgnoreCase)
                    ? ChartKind.Bar : ChartKind.Column;
            }
            if (First(plot, "lineChart") != null) return ChartKind.Line;
            if (First(plot, "areaChart") != null) return ChartKind.Area;
            if (First(plot, "pieChart") != null || First(plot, "ofPieChart") != null)
                return ChartKind.Pie;
            if (First(plot, "doughnutChart") != null) return ChartKind.Doughnut;
            if (First(plot, "scatterChart") != null) return ChartKind.Scatter;
            if (First(plot, "radarChart") != null) return ChartKind.Radar;
            if (First(plot, "comboChart") != null) return ChartKind.Combo;
            return ChartKind.Unknown;
        }

        private static string KindElement(ChartKind kind)
        {
            switch (kind)
            {
                case ChartKind.Bar:
                case ChartKind.Column: return "barChart";
                case ChartKind.Line: return "lineChart";
                case ChartKind.Area: return "areaChart";
                case ChartKind.Pie: return "pieChart";
                case ChartKind.Doughnut: return "doughnutChart";
                case ChartKind.Scatter: return "scatterChart";
                case ChartKind.Radar: return "radarChart";
                default: return "";
            }
        }

        private static List<double> CachedNumbers(XElement ser, string name)
        {
            List<double> list = new List<double>();
            XElement el = First(ser, name);
            XElement numRef = First(el, "numRef");
            XElement cache = First(numRef, "numCache");
            if (cache == null) return list;
            foreach (XElement pt in Children(cache, "pt"))
            {
                XElement v = First(pt, "v");
                double d;
                if (v != null && double.TryParse(v.Value, NumberStyles.Float,
                    CultureInfo.InvariantCulture, out d)) list.Add(d);
                else list.Add(0);
            }
            return list;
        }

        private static List<string> CachedTexts(XElement ser, string name)
        {
            List<string> list = new List<string>();
            XElement el = First(ser, name);
            XElement strRef = First(el, "strRef");
            XElement cache = First(strRef, "strCache");
            if (cache == null) return list;
            foreach (XElement pt in Children(cache, "pt"))
            {
                XElement v = First(pt, "v");
                list.Add(v == null ? "" : v.Value);
            }
            return list;
        }

        private static bool LooksLikeReference(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;
            return text[0] == '=' || text[0] == '\'' || text.IndexOf('!') >= 0 ||
                   text.IndexOf('$') >= 0;
        }

        private static string SeriesText(XElement ser)
        {
            XElement tx = First(ser, "tx");
            if (tx == null) return "";
            XElement strRef = First(tx, "strRef");
            if (strRef != null)
            {
                XElement f = First(strRef, "f");
                if (f != null) return f.Value;
                XElement cache = First(strRef, "strCache");
                if (cache != null) return TextOf(cache);
            }
            XElement v = First(tx, "v");
            return v == null ? "" : v.Value;
        }

        private static string SeriesRef(XElement ser, string name)
        {
            XElement el = First(ser, name);
            if (el == null) return null;
            XElement strRef = First(el, "strRef");
            if (strRef != null) { XElement f = First(strRef, "f"); return f == null ? null : f.Value; }
            XElement numRef = First(el, "numRef");
            if (numRef != null) { XElement f = First(numRef, "f"); return f == null ? null : f.Value; }
            XElement multi = First(el, "multiLvlStrRef");
            if (multi != null) { XElement f = First(multi, "f"); return f == null ? null : f.Value; }
            return null;
        }

        private static Color SeriesColor(XElement ser)
        {
            XElement spPr = First(ser, "spPr");
            if (spPr == null) return Color.Empty;
            XElement clr = First(spPr, "srgbClr");
            if (clr == null) return Color.Empty;
            string val = Attr(clr, "val");
            if (string.IsNullOrEmpty(val) || val.Length != 6) return Color.Empty;
            try
            {
                return Color.FromArgb(
                    int.Parse(val.Substring(0, 2), NumberStyles.HexNumber),
                    int.Parse(val.Substring(2, 2), NumberStyles.HexNumber),
                    int.Parse(val.Substring(4, 2), NumberStyles.HexNumber));
            }
            catch { return Color.Empty; }
        }

        private static string TextOf(XElement e)
        {
            if (e == null) return "";
            List<string> parts = new List<string>();
            foreach (XElement t in Descendants(e, "t")) if (t.Value != null) parts.Add(t.Value);
            return string.Join(" ", parts.ToArray()).Trim();
        }

        private static XElement First(XElement parent, string name)
        {
            if (parent == null) return null;
            foreach (XElement e in parent.Elements())
                if (string.Equals(e.Name.LocalName, name, StringComparison.OrdinalIgnoreCase)) return e;
            return null;
        }

        private static IEnumerable<XElement> Children(XElement parent, string name)
        {
            if (parent == null) yield break;
            foreach (XElement e in parent.Elements())
                if (string.Equals(e.Name.LocalName, name, StringComparison.OrdinalIgnoreCase)) yield return e;
        }

        private static IEnumerable<XElement> Descendants(XElement parent, string name)
        {
            if (parent == null) yield break;
            foreach (XElement e in parent.Descendants())
                if (string.Equals(e.Name.LocalName, name, StringComparison.OrdinalIgnoreCase)) yield return e;
        }

        private static string Attr(XElement e, string name)
        {
            if (e == null) return null;
            XAttribute a = e.Attribute(name);
            return a == null ? null : a.Value;
        }
    }

    public static class ChartWriter
    {
        public static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        public static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        public static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        public static readonly XNamespace Xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

        public static XDocument ChartXml(SheetChart chart)
        {
            XElement chartEl = new XElement(C + "chart");
            if (!string.IsNullOrEmpty(chart.Title))
            {
                chartEl.Add(new XElement(C + "title",
                    new XElement(C + "tx",
                        new XElement(C + "rich",
                            new XElement(A + "bodyPr"),
                            new XElement(A + "p",
                                new XElement(A + "r",
                                    new XElement(A + "t", chart.Title)))))));
            }
            XElement plot = new XElement(C + "plotArea", new XElement(C + "layout"));
            ChartKind kind = chart.Kind == ChartKind.Unknown ? ChartKind.Column : chart.Kind;
            XElement typeEl = new XElement(C + KindName(kind));
            if (kind == ChartKind.Bar || kind == ChartKind.Column)
            {
                typeEl.Add(new XElement(C + "barDir", new XAttribute("val",
                    kind == ChartKind.Bar ? "bar" : "col")));
                typeEl.Add(new XElement(C + "grouping", new XAttribute("val",
                    chart.Stacked ? "stacked" : "clustered")));
                typeEl.Add(new XElement(C + "varyColors", new XAttribute("val", "0")));
            }
            else if (kind == ChartKind.Pie || kind == ChartKind.Doughnut)
            {
                typeEl.Add(new XElement(C + "varyColors", new XAttribute("val", "1")));
            }
            else if (kind == ChartKind.Scatter)
            {
                typeEl.Add(new XElement(C + "scatterStyle", new XAttribute("val", "lineMarker")));
            }

            for (int i = 0; i < chart.Series.Count; i++)
            {
                ChartSeries s = chart.Series[i];
                XElement ser = new XElement(C + "ser",
                    new XElement(C + "idx", new XAttribute("val", i)),
                    new XElement(C + "order", new XAttribute("val", i)));
                if (!string.IsNullOrEmpty(s.NameFormula))
                {
                    ser.Add(new XElement(C + "tx",
                        new XElement(C + "strRef", new XElement(C + "f", s.NameFormula))));
                }
                else if (!string.IsNullOrEmpty(s.Name))
                {
                    ser.Add(new XElement(C + "tx", new XElement(C + "v", s.Name)));
                }
                if (!s.Color.IsEmpty)
                {
                    ser.Add(new XElement(C + "spPr",
                        new XElement(A + "solidFill",
                            new XElement(A + "srgbClr", new XAttribute("val", Hex(s.Color))))));
                }
                if (kind == ChartKind.Scatter)
                {
                    if (!string.IsNullOrEmpty(s.XValuesFormula))
                        ser.Add(new XElement(C + "xVal", NumRef(s.XValuesFormula)));
                    if (!string.IsNullOrEmpty(s.ValuesFormula))
                        ser.Add(new XElement(C + "yVal", NumRef(s.ValuesFormula)));
                }
                else
                {
                    if (!string.IsNullOrEmpty(s.CategoriesFormula))
                        ser.Add(new XElement(C + "cat",
                            new XElement(C + "numRef", new XElement(C + "f", s.CategoriesFormula))));
                    if (!string.IsNullOrEmpty(s.ValuesFormula))
                        ser.Add(new XElement(C + "val", NumRef(s.ValuesFormula)));
                }
                typeEl.Add(ser);
            }
            if (kind == ChartKind.Bar || kind == ChartKind.Column)
                typeEl.Add(new XElement(C + "gapWidth", new XAttribute("val", "150")));
            typeEl.Add(new XElement(C + "axId", new XAttribute("val", "10")));
            typeEl.Add(new XElement(C + "axId", new XAttribute("val", "100")));
            plot.Add(typeEl);
            if (kind != ChartKind.Pie && kind != ChartKind.Doughnut)
            {
                plot.Add(new XElement(C + "catAx",
                    new XElement(C + "axId", new XAttribute("val", "10")),
                    new XElement(C + "scaling", new XElement(C + "orientation", new XAttribute("val", "minMax"))),
                    new XElement(C + "delete", new XAttribute("val", chart.ShowAxes ? "0" : "1")),
                    new XElement(C + "axPos", new XAttribute("val", "b")),
                    new XElement(C + "majorTickMark", new XAttribute("val", "none")),
                    new XElement(C + "minorTickMark", new XAttribute("val", "none")),
                    new XElement(C + "tickLblPos", new XAttribute("val", "nextTo")),
                    new XElement(C + "crossAx", new XAttribute("val", "100")),
                    new XElement(C + "crosses", new XAttribute("val", "autoZero")),
                    new XElement(C + "auto", new XAttribute("val", "1")),
                    new XElement(C + "lblAlgn", new XAttribute("val", "ctr")),
                    new XElement(C + "lblOffset", new XAttribute("val", "100")),
                    new XElement(C + "noMultiLvlLbl", new XAttribute("val", "0"))));
                plot.Add(new XElement(C + "valAx",
                    new XElement(C + "axId", new XAttribute("val", "100")),
                    new XElement(C + "scaling", new XElement(C + "orientation", new XAttribute("val", "minMax"))),
                    new XElement(C + "delete", new XAttribute("val", chart.ShowAxes ? "0" : "1")),
                    new XElement(C + "axPos", new XAttribute("val", "l")),
                    new XElement(C + "majorGridlines"),
                    new XElement(C + "majorTickMark", new XAttribute("val", "none")),
                    new XElement(C + "minorTickMark", new XAttribute("val", "none")),
                    new XElement(C + "tickLblPos", new XAttribute("val", "nextTo")),
                    new XElement(C + "crossAx", new XAttribute("val", "10")),
                    new XElement(C + "crosses", new XAttribute("val", "autoZero")),
                    new XElement(C + "crossBetween", new XAttribute("val", "between")),
                    new XElement(C + "majorUnit", new XAttribute("val", "1")),
                    new XElement(C + "minorUnit", new XAttribute("val", "1"))));
            }
            chartEl.Add(plot);
            if (chart.ShowLegend)
                chartEl.Add(new XElement(C + "legend",
                    new XElement(C + "legendPos", new XAttribute("val", kind == ChartKind.Pie ? "r" : "b")),
                    new XElement(C + "overlay", new XAttribute("val", "0"))));
            chartEl.Add(new XElement(C + "plotVisOnly",
                new XAttribute("val", chart.PlotVisibleOnly ? "1" : "0")));
            chartEl.Add(new XElement(C + "dispBlanksAs", new XAttribute("val", "gap")));
            return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"),
                new XElement(C + "chartSpace",
                    new XAttribute(XNamespace.Xmlns + "c", C.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "a", A.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "r", R.NamespaceName),
                    new XElement(C + "lang", new XAttribute("val", "es-ES")),
                    new XElement(C + "roundedCorners", new XAttribute("val", "0")),
                    chartEl));
        }

        public static XDocument DrawingXml(IList<SheetChart> charts)
        {
            XElement root = new XElement(Xdr + "wsDr",
                new XAttribute(XNamespace.Xmlns + "xdr", Xdr.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "a", A.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "c", C.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "r", R.NamespaceName));
            for (int i = 0; i < charts.Count; i++)
                root.Add(DrawingAnchorXml(charts[i], i + 2, "rId" + (i + 1)));
            return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
        }

        public static XElement DrawingAnchorXml(SheetChart chart, int shapeId, string chartRid)
        {
            long cx = Math.Max(40, chart.WidthPx) * 9525L;
            long cy = Math.Max(30, chart.HeightPx) * 9525L;
            return new XElement(Xdr + "oneCellAnchor",
                new XElement(Xdr + "from",
                    new XElement(Xdr + "col", chart.FromCol),
                    new XElement(Xdr + "colOff", chart.FromColOff),
                    new XElement(Xdr + "row", chart.FromRow),
                    new XElement(Xdr + "rowOff", chart.FromRowOff)),
                new XElement(Xdr + "ext", new XAttribute("cx", cx), new XAttribute("cy", cy)),
                new XElement(Xdr + "graphicFrame",
                    new XElement(Xdr + "nvGraphicFramePr",
                        new XElement(Xdr + "cNvPr", new XAttribute("id", shapeId),
                            new XAttribute("name", "Gráfico " + (shapeId - 1))),
                        new XElement(Xdr + "cNvGraphicFramePr")),
                    new XElement(Xdr + "xfrm",
                        new XElement(A + "off", new XAttribute("x", "0"), new XAttribute("y", "0")),
                        new XElement(A + "ext", new XAttribute("cx", cx), new XAttribute("cy", cy))),
                    new XElement(A + "graphic",
                        new XElement(A + "graphicData", new XAttribute("uri", C.NamespaceName),
                            new XElement(C + "chart", new XAttribute(R + "id", chartRid))))),
                new XElement(Xdr + "clientData"));
        }

        public static XDocument DrawingRels(IList<string> chartParts)
        {
            XElement rels = new XElement(NsPkgRel + "Relationships");
            for (int i = 0; i < chartParts.Count; i++)
                rels.Add(new XElement(NsPkgRel + "Relationship",
                    new XAttribute("Id", "rId" + (i + 1)),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"),
                    new XAttribute("Target", chartParts[i])));
            return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), rels);
        }

        private static readonly XNamespace NsPkgRel =
            "http://schemas.openxmlformats.org/package/2006/relationships";

        private static XElement NumRef(string formula)
        {
            return new XElement(C + "numRef", new XElement(C + "f", formula));
        }

        private static string KindName(ChartKind kind)
        {
            switch (kind)
            {
                case ChartKind.Bar: return "barChart";
                case ChartKind.Line: return "lineChart";
                case ChartKind.Area: return "areaChart";
                case ChartKind.Pie: return "pieChart";
                case ChartKind.Doughnut: return "doughnutChart";
                case ChartKind.Scatter: return "scatterChart";
                case ChartKind.Radar: return "radarChart";
                default: return "barChart";
            }
        }

        private static string Hex(Color c)
        {
            return c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
        }
    }

    public sealed class ChartInsertDialog : Form
    {
        public ChartKind Kind = ChartKind.Column;
        public string ChartTitle = "Gráfico";
        public bool ShowLegend = true;
        public bool ShowAxes = true;
        public bool Stacked;
        private readonly ComboBox _type;
        private readonly TextBox _title;
        private readonly CheckBox _legend;
        private readonly CheckBox _axes;
        private readonly CheckBox _stacked;

        public ChartInsertDialog(string defaultTitle)
            : this(defaultTitle, ChartKind.Column, true, true, false) { }

        public ChartInsertDialog(string defaultTitle, ChartKind kind, bool legend, bool axes,
                                 bool stacked = false)
        {
            Text = "Insertar gráfico";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(390, 205);
            Font = new Font("Segoe UI", 9f);

            Label lt = new Label(); lt.Text = "Tipo:"; lt.SetBounds(14, 18, 70, 20); Controls.Add(lt);
            _type = new ComboBox(); _type.DropDownStyle = ComboBoxStyle.DropDownList;
            _type.Items.Add("Columnas"); _type.Items.Add("Barras"); _type.Items.Add("Líneas");
            _type.Items.Add("Área"); _type.Items.Add("Circular"); _type.Items.Add("Dispersión");
            _type.Items.Add("Anillo"); _type.Items.Add("Radar");
            _type.SelectedIndex = kind == ChartKind.Bar ? 1 :
                kind == ChartKind.Line ? 2 :
                kind == ChartKind.Area ? 3 :
                kind == ChartKind.Pie ? 4 :
                kind == ChartKind.Scatter ? 5 :
                kind == ChartKind.Doughnut ? 6 :
                kind == ChartKind.Radar ? 7 : 0;
            _type.SetBounds(90, 15, 270, 24); Controls.Add(_type);

            Label ln = new Label(); ln.Text = "Título:"; ln.SetBounds(14, 52, 70, 20); Controls.Add(ln);
            _title = new TextBox(); _title.Text = defaultTitle ?? "Gráfico";
            _title.SetBounds(90, 49, 270, 24); Controls.Add(_title);

            _legend = new CheckBox(); _legend.Text = "Mostrar leyenda"; _legend.Checked = legend;
            _legend.SetBounds(90, 82, 150, 22); Controls.Add(_legend);
            _axes = new CheckBox(); _axes.Text = "Mostrar ejes"; _axes.Checked = axes;
            _axes.SetBounds(240, 82, 120, 22); Controls.Add(_axes);
            _stacked = new CheckBox(); _stacked.Text = "Apilado"; _stacked.Checked = stacked;
            _stacked.SetBounds(90, 108, 120, 22); Controls.Add(_stacked);
            _type.SelectedIndexChanged += delegate
            {
                _stacked.Enabled = _type.SelectedIndex == 0 || _type.SelectedIndex == 1;
            };
            _stacked.Enabled = _type.SelectedIndex == 0 || _type.SelectedIndex == 1;

            Button ok = new Button(); ok.Text = "Insertar"; ok.SetBounds(180, 145, 90, 30);
            ok.Click += delegate
            {
                Kind = _type.SelectedIndex == 1 ? ChartKind.Bar :
                       _type.SelectedIndex == 2 ? ChartKind.Line :
                       _type.SelectedIndex == 3 ? ChartKind.Area :
                       _type.SelectedIndex == 4 ? ChartKind.Pie :
                       _type.SelectedIndex == 5 ? ChartKind.Scatter :
                       _type.SelectedIndex == 6 ? ChartKind.Doughnut :
                       _type.SelectedIndex == 7 ? ChartKind.Radar : ChartKind.Column;
                ChartTitle = _title.Text ?? "";
                ShowLegend = _legend.Checked;
                ShowAxes = _axes.Checked;
                Stacked = _stacked.Checked && (Kind == ChartKind.Column || Kind == ChartKind.Bar);
                DialogResult = DialogResult.OK;
                Close();
            };
            Controls.Add(ok);
            Button cancel = new Button(); cancel.Text = "Cancelar"; cancel.SetBounds(278, 145, 90, 30);
            cancel.DialogResult = DialogResult.Cancel; Controls.Add(cancel);
            AcceptButton = ok; CancelButton = cancel;
        }
    }

    public sealed class ChartSeriesDialog : Form
    {
        private readonly List<ChartSeries> _series = new List<ChartSeries>();
        private readonly ListBox _list;
        private readonly TextBox _name;
        private readonly TextBox _values;
        private readonly TextBox _categories;
        private int _index = -1;

        public List<ChartSeries> Result { get { return _series; } }

        public ChartSeriesDialog(SheetChart chart)
        {
            Text = "Series del gráfico";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(520, 300);
            BackColor = Color.White;
            Font = new Font("Segoe UI", 9f);
            foreach (ChartSeries s in chart.Series) _series.Add(s.Clone());

            _list = new ListBox();
            _list.SetBounds(12, 12, 150, 220);
            _list.IntegralHeight = false;
            Controls.Add(_list);
            foreach (ChartSeries s in _series) _list.Items.Add(SeriesLabel(s));
            _list.SelectedIndexChanged += delegate { SaveCurrent(); LoadCurrent(); };

            Label ln = new Label(); ln.Text = "Nombre:"; ln.SetBounds(176, 14, 80, 20); Controls.Add(ln);
            _name = new TextBox(); _name.SetBounds(258, 12, 246, 23); Controls.Add(_name);
            Label lv = new Label(); lv.Text = "Valores:"; lv.SetBounds(176, 48, 80, 20); Controls.Add(lv);
            _values = new TextBox(); _values.SetBounds(258, 46, 246, 23); Controls.Add(_values);
            Label lc = new Label(); lc.Text = "Categorías:"; lc.SetBounds(176, 82, 80, 20); Controls.Add(lc);
            _categories = new TextBox(); _categories.SetBounds(258, 80, 246, 23); Controls.Add(_categories);
            Label help = new Label();
            help.Text = "Use referencias como Hoja1!A2:A10 o =Hoja1!$A$2:$A$10.";
            help.ForeColor = Color.FromArgb(110, 110, 110);
            help.SetBounds(176, 112, 330, 34);
            Controls.Add(help);

            Button add = new Button(); add.Text = "Agregar"; add.SetBounds(176, 152, 78, 28);
            add.Click += delegate
            {
                SaveCurrent();
                ChartSeries s = new ChartSeries();
                s.Name = "Serie " + (_series.Count + 1);
                _series.Add(s);
                _list.Items.Add(SeriesLabel(s));
                _list.SelectedIndex = _series.Count - 1;
            };
            Controls.Add(add);
            Button remove = new Button(); remove.Text = "Quitar"; remove.SetBounds(260, 152, 78, 28);
            remove.Click += delegate
            {
                if (_index < 0 || _index >= _series.Count) return;
                _series.RemoveAt(_index);
                _list.Items.RemoveAt(_index);
                _index = Math.Min(_index, _series.Count - 1);
                if (_index >= 0) _list.SelectedIndex = _index;
            };
            Controls.Add(remove);

            Button ok = new Button(); ok.Text = "Aceptar"; ok.SetBounds(330, 246, 84, 30);
            ok.DialogResult = DialogResult.OK;
            ok.Click += delegate { SaveCurrent(); };
            Controls.Add(ok);
            Button cancel = new Button(); cancel.Text = "Cancelar"; cancel.SetBounds(420, 246, 84, 30);
            cancel.DialogResult = DialogResult.Cancel; Controls.Add(cancel);
            AcceptButton = ok; CancelButton = cancel;
            Shown += delegate { if (_list.Items.Count > 0) _list.SelectedIndex = 0; };
        }

        private static string SeriesLabel(ChartSeries s)
        {
            return string.IsNullOrEmpty(s.Name) ? "(sin nombre)" : s.Name;
        }

        private void SaveCurrent()
        {
            if (_index < 0 || _index >= _series.Count) return;
            ChartSeries s = _series[_index];
            string newName = _name.Text ?? "";
            if (!string.Equals(newName, s.Name, StringComparison.Ordinal)) s.NameFormula = null;
            s.Name = newName;
            s.ValuesFormula = _values.Text ?? "";
            s.CategoriesFormula = _categories.Text ?? "";
            if (_index < _list.Items.Count) _list.Items[_index] = SeriesLabel(s);
        }

        private void LoadCurrent()
        {
            _index = _list.SelectedIndex;
            bool has = _index >= 0 && _index < _series.Count;
            _name.Enabled = _values.Enabled = _categories.Enabled = has;
            if (!has) { _name.Text = _values.Text = _categories.Text = ""; return; }
            ChartSeries s = _series[_index];
            _name.Text = s.Name ?? "";
            _values.Text = s.ValuesFormula ?? "";
            _categories.Text = s.CategoriesFormula ?? "";
        }
    }

    public static class ChartRenderer
    {
        private static readonly Color[] Palette = new Color[]
        {
            Color.FromArgb(68,114,196), Color.FromArgb(237,125,49),
            Color.FromArgb(165,165,165), Color.FromArgb(255,192,0),
            Color.FromArgb(91,155,213), Color.FromArgb(112,173,71),
            Color.FromArgb(255,99,99), Color.FromArgb(128,100,162)
        };

        public static void Draw(Graphics g, RectangleF bounds, SheetChart chart, Workbook wb, Sheet host)
        {
            if (chart == null || bounds.Width < 8 || bounds.Height < 8) return;
            GraphicsState state = g.Save();
            try
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                float scale = Math.Max(0.25f, bounds.Width / Math.Max(1, chart.WidthPx));
                using (SolidBrush bg = new SolidBrush(Color.White)) g.FillRectangle(bg, bounds);
                using (Pen border = new Pen(Color.FromArgb(190, 190, 190), 1f))
                    g.DrawRectangle(border, bounds.X, bounds.Y, bounds.Width - 1, bounds.Height - 1);

                float pad = 6f * scale;
                float titleH = string.IsNullOrEmpty(chart.Title) ? 0 : 18f * scale;
                bool pie = chart.Kind == ChartKind.Pie || chart.Kind == ChartKind.Doughnut;
                float legendH = chart.ShowLegend && (chart.Series.Count > 1 || pie) ? 16f * scale : 0;
                float plotTop = bounds.Y + pad + titleH + (chart.LegendTop ? legendH : 0);
                RectangleF plot = new RectangleF(bounds.X + pad, plotTop,
                    Math.Max(1, bounds.Width - pad * 2),
                    Math.Max(1, bounds.Height - pad * 2 - titleH - legendH));

                using (Font tf = new Font("Segoe UI", Math.Max(7f, 10f * scale), FontStyle.Bold))
                using (SolidBrush tb = new SolidBrush(Color.FromArgb(35, 35, 35)))
                    if (!string.IsNullOrEmpty(chart.Title))
                        g.DrawString(chart.Title, tf, tb,
                            new RectangleF(bounds.X + pad, bounds.Y + 2 * scale, bounds.Width - pad * 2, titleH),
                            new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center, Trimming = StringTrimming.EllipsisCharacter });

                List<double[]> values = new List<double[]>();
                List<double[]> xvalues = new List<double[]>();
                List<string[]> cats = new List<string[]>();
                int maxLen = 0;
                for (int i = 0; i < chart.Series.Count; i++)
                {
                    ChartSeries s = chart.Series[i];
                    double[] vals = ChartData.ReadNumbers(wb, host, s.ValuesFormula,
                        chart.PlotVisibleOnly);
                    if (s.CachedValues.Count > 0 &&
                        (vals.Length == 0 || vals.Length != s.CachedValues.Count))
                        vals = s.CachedValues.ToArray();
                    double[] xs = chart.Kind == ChartKind.Scatter
                        ? ChartData.ReadNumbers(wb, host, s.XValuesFormula, chart.PlotVisibleOnly) : null;
                    if (chart.Kind == ChartKind.Scatter && (xs == null || xs.Length == 0) &&
                        s.CachedXValues.Count > 0) xs = s.CachedXValues.ToArray();
                    string[] cs = ChartData.ReadTexts(wb, host, s.CategoriesFormula,
                        chart.PlotVisibleOnly);
                    if (s.CachedCategories.Count > 0 &&
                        (cs.Length == 0 || cs.Length != s.CachedCategories.Count))
                        cs = s.CachedCategories.ToArray();
                    values.Add(vals);
                    xvalues.Add(xs);
                    cats.Add(cs);
                    maxLen = Math.Max(maxLen, vals.Length);
                }
                for (int i = 0; i < chart.Series.Count; i++)
                {
                    ChartSeries s = chart.Series[i];
                    if (string.IsNullOrEmpty(s.Name))
                        s.Name = ChartData.SeriesName(wb, host, s, i);
                }
                if (maxLen == 0)
                {
                    using (Font f = new Font("Segoe UI", Math.Max(7f, 9f * scale)))
                    using (SolidBrush b = new SolidBrush(Color.Gray))
                        g.DrawString("Sin datos", f, b, plot, new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center });
                    return;
                }

                if (chart.Kind == ChartKind.Pie || chart.Kind == ChartKind.Doughnut)
                {
                    if (chart.OfPie)
                        DrawPieOfPie(g, plot, chart, values, cats.Count > 0 ? cats[0] : null, scale);
                    else
                        DrawPie(g, plot, chart, values, cats.Count > 0 ? cats[0] : null, scale);
                }
                else
                    DrawCartesian(g, plot, chart, values, xvalues, cats, maxLen, scale);

                if (legendH > 0)
                {
                    float legendY = chart.LegendTop ? bounds.Y + pad + titleH : bounds.Bottom - legendH;
                    DrawLegend(g, bounds, chart, scale, legendH, legendY,
                        pie && cats.Count > 0 ? cats[0] : null);
                }
            }
            catch (Exception ex)
            {
                // Un gráfico no debe romper la hoja ni la impresión, pero
                // conservamos el diagnóstico para que no desaparezca en silencio.
                chart.LastError = ex.GetType().Name + ": " + ex.Message;
                try
                {
                    using (Font f = new Font("Segoe UI", 7.5f))
                    using (SolidBrush b = new SolidBrush(Color.FromArgb(170, 30, 30)))
                        g.DrawString("Gráfico no representable", f, b, bounds.Left + 8, bounds.Top + 8);
                }
                catch { }
            }
            finally { g.Restore(state); }
        }

        private static void DrawCartesian(Graphics g, RectangleF plot, SheetChart chart,
                                          List<double[]> series, List<double[]> xseries,
                                          List<string[]> cats, int count, float scale)
        {
            bool bar = chart.Kind == ChartKind.Bar;
            bool scatter = chart.Kind == ChartKind.Scatter;
            float left = 34f * scale;
            float bottom = 22f * scale;
            RectangleF area = new RectangleF(plot.X + left, plot.Y + 4 * scale,
                Math.Max(1, plot.Width - left - 6 * scale),
                Math.Max(1, plot.Height - bottom - 4 * scale));
            if (scatter)
            {
                DrawScatter(g, area, chart, series, xseries, scale);
                return;
            }
            if (chart.Kind == ChartKind.Radar)
            {
                DrawRadar(g, area, chart, series, cats, scale);
                return;
            }
            double min = 0, max = 0, min2 = 0, max2 = 0;
            bool any = false, any2 = false;
            for (int s = 0; s < series.Count; s++)
            {
                bool sec = s < chart.Series.Count && chart.Series[s].Secondary;
                foreach (double v in series[s])
                {
                    if (double.IsNaN(v) || double.IsInfinity(v)) continue;
                    if (sec)
                    {
                        if (!any2) { min2 = max2 = v; any2 = true; }
                        else { if (v < min2) min2 = v; if (v > max2) max2 = v; }
                    }
                    else
                    {
                        if (!any) { min = max = v; any = true; }
                        else { if (v < min) min = v; if (v > max) max = v; }
                    }
                }
            }
            if (!any && any2) { min = min2; max = max2; any = true; }
            if (!any2) { min2 = min; max2 = max; }
            if (!any) return;
            if (chart.Stacked && (chart.Kind == ChartKind.Column || chart.Kind == ChartKind.Bar))
            {
                min = 0; max = 0;
                for (int i = 0; i < count; i++)
                {
                    double pos = 0, neg = 0;
                    for (int s = 0; s < series.Count; s++)
                    {
                        if (chart.Series[s].Secondary) continue;
                        double v = i < series[s].Length ? series[s][i] : 0;
                        if (v >= 0) pos += v; else neg += v;
                    }
                    if (pos > max) max = pos;
                    if (neg < min) min = neg;
                }
            }
            if (min > 0) min = 0;
            if (max < 0) max = 0;
            if (Math.Abs(max - min) < 1e-9) max = min + 1;
            if (min2 > 0) min2 = 0;
            if (max2 < 0) max2 = 0;
            if (Math.Abs(max2 - min2) < 1e-9) max2 = min2 + 1;
            double span = max - min;
            double span2 = max2 - min2;
            if (any2) area.Width = Math.Max(1, area.Width - 30 * scale);

            using (Pen axis = new Pen(Color.FromArgb(120, 120, 120), 1f))
            using (Font lf = new Font("Segoe UI", Math.Max(5.5f, 7f * scale)))
            using (SolidBrush lb = new SolidBrush(Color.FromArgb(80, 80, 80)))
            {
                g.DrawLine(axis, area.Left, area.Top, area.Left, area.Bottom);
                g.DrawLine(axis, area.Left, area.Bottom, area.Right, area.Bottom);
                for (int t = 0; t <= 4; t++)
                {
                    double v = max - span * t / 4.0;
                    float y = area.Top + area.Height * t / 4f;
                    using (Pen grid = new Pen(Color.FromArgb(235, 235, 235), 1f))
                        g.DrawLine(grid, area.Left, y, area.Right, y);
                    string text = NumberFormatter.FormatGeneral(v);
                    float lx = Math.Max(plot.Left + 1, area.Left - 30 * scale);
                    g.DrawString(text, lf, lb, new PointF(lx, y - 7 * scale));
                }

                if (any2)
                {
                    g.DrawLine(axis, area.Right, area.Top, area.Right, area.Bottom);
                    for (int t = 0; t <= 4; t++)
                    {
                        double v = max2 - span2 * t / 4.0;
                        float y = area.Top + area.Height * t / 4f;
                        string text = NumberFormatter.FormatGeneral(v);
                        g.DrawString(text, lf, lb, new PointF(area.Right + 2 * scale, y - 7 * scale));
                    }
                }
                int n = Math.Max(1, count);
                string[] labels = cats.Count > 0 ? cats[0] : null;
                float slot = (bar ? area.Height : area.Width) / n;
                float group = slot * 0.8f;
                float bw = Math.Max(1f, group / Math.Max(1, series.Count));
                bool combo = chart.Kind == ChartKind.Combo;
                bool drawBars = chart.Kind == ChartKind.Column || chart.Kind == ChartKind.Bar || combo;
                for (int i = 0; i < n; i++)
                {
                    float x = area.Left + (bar ? 0 : i * slot);
                    float y = bar ? area.Top + i * slot : area.Bottom;
                    if (drawBars)
                    {
                        if (chart.Stacked)
                        {
                            double positive = 0, negative = 0;
                            for (int s = 0; s < series.Count; s++)
                            {
                                if (combo && SeriesKind(chart, s) == ChartKind.Line) continue;
                                double v = i < series[s].Length ? series[s][i] : 0;
                                float h = (float)(Math.Abs(v) / span * (bar ? area.Width : area.Height));
                                using (SolidBrush b = new SolidBrush(SeriesColor(chart, s)))
                                {
                                    if (bar)
                                    {
                                        float xx = v >= 0 ? area.Left + (float)(positive / span * area.Width)
                                                          : area.Left + (float)(negative / span * area.Width) - h;
                                        float yy = y + (group - bw) / 2;
                                        g.FillRectangle(b, xx, yy, h, Math.Max(1, bw - 1));
                                    }
                                    else
                                    {
                                        float yy;
                                        if (v >= 0)
                                        {
                                            yy = area.Bottom - (float)(positive / span * area.Height) - h;
                                            positive += v;
                                        }
                                        else
                                        {
                                            yy = area.Bottom - (float)(negative / span * area.Height);
                                            negative += v;
                                        }
                                        g.FillRectangle(b, x + (group - bw) / 2, yy,
                                            Math.Max(1, bw - 1), h);
                                    }
                                }
                                if (v >= 0 && bar) positive += v;
                                else if (v < 0 && bar) negative += v;
                            }
                        }
                        else
                        {
                            for (int s = 0; s < series.Count; s++)
                            {
                                if (combo && SeriesKind(chart, s) == ChartKind.Line) continue;
                                double v = i < series[s].Length ? series[s][i] : 0;
                                float h = (float)(Math.Abs(v) / span * (bar ? area.Width : area.Height));
                                if (bar)
                                {
                                    float yy = y + (group - bw * series.Count) / 2 + s * bw;
                                    float xx = v >= 0 ? area.Left : area.Left - h;
                                    using (SolidBrush b = new SolidBrush(SeriesColor(chart, s)))
                                        g.FillRectangle(b, xx, yy, h, Math.Max(1, bw - 1));
                                }
                                else
                                {
                                    float xx = x + (group - bw * series.Count) / 2 + s * bw;
                                    float yy = v >= 0 ? area.Bottom - h : area.Bottom;
                                    using (SolidBrush b = new SolidBrush(SeriesColor(chart, s)))
                                        g.FillRectangle(b, xx, yy, Math.Max(1, bw - 1), h);
                                }
                            }
                        }
                    }
                    if (!bar && labels != null && i < labels.Length && labels[i] != null && n <= 40)
                    {
                        string text = labels[i] ?? "";
                        if (text.Length > 12) text = text.Substring(0, 11) + "…";
                        using (StringFormat sf = new StringFormat { Alignment = StringAlignment.Center })
                            g.DrawString(text, lf, lb, new RectangleF(x, area.Bottom + 2, slot, bottom - 2), sf);
                    }
                }

                if (chart.Kind == ChartKind.Line || chart.Kind == ChartKind.Area || combo)
                {
                    for (int s = 0; s < series.Count; s++)
                    {
                        if (combo && SeriesKind(chart, s) != ChartKind.Line) continue;
                        double[] vals = series[s];
                        if (vals.Length < 2) continue;
                        bool secondarySeries = s < chart.Series.Count && chart.Series[s].Secondary;
                        double lineMin = secondarySeries ? min2 : min;
                        double lineSpan = secondarySeries ? span2 : span;
                        PointF[] pts = new PointF[vals.Length];
                        for (int i = 0; i < vals.Length; i++)
                        {
                            float x = area.Left + (i + 0.5f) * area.Width / vals.Length;
                            float y = area.Bottom - (float)((vals[i] - lineMin) / lineSpan * area.Height);
                            pts[i] = new PointF(x, y);
                        }
                        if (chart.Kind == ChartKind.Area)
                        {
                            PointF[] poly = new PointF[pts.Length + 2];
                            Array.Copy(pts, poly, pts.Length);
                            poly[pts.Length] = new PointF(pts[pts.Length - 1].X, area.Bottom);
                            poly[pts.Length + 1] = new PointF(pts[0].X, area.Bottom);
                            using (SolidBrush b = new SolidBrush(Color.FromArgb(120, SeriesColor(chart, s))))
                                g.FillPolygon(b, poly);
                        }
                        using (Pen p = new Pen(SeriesColor(chart, s), Math.Max(1f, 1.5f * scale)))
                            g.DrawLines(p, pts);
                    }
                }
            }
        }

        private static void DrawRadar(Graphics g, RectangleF area, SheetChart chart,
                                      List<double[]> series, List<string[]> cats, float scale)
        {
            int n = 0;
            double max = 0;
            foreach (double[] values in series)
            {
                n = Math.Max(n, values.Length);
                foreach (double v in values) if (!double.IsNaN(v)) max = Math.Max(max, Math.Abs(v));
            }
            if (n < 3 || max <= 0) return;
            float cx = area.Left + area.Width / 2f;
            float cy = area.Top + area.Height / 2f;
            float radius = Math.Min(area.Width, area.Height) * 0.42f;
            using (Pen grid = new Pen(Color.FromArgb(225, 225, 225), 1f))
            using (SolidBrush text = new SolidBrush(Color.FromArgb(90, 90, 90)))
            using (Font f = new Font("Segoe UI", Math.Max(5.5f, 7f * scale)))
            {
                for (int ring = 1; ring <= 4; ring++)
                {
                    float rr = radius * ring / 4f;
                    PointF[] circle = new PointF[n];
                    for (int i = 0; i < n; i++)
                    {
                        double a = -Math.PI / 2 + i * 2 * Math.PI / n;
                        circle[i] = new PointF(cx + (float)Math.Cos(a) * rr,
                                               cy + (float)Math.Sin(a) * rr);
                    }
                    g.DrawPolygon(grid, circle);
                }
                for (int i = 0; i < n; i++)
                {
                    double a = -Math.PI / 2 + i * 2 * Math.PI / n;
                    g.DrawLine(grid, cx, cy,
                        cx + (float)Math.Cos(a) * radius, cy + (float)Math.Sin(a) * radius);
                    if (cats.Count > 0 && i < cats[0].Length)
                        g.DrawString(cats[0][i], f, text,
                            new PointF(cx + (float)Math.Cos(a) * (radius + 4 * scale) - 8,
                                       cy + (float)Math.Sin(a) * (radius + 4 * scale) - 6));
                }
                for (int s = 0; s < series.Count; s++)
                {
                    PointF[] pts = new PointF[series[s].Length];
                    for (int i = 0; i < series[s].Length; i++)
                    {
                        double a = -Math.PI / 2 + i * 2 * Math.PI / n;
                        float rr = radius * (float)Math.Min(1.0, Math.Abs(series[s][i]) / max);
                        pts[i] = new PointF(cx + (float)Math.Cos(a) * rr,
                                            cy + (float)Math.Sin(a) * rr);
                    }
                    if (pts.Length > 2)
                        using (SolidBrush fill = new SolidBrush(Color.FromArgb(60, SeriesColor(chart, s))))
                            g.FillPolygon(fill, pts);
                    using (Pen line = new Pen(SeriesColor(chart, s), Math.Max(1f, 1.4f * scale)))
                        g.DrawPolygon(line, pts);
                }
            }
        }

        private static void DrawScatter(Graphics g, RectangleF area, SheetChart chart,
                                        List<double[]> yseries, List<double[]> xseries, float scale)
        {
            double xmin = double.MaxValue, xmax = double.MinValue;
            double ymin = double.MaxValue, ymax = double.MinValue;
            for (int s = 0; s < yseries.Count; s++)
            {
                double[] xs = s < xseries.Count ? xseries[s] : null;
                bool indexX = true;
                if (xs != null) foreach (double xv in xs) if (xv != 0) { indexX = false; break; }
                for (int i = 0; i < yseries[s].Length; i++)
                {
                    double x = !indexX && i < xs.Length ? xs[i] : i + 1;
                    double y = yseries[s][i];
                    if (double.IsNaN(x) || double.IsNaN(y)) continue;
                    xmin = Math.Min(xmin, x); xmax = Math.Max(xmax, x);
                    ymin = Math.Min(ymin, y); ymax = Math.Max(ymax, y);
                }
            }
            if (xmin == double.MaxValue) return;
            if (xmax - xmin < 1e-9) xmax = xmin + 1;
            if (ymax - ymin < 1e-9) ymax = ymin + 1;
            using (Pen axis = new Pen(Color.FromArgb(120, 120, 120), 1f))
            {
                g.DrawLine(axis, area.Left, area.Top, area.Left, area.Bottom);
                g.DrawLine(axis, area.Left, area.Bottom, area.Right, area.Bottom);
                for (int s = 0; s < yseries.Count; s++)
                {
                    double[] xs = s < xseries.Count ? xseries[s] : null;
                    bool indexX = true;
                    if (xs != null) foreach (double xv in xs) if (xv != 0) { indexX = false; break; }
                    PointF[] pts = new PointF[yseries[s].Length];
                    for (int i = 0; i < yseries[s].Length; i++)
                    {
                        double x = !indexX && i < xs.Length ? xs[i] : i + 1;
                        pts[i] = new PointF(
                            area.Left + (float)((x - xmin) / (xmax - xmin) * area.Width),
                            area.Bottom - (float)((yseries[s][i] - ymin) / (ymax - ymin) * area.Height));
                    }
                    using (Pen line = new Pen(SeriesColor(chart, s), Math.Max(1f, 1.2f * scale)))
                    {
                        if (pts.Length > 1) g.DrawLines(line, pts);
                        using (SolidBrush b = new SolidBrush(SeriesColor(chart, s)))
                            foreach (PointF p in pts) g.FillEllipse(b, p.X - 2.5f, p.Y - 2.5f, 5f, 5f);
                    }
                }
            }
        }

        private static void DrawPieOfPie(Graphics g, RectangleF plot, SheetChart chart,
                                         List<double[]> series, string[] labels, float scale)
        {
            double[] vals = series.Count > 0 ? series[0] : new double[0];
            double total = 0;
            foreach (double v in vals) if (v > 0) total += v;
            if (total <= 0) return;
            List<int> small = new List<int>();
            for (int i = 0; i < vals.Length; i++)
            {
                if (vals[i] <= 0) continue;
                double pct = vals[i] / total * 100.0;
                bool isSmall;
                if (string.Equals(chart.SplitType, "pos", StringComparison.OrdinalIgnoreCase))
                    isSmall = i + 1 > chart.SplitPos;
                else if (string.Equals(chart.SplitType, "val", StringComparison.OrdinalIgnoreCase))
                    isSmall = vals[i] <= chart.SplitPos;
                else
                    isSmall = pct <= chart.SplitPos;
                if (isSmall) small.Add(i);
            }
            if (small.Count == 0 || small.Count >= vals.Length)
            {
                DrawPie(g, plot, chart, series, labels, scale);
                return;
            }
            List<double> mainValues = new List<double>();
            List<string> mainLabels = new List<string>();
            List<Color> mainColors = new List<Color>();
            double smallTotal = 0;
            for (int i = 0; i < vals.Length; i++)
            {
                if (vals[i] <= 0) continue;
                if (small.Contains(i)) { smallTotal += vals[i]; continue; }
                mainValues.Add(vals[i]);
                mainLabels.Add(labels != null && i < labels.Length ? labels[i] : "");
                mainColors.Add(PieColor(i, chart));
            }
            if (smallTotal > 0)
            {
                mainValues.Add(smallTotal);
                mainLabels.Add("Otros");
                mainColors.Add(Color.FromArgb(190, 190, 190));
            }
            float mainW = plot.Width * 0.54f;
            float mainSize = Math.Min(mainW, plot.Height) * 0.82f;
            RectangleF mainPie = new RectangleF(plot.X + 2, plot.Bottom - mainSize - 2,
                mainSize, mainSize);
            float secondAvail = Math.Max(80, plot.Right - mainPie.Right - 8);
            float secondMax = Math.Min(secondAvail, plot.Height * 0.78f);
            float secondSize = secondMax * (float)Math.Max(0.25, Math.Min(1.5, chart.SecondPieSize / 100.0));
            float secondLeft = mainPie.Right + 8;
            if (secondLeft + secondSize > plot.Right - 2)
                secondLeft = Math.Max(mainPie.Right + 4, plot.Right - secondSize - 2);
            RectangleF secondPie = new RectangleF(secondLeft, plot.Y + 4,
                secondSize, secondSize);
            using (Font f = new Font("Segoe UI", Math.Max(5f, 6.8f * scale)))
            using (SolidBrush lb = new SolidBrush(Color.FromArgb(75, 75, 75)))
            {
                RectangleF mainLabelBounds = new RectangleF(plot.X, plot.Y,
                    Math.Max(1, mainPie.Right - plot.X + 18), plot.Height);
                DrawPieSlices(g, mainPie, mainValues, mainLabels, mainColors, total, f, lb, 6.0f,
                    mainLabelBounds, true);
                List<double> detailValues = new List<double>();
                List<string> detailLabels = new List<string>();
                List<Color> detailColors = new List<Color>();
                foreach (int i in small)
                {
                    detailValues.Add(vals[i]);
                    detailLabels.Add(labels != null && i < labels.Length ? labels[i] : "");
                    detailColors.Add(PieColor(i, chart));
                }
                if (string.Equals(chart.OfPieType, "bar", StringComparison.OrdinalIgnoreCase))
                    DrawDetailBars(g, secondPie, detailValues, detailLabels, detailColors, total, f, lb);
                else
                    DrawPieSlices(g, secondPie, detailValues, detailLabels, detailColors, total, f, lb, 4.0f, plot, true);
                using (Pen line = new Pen(Color.FromArgb(150, 150, 150), Math.Max(0.7f, scale)))
                {
                    g.DrawLine(line, mainPie.Right - 2, mainPie.Top + mainPie.Height * 0.42f,
                        secondPie.Left, secondPie.Top + secondPie.Height * 0.55f);
                }
            }
        }

        private static void DrawDetailBars(Graphics g, RectangleF area, List<double> values,
                                            List<string> labels, List<Color> colors, double total,
                                            Font font, Brush brush)
        {
            double max = 0;
            foreach (double v in values) if (v > max) max = v;
            if (max <= 0) return;
            float rowH = area.Height / Math.Max(1, values.Count);
            for (int i = 0; i < values.Count; i++)
            {
                float y = area.Y + i * rowH;
                float w = Math.Max(2, (float)(values[i] / max * (area.Width * 0.38f)));
                using (SolidBrush b = new SolidBrush(colors[i]))
                    g.FillRectangle(b, area.X + 2, y + 1, w, Math.Max(1, rowH - 2));
                string text = labels != null && i < labels.Count ? labels[i] : "";
                if (text.Length > 15) text = text.Substring(0, 14) + "…";
                text += " " + (values[i] / total * 100).ToString("0",
                    CultureInfo.GetCultureInfo("es-CO")) + "%";
                g.DrawString(text, font, brush,
                    new PointF(area.X + area.Width * 0.42f, y + Math.Max(0, (rowH - font.Height) / 2)));
            }
        }

        private static void DrawPieSlices(Graphics g, RectangleF pie, List<double> values,
                                          List<string> labels, List<Color> colors, double total,
                                          Font font, Brush brush, float minSweepForLabel,
                                          RectangleF labelBounds, bool outside)
        {
            double sum = 0;
            foreach (double v in values) if (v > 0) sum += v;
            if (sum <= 0) return;
            float[] starts = new float[values.Count];
            float[] sweeps = new float[values.Count];
            float cursor = 0;
            for (int i = 0; i < values.Count; i++)
            {
                if (values[i] <= 0) continue;
                float sweep = (float)(values[i] / sum * 360);
                starts[i] = cursor;
                sweeps[i] = sweep;
                using (SolidBrush b = new SolidBrush(colors[i]))
                    g.FillPie(b, pie.X, pie.Y, pie.Width, pie.Height, cursor, sweep);
                cursor += sweep;
            }
            List<RectangleF> usedLabels = new List<RectangleF>();
            for (int i = 0; i < values.Count; i++)
            {
                if (values[i] <= 0 || sweeps[i] < minSweepForLabel ||
                    labels == null || i >= labels.Count) continue;
                double angle = (starts[i] + sweeps[i] / 2) * Math.PI / 180.0;
                float rr = pie.Width * 0.34f;
                float lx = pie.X + pie.Width / 2 + (float)Math.Cos(angle) * rr;
                float ly = pie.Y + pie.Height / 2 + (float)Math.Sin(angle) * rr;
                string text = labels[i] ?? "";
                if (text.Length > 16) text = text.Substring(0, 15) + "…";
                text += " " + (values[i] / total * 100).ToString("0",
                    CultureInfo.GetCultureInfo("es-CO")) + "%";
                SizeF sz = g.MeasureString(text, font);
                if (outside)
                {
                    float edgeX = pie.X + pie.Width / 2 + (float)Math.Cos(angle) * (pie.Width / 2);
                    float edgeY = pie.Y + pie.Height / 2 + (float)Math.Sin(angle) * (pie.Height / 2);
                    bool rightPie = pie.X > labelBounds.X + labelBounds.Width * 0.35f;
                    float tx = pie.X + pie.Width / 2 + (float)Math.Cos(angle) * (pie.Width * 0.72f);
                    float ty = pie.Y + pie.Height / 2 + (float)Math.Sin(angle) * (pie.Height * 0.72f);
                    if (rightPie) tx = pie.Right + 2;
                    else if (Math.Cos(angle) >= 0) tx += 2;
                    else tx -= sz.Width + 2;
                    ty -= sz.Height / 2;
                    tx = Math.Max(labelBounds.X + 1,
                        Math.Min(labelBounds.Right - sz.Width - 1, tx));
                    ty = Math.Max(labelBounds.Y + 1,
                        Math.Min(labelBounds.Bottom - sz.Height - 1, ty));
                    RectangleF labelRect = new RectangleF(tx, ty, sz.Width, sz.Height);
                    for (int attempt = 0; attempt < 12; attempt++)
                    {
                        bool collide = false;
                        foreach (RectangleF used in usedLabels)
                            if (used.IntersectsWith(labelRect)) { collide = true; break; }
                        if (!collide) break;
                        ty = Math.Min(labelBounds.Bottom - sz.Height - 1, ty + sz.Height + 1);
                        labelRect = new RectangleF(tx, ty, sz.Width, sz.Height);
                    }
                    usedLabels.Add(labelRect);
                    using (Pen line = new Pen(Color.FromArgb(170, 170, 170),
                        Math.Max(0.6f, 0.7f)))
                        g.DrawLine(line, edgeX, edgeY,
                            Math.Cos(angle) >= 0 || rightPie ? tx : tx + sz.Width,
                            ty + sz.Height / 2);
                    g.DrawString(text, font, brush, tx, ty);
                }
                else
                {
                    lx -= sz.Width / 2; ly -= sz.Height / 2;
                    lx = Math.Max(pie.X + 1, Math.Min(pie.Right - sz.Width - 1, lx));
                    ly = Math.Max(pie.Y + 1, Math.Min(pie.Bottom - sz.Height - 1, ly));
                    g.DrawString(text, font, brush, lx, ly);
                }
            }
        }

        private static void DrawPie(Graphics g, RectangleF plot, SheetChart chart, List<double[]> series,
                                    string[] labels, float scale)
        {
            double[] vals = series.Count > 0 ? series[0] : new double[0];
            double total = 0;
            foreach (double v in vals) if (v > 0) total += v;
            if (total <= 0) return;
            float size = Math.Min(plot.Width, plot.Height) * 0.72f;
            RectangleF pie = new RectangleF(plot.X + (plot.Width - size) / 2,
                plot.Y + (plot.Height - size) / 2, size, size);
            float start = 0;
            using (Font labelFont = new Font("Segoe UI", Math.Max(5.5f, 7.2f * scale)))
            using (SolidBrush labelBrush = new SolidBrush(Color.FromArgb(75, 75, 75)))
            {
                for (int i = 0; i < vals.Length; i++)
                {
                    if (vals[i] <= 0) continue;
                    float sweep = (float)(vals[i] / total * 360);
                    using (SolidBrush b = new SolidBrush(PieColor(i, chart)))
                        g.FillPie(b, pie.X, pie.Y, pie.Width, pie.Height, start, sweep);
                    if (labels != null && i < labels.Length && sweep >= 12f)
                    {
                        double angle = (start + sweep / 2) * Math.PI / 180.0;
                        float rr = size * 0.43f;
                        float lx = pie.X + pie.Width / 2 + (float)Math.Cos(angle) * rr;
                        float ly = pie.Y + pie.Height / 2 + (float)Math.Sin(angle) * rr;
                        string label = labels[i] ?? "";
                        if (label.Length > 18) label = label.Substring(0, 17) + "…";
                        label += " " + (vals[i] / total * 100).ToString("0", CultureInfo.GetCultureInfo("es-CO")) + "%";
                        SizeF sz = g.MeasureString(label, labelFont);
                        lx -= sz.Width / 2;
                        ly -= sz.Height / 2;
                        lx = Math.Max(plot.X + 2, Math.Min(plot.Right - sz.Width - 2, lx));
                        ly = Math.Max(plot.Y + 2, Math.Min(plot.Bottom - sz.Height - 2, ly));
                        g.DrawString(label, labelFont, labelBrush, lx, ly);
                    }
                    start += sweep;
                }
            }
        }

        private static Color PieColor(int index, SheetChart chart)
        {
            Color[] theme = chart != null && chart.ThemePalette != null
                ? chart.ThemePalette : ChartData.ThemePalette;
            if (theme != null && theme.Length > 0)
            {
                Color baseColor = theme[index % theme.Length];
                int cycle = index / theme.Length;
                if (cycle == 0) return baseColor;
                float factor = cycle % 2 == 0 ? 1.12f : 0.82f;
                return Color.FromArgb(
                    Math.Min(255, (int)(baseColor.R * factor)),
                    Math.Min(255, (int)(baseColor.G * factor)),
                    Math.Min(255, (int)(baseColor.B * factor)));
            }
            return Palette[index % Palette.Length];
        }

        private static void DrawLegend(Graphics g, RectangleF bounds, SheetChart chart, float scale,
                                        float height, float y, string[] pieLabels)
        {
            using (Font f = new Font("Segoe UI", Math.Max(5.5f, 7f * scale)))
            {
                float x = bounds.Left + 8 * scale;
                int count = pieLabels != null ? pieLabels.Length : chart.Series.Count;
                for (int i = 0; i < count; i++)
                {
                    using (SolidBrush b = new SolidBrush(
                        pieLabels != null ? Palette[i % Palette.Length] : SeriesColor(chart, i)))
                        g.FillRectangle(b, x, y + 3 * scale, 8 * scale, 8 * scale);
                    string name = pieLabels != null
                        ? (i < pieLabels.Length ? pieLabels[i] : "")
                        : chart.Series[i].Name;
                    if (string.IsNullOrEmpty(name)) name = "Serie " + (i + 1);
                    SizeF sz = g.MeasureString(name, f);
                    using (SolidBrush tb = new SolidBrush(Color.FromArgb(70, 70, 70)))
                        g.DrawString(name, f, tb, new PointF(x + 11 * scale, y));
                    x += sz.Width + 18 * scale;
                    if (x > bounds.Right - 40 * scale) { x = bounds.Left + 8 * scale; y += 12 * scale; }
                }
            }
        }

        private static ChartKind SeriesKind(SheetChart chart, int index)
        {
            if (index >= 0 && index < chart.Series.Count && chart.Series[index].Kind != ChartKind.Unknown)
                return chart.Series[index].Kind;
            return chart.Kind;
        }

        private static Color SeriesColor(SheetChart chart, int index)
        {
            if (index < chart.Series.Count && !chart.Series[index].Color.IsEmpty)
                return chart.Series[index].Color;
            Color[] theme = chart != null && chart.ThemePalette != null
                ? chart.ThemePalette : ChartData.ThemePalette;
            if (theme != null && theme.Length > 0) return theme[index % theme.Length];
            return Palette[index % Palette.Length];
        }
    }
}
