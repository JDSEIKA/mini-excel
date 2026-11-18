using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Text;

namespace CalcRapido
{
    public static class Limits
    {
        public const int MaxRows = 1048576;
        public const int MaxCols = 16384;
    }

    public static class AddressUtil
    {
        public static string ColumnName(int col)
        {
            if (col < 0) return "#";
            string s = "";
            col++;
            while (col > 0)
            {
                int m = (col - 1) % 26;
                s = (char)('A' + m) + s;
                col = (col - 1) / 26;
            }
            return s;
        }

        public static bool TryColumnIndex(string letters, out int col)
        {
            col = 0;
            if (string.IsNullOrEmpty(letters) || letters.Length > 3) return false;
            for (int i = 0; i < letters.Length; i++)
            {
                char ch = char.ToUpperInvariant(letters[i]);
                if (ch < 'A' || ch > 'Z') return false;
                col = col * 26 + (ch - 'A' + 1);
            }
            col--;
            return col >= 0 && col < Limits.MaxCols;
        }

        public static string CellName(int row, int col)
        {
            return ColumnName(col) + (row + 1).ToString(CultureInfo.InvariantCulture);
        }

        public static string RangeName(int r1, int c1, int r2, int c2)
        {
            if (r1 == r2 && c1 == c2) return CellName(r1, c1);
            return CellName(r1, c1) + ":" + CellName(r2, c2);
        }

        public static bool TryParseCell(string text, out int row, out int col, out bool absRow, out bool absCol)
        {
            row = -1; col = -1; absRow = false; absCol = false;
            if (string.IsNullOrEmpty(text)) return false;
            string s = text.Trim();
            int i = 0;
            if (i < s.Length && s[i] == '$') { absCol = true; i++; }
            int start = i;
            while (i < s.Length && char.IsLetter(s[i])) i++;
            if (i == start || i - start > 3) return false;
            string letters = s.Substring(start, i - start);
            if (i < s.Length && s[i] == '$') { absRow = true; i++; }
            int numStart = i;
            while (i < s.Length && char.IsDigit(s[i])) i++;
            if (i == numStart || i != s.Length) return false;
            long n;
            if (!long.TryParse(s.Substring(numStart), NumberStyles.None, CultureInfo.InvariantCulture, out n)) return false;
            if (n < 1 || n > Limits.MaxRows) return false;
            if (!TryColumnIndex(letters, out col)) return false;
            row = (int)(n - 1);
            return true;
        }

        public static bool TryParseRange(string text, out int r1, out int c1, out int r2, out int c2)
        {
            r1 = c1 = r2 = c2 = 0;
            if (string.IsNullOrEmpty(text)) return false;
            string[] parts = text.Split(':');
            bool ar, ac;
            if (parts.Length == 1)
            {
                if (!TryParseCell(parts[0], out r1, out c1, out ar, out ac)) return false;
                r2 = r1; c2 = c1;
                return true;
            }
            if (parts.Length != 2) return false;
            bool br, bc;
            if (!TryParseCell(parts[0], out r1, out c1, out ar, out ac)) return false;
            if (!TryParseCell(parts[1], out r2, out c2, out br, out bc)) return false;
            if (r1 > r2) { int t = r1; r1 = r2; r2 = t; }
            if (c1 > c2) { int t = c1; c1 = c2; c2 = t; }
            return true;
        }
    }

    public enum ValueKind { Empty, Number, Text, Bool, Error, Range, Array }

    public static class Errors
    {
        public const string Div0 = "#\u00A1DIV/0!";
        public const string Value = "#\u00A1VALOR!";
        public const string Name = "#\u00BFNOMBRE?";
        public const string Ref = "#\u00A1REF!";
        public const string NA = "#N/A";
        public const string Num = "#\u00A1NUM!";
        public const string Circular = "#\u00A1CIRC!";
        public const string Syntax = "#\u00A1SINTAXIS!";
        public const string Error = "#\u00A1ERROR!";
        public const string Spill = "#\u00A1DESBORDAMIENTO!";
    }

    public sealed class RangeRef
    {
        public Sheet Sheet;
        public int Row1;
        public int Col1;
        public int Row2;
        public int Col2;

        public RangeRef(Sheet sheet, int r1, int c1, int r2, int c2)
        {
            Sheet = sheet; Row1 = Math.Min(r1, r2); Col1 = Math.Min(c1, c2);
            Row2 = Math.Max(r1, r2); Col2 = Math.Max(c1, c2);
        }

        public int Rows { get { return Row2 - Row1 + 1; } }
        public int Cols { get { return Col2 - Col1 + 1; } }
    }

    public struct CellValue
    {
        public ValueKind Kind;
        public double Number;
        public string Text;
        public bool Bool;
        public string Error;
        public RangeRef Range;
        public List<CellValue> Values;
        public int ArrayRows;
        public int ArrayCols;

        public static CellValue Empty { get { CellValue v = new CellValue(); v.Kind = ValueKind.Empty; return v; } }
        public static CellValue FromNumber(double d)
        {
            if (double.IsNaN(d) || double.IsInfinity(d)) return FromError(Errors.Num);
            CellValue v = new CellValue(); v.Kind = ValueKind.Number; v.Number = d; return v;
        }
        public static CellValue FromText(string s) { CellValue v = new CellValue(); v.Kind = ValueKind.Text; v.Text = s == null ? "" : s; return v; }
        public static CellValue FromBool(bool b) { CellValue v = new CellValue(); v.Kind = ValueKind.Bool; v.Bool = b; return v; }
        public static CellValue FromError(string e) { CellValue v = new CellValue(); v.Kind = ValueKind.Error; v.Error = e; return v; }
        public static CellValue FromRange(RangeRef r) { CellValue v = new CellValue(); v.Kind = ValueKind.Range; v.Range = r; return v; }
        public static CellValue FromArray(List<CellValue> values)
        {
            List<CellValue> list = values ?? new List<CellValue>();
            return FromArray(list, 1, Math.Max(1, list.Count));
        }
        public static CellValue FromArray(List<CellValue> values, int rows, int cols)
        {
            CellValue v = new CellValue(); v.Kind = ValueKind.Array;
            v.Values = values ?? new List<CellValue>();
            v.ArrayRows = Math.Max(1, rows);
            v.ArrayCols = Math.Max(1, cols);
            return v;
        }

        public bool IsError { get { return Kind == ValueKind.Error; } }
        public bool IsEmpty { get { return Kind == ValueKind.Empty; } }

        public string ToDisplayString()
        {
            switch (Kind)
            {
                case ValueKind.Empty: return "";
                case ValueKind.Text: return Text ?? "";
                case ValueKind.Number: return NumberFormatter.FormatGeneral(Number);
                case ValueKind.Bool: return Bool ? "VERDADERO" : "FALSO";
                case ValueKind.Error: return Error ?? Errors.Error;
                case ValueKind.Array:
                    return Values != null && Values.Count > 0 ? Values[0].ToDisplayString() : "";
                default: return "";
            }
        }

        public string ToFormulaText()
        {
            switch (Kind)
            {
                case ValueKind.Empty: return "";
                case ValueKind.Text: return Text ?? "";
                case ValueKind.Number: return Number.ToString("G15", CultureInfo.InvariantCulture);
                case ValueKind.Bool: return Bool ? "VERDADERO" : "FALSO";
                case ValueKind.Error: return Error ?? Errors.Error;
                default: return "";
            }
        }
    }

    [Flags]
    public enum BorderFlags
    {
        None = 0,
        Left = 1,
        Top = 2,
        Right = 4,
        Bottom = 8,
        All = 15
    }

    public enum NumberFormatKind
    {
        General = 0,
        Number = 1,
        Currency = 2,
        Percent = 3,
        Date = 4,
        Time = 5,
        Text = 6,
        Custom = 7
    }

    public sealed class CellStyle
    {
        public bool Bold;
        public bool Italic;
        public bool Underline;
        public bool WrapText;
        public int Align = -1; // -1 general, 0 izquierda, 1 centro, 2 derecha
        public NumberFormatKind Format = NumberFormatKind.General;
        public int Decimals = 2;
        public Color FontColor = Color.FromArgb(31, 31, 31);
        public Color FillColor = Color.Empty;
        public BorderFlags Borders = BorderFlags.None;
        public float FontSize = 10f;
        public string FontName = "Segoe UI";
        public string CustomFormat;

        public static readonly CellStyle Default = new CellStyle();

        public CellStyle Clone()
        {
            CellStyle s = new CellStyle();
            s.Bold = Bold; s.Italic = Italic; s.Underline = Underline;
            s.WrapText = WrapText;
            s.Align = Align; s.Format = Format; s.Decimals = Decimals;
            s.FontColor = FontColor; s.FillColor = FillColor; s.Borders = Borders;
            s.FontSize = FontSize; s.FontName = FontName; s.CustomFormat = CustomFormat;
            return s;
        }

        public bool IsDefault()
        {
            return !Bold && !Italic && !Underline && !WrapText && Align == -1 &&
                   Format == NumberFormatKind.General &&
                   Decimals == 2 && FontColor.ToArgb() == Default.FontColor.ToArgb() &&
                   FillColor.IsEmpty && Borders == BorderFlags.None &&
                   Math.Abs(FontSize - 10f) < 0.01f && FontName == "Segoe UI" &&
                   string.IsNullOrEmpty(CustomFormat);
        }

        public void SetFormat(NumberFormatKind format)
        {
            Format = format;
            CustomFormat = null;
        }

        public void SetDecimals(int decimals)
        {
            Decimals = Math.Max(0, Math.Min(10, decimals));
            if (Format == NumberFormatKind.Date || Format == NumberFormatKind.Time) return;
            if (!string.IsNullOrEmpty(CustomFormat))
                CustomFormat = RewriteDecimals(CustomFormat, Decimals);
        }

        public void SetAccounting(int decimals)
        {
            Decimals = Math.Max(0, Math.Min(10, decimals));
            Format = NumberFormatKind.Custom;
            string zeros = Decimals > 0 ? "." + new string('0', Decimals) : "";
            CustomFormat = "_(\"$ \"* #,##0" + zeros + "_);_(\"$ \"* \\(#,##0" + zeros +
                            "\\);_(\"$ \"* \\-??" + zeros + "_);_(@_)";
        }

        private static bool IsDateFormatCode(string code)
        {
            if (string.IsNullOrEmpty(code)) return false;
            bool quoted = false;
            int bracket = 0;
            for (int i = 0; i < code.Length; i++)
            {
                char c = code[i];
                if (c == '"') { quoted = !quoted; continue; }
                if (quoted) continue;
                if (c == '\\') { i++; continue; }
                if (c == '[')
                {
                    int end = code.IndexOf(']', i + 1);
                    if (end > i)
                    {
                        string inner = code.Substring(i + 1, end - i - 1).ToLowerInvariant();
                        if (inner.IndexOf('h') >= 0 || inner.IndexOf('m') >= 0 || inner.IndexOf('s') >= 0)
                            return true;
                        i = end;
                        continue;
                    }
                    bracket++;
                    continue;
                }
                if (c == ']' && bracket > 0) { bracket--; continue; }
                if (bracket > 0) continue;
                char l = char.ToLowerInvariant(c);
                if (l == 'y' || l == 'd' || l == 'h' || l == 's') return true;
            }
            return false;
        }

        public static string RewriteDecimals(string code, int decimals)
        {
            if (string.IsNullOrEmpty(code)) return code;
            if (IsDateFormatCode(code)) return code;
            decimals = Math.Max(0, Math.Min(10, decimals));
            string[] sections = SplitFormatSections(code);
            for (int i = 0; i < sections.Length; i++)
                sections[i] = RewriteSectionDecimals(sections[i], decimals);
            return string.Join(";", sections);
        }

        private static string[] SplitFormatSections(string code)
        {
            List<string> parts = new List<string>();
            StringBuilder sb = new StringBuilder();
            bool quoted = false;
            int bracket = 0;
            for (int i = 0; i < code.Length; i++)
            {
                char c = code[i];
                if (c == '"') { quoted = !quoted; sb.Append(c); continue; }
                if (quoted) { sb.Append(c); continue; }
                if (c == '\\' && i + 1 < code.Length) { sb.Append(c); sb.Append(code[i + 1]); i++; continue; }
                if (c == '[') { bracket++; sb.Append(c); continue; }
                if (c == ']' && bracket > 0) { bracket--; sb.Append(c); continue; }
                if (c == ';' && bracket == 0) { parts.Add(sb.ToString()); sb.Length = 0; continue; }
                sb.Append(c);
            }
            parts.Add(sb.ToString());
            return parts.ToArray();
        }

        private static string RewriteSectionDecimals(string section, int decimals)
        {
            if (string.IsNullOrEmpty(section)) return section;
            bool quoted = false;
            int bracket = 0;
            int first = -1, last = -1, dot = -1, exponent = -1;
            for (int i = 0; i < section.Length; i++)
            {
                char c = section[i];
                if (c == '"') { quoted = !quoted; continue; }
                if (quoted) continue;
                if (c == '\\') { i++; continue; }
                if (c == '[') { bracket++; continue; }
                if (c == ']' && bracket > 0) { bracket--; continue; }
                if (bracket > 0) continue;
                if (c == '0' || c == '#' || c == '?')
                {
                    if (first < 0) first = i;
                    if (exponent < 0) last = i;
                }
                else if (c == '.' && dot < 0 && first >= 0) dot = i;
                else if ((c == 'E' || c == 'e') && (i + 1 < section.Length && (section[i + 1] == '+' || section[i + 1] == '-')))
                {
                    exponent = i;
                    break;
                }
            }
            if (first < 0 || last < first) return section;
            int integerEnd = dot >= 0 && dot > first ? dot - 1 : last;
            while (integerEnd > first && (section[integerEnd] == ',' || section[integerEnd] == ' ')) integerEnd--;
            int suffixStart = last + 1;
            if (dot > last) suffixStart = dot + 1;
            if (exponent >= 0) suffixStart = exponent;
            string prefix = section.Substring(0, first);
            string integerPart = section.Substring(first, integerEnd - first + 1);
            string suffix = section.Substring(suffixStart);
            string zeros = decimals > 0 ? "." + new string('0', decimals) : "";
            return prefix + integerPart + zeros + suffix;
        }

        public static CellStyle CopyOf(CellStyle s)
        {
            return s == null ? new CellStyle() : s.Clone();
        }
    }

    public sealed class MergeRange
    {
        public int Row1;
        public int Col1;
        public int Row2;
        public int Col2;

        public MergeRange(int r1, int c1, int r2, int c2)
        {
            Row1 = Math.Min(r1, r2); Col1 = Math.Min(c1, c2);
            Row2 = Math.Max(r1, r2); Col2 = Math.Max(c1, c2);
        }

        public bool Contains(int row, int col)
        {
            return row >= Row1 && row <= Row2 && col >= Col1 && col <= Col2;
        }

        public bool Intersects(int r1, int c1, int r2, int c2)
        {
            return !(r2 < Row1 || r1 > Row2 || c2 < Col1 || c1 > Col2);
        }

        public MergeRange Clone()
        {
            return new MergeRange(Row1, Col1, Row2, Col2);
        }

        public string ToRangeString()
        {
            return AddressUtil.RangeName(Row1, Col1, Row2, Col2);
        }
    }

    public sealed class DefinedName
    {
        public string Name;
        public string Formula;
        public int LocalSheetId = -1;
        public bool Hidden;
        public string Comment;

        public DefinedName Clone()
        {
            DefinedName d = new DefinedName();
            d.Name = Name; d.Formula = Formula; d.LocalSheetId = LocalSheetId;
            d.Hidden = Hidden; d.Comment = Comment;
            return d;
        }
    }

    public sealed class XlsxSourceInfo
    {
        public string Path;
        public long Length;
        public DateTime LastWriteUtc;
        public string WorkbookPart = "xl/workbook.xml";
        public string StylesPart = "xl/styles.xml";
        public string ThemePart = "xl/theme/theme1.xml";
        public string SharedStringsPart = "xl/sharedStrings.xml";
        public bool Date1904;

        public XlsxSourceInfo Clone()
        {
            XlsxSourceInfo s = new XlsxSourceInfo();
            s.Path = Path; s.Length = Length; s.LastWriteUtc = LastWriteUtc;
            s.WorkbookPart = WorkbookPart; s.StylesPart = StylesPart;
            s.ThemePart = ThemePart; s.SharedStringsPart = SharedStringsPart;
            s.Date1904 = Date1904;
            return s;
        }
    }

    public sealed class Cell
    {
        public string Raw;
        public string Formula;
        public CellValue Value;
        public CellValue CachedValue;
        public bool HasCachedValue;
        public CellStyle Style;
        internal FormulaNode Ast;
        internal bool AstReady;
        public int SourceStyleIndex = -1;
        public bool StyleDirty;
        public bool FormulaDirty;
        public bool Dirty;
        public string SourceFormulaText;
        public string SourceFormulaType;
        public string SourceFormulaRef;
        public bool SourceFormulaAca;
        public bool SourceFormulaBx;
        public string Hyperlink;
        public string HyperlinkTooltip;
        public bool HyperlinkInternal;
        public string Comment;
        public long SpillOwner;

        public bool IsFormula { get { return !string.IsNullOrEmpty(Formula); } }

        public CellStyle EffectiveStyle
        {
            get { return Style == null ? CellStyle.Default : Style; }
        }

        public Cell Clone()
        {
            Cell c = new Cell();
            c.Raw = Raw;
            c.Formula = Formula;
            c.Value = Value;
            c.CachedValue = CachedValue;
            c.HasCachedValue = HasCachedValue;
            c.Style = Style == null ? null : Style.Clone();
            c.Ast = null;
            c.AstReady = false;
            c.SourceStyleIndex = SourceStyleIndex;
            c.StyleDirty = StyleDirty;
            c.FormulaDirty = FormulaDirty;
            c.Dirty = Dirty;
            c.SourceFormulaText = SourceFormulaText;
            c.SourceFormulaType = SourceFormulaType;
            c.SourceFormulaRef = SourceFormulaRef;
            c.SourceFormulaAca = SourceFormulaAca;
            c.SourceFormulaBx = SourceFormulaBx;
            c.Hyperlink = Hyperlink;
            c.HyperlinkTooltip = HyperlinkTooltip;
            c.HyperlinkInternal = HyperlinkInternal;
            c.Comment = Comment;
            c.SpillOwner = SpillOwner;
            return c;
        }

        public void SetInput(string text)
        {
            Ast = null;
            AstReady = false;
            SpillOwner = 0;
            Dirty = true;
            CachedValue = CellValue.Empty;
            HasCachedValue = false;
            SourceFormulaType = null;
            SourceFormulaRef = null;
            SourceFormulaAca = false;
            SourceFormulaBx = false;
            if (string.IsNullOrEmpty(text))
            {
                Raw = null;
                Formula = null;
                FormulaDirty = true;
                Value = CellValue.Empty;
                return;
            }
            if (text[0] == '\'')
            {
                Raw = text.Substring(1);
                Formula = null;
                FormulaDirty = false;
                Value = CellValue.FromText(Raw);
                return;
            }
            if (text[0] == '=')
            {
                Raw = text;
                Formula = text;
                FormulaDirty = true;
                Value = CellValue.Empty;
                string fupper = text.ToUpperInvariant();
                if (fupper.StartsWith("=HOY(", StringComparison.Ordinal) ||
                    fupper.StartsWith("=TODAY(", StringComparison.Ordinal) ||
                    fupper.StartsWith("=FECHA(", StringComparison.Ordinal) ||
                    fupper.StartsWith("=DATE(", StringComparison.Ordinal) ||
                    fupper.StartsWith("=AHORA(", StringComparison.Ordinal) ||
                    fupper.StartsWith("=NOW(", StringComparison.Ordinal))
                {
                    if (Style == null) Style = new CellStyle();
                    if (Style.Format == NumberFormatKind.General)
                    {
                        Style.Format = NumberFormatKind.Date;
                        StyleDirty = true;
                    }
                }
                return;
            }
            Raw = text;
            Formula = null;
            FormulaDirty = false;
            string upper = text.Trim().ToUpperInvariant();
            if (upper == "VERDADERO" || upper == "TRUE") { Value = CellValue.FromBool(true); return; }
            if (upper == "FALSO" || upper == "FALSE") { Value = CellValue.FromBool(false); return; }
            double n;
            if (NumberParser.TryParse(text, out n))
            {
                Value = CellValue.FromNumber(n);
                string trimmed = text.Trim();
                if (trimmed.EndsWith("%", StringComparison.Ordinal))
                {
                    if (Style == null) Style = new CellStyle();
                    if (Style.Format == NumberFormatKind.General)
                    {
                        Style.Format = NumberFormatKind.Percent;
                        Style.Decimals = PercentDecimals(trimmed);
                        StyleDirty = true;
                    }
                }
            }
            else Value = CellValue.FromText(text);
        }

        public void SetTextLiteral(string text)
        {
            Ast = null;
            AstReady = false;
            Dirty = true;
            Formula = null;
            FormulaDirty = false;
            CachedValue = CellValue.Empty;
            HasCachedValue = false;
            SourceFormulaType = null;
            SourceFormulaRef = null;
            SourceFormulaAca = false;
            SourceFormulaBx = false;
            if (string.IsNullOrEmpty(text))
            {
                Raw = null;
                Value = CellValue.Empty;
                return;
            }
            Raw = text;
            Value = CellValue.FromText(text);
        }

        public void SetNumber(double value)
        {
            Ast = null;
            AstReady = false;
            Dirty = true;
            Formula = null;
            FormulaDirty = false;
            CachedValue = CellValue.Empty;
            HasCachedValue = false;
            SourceFormulaType = null;
            SourceFormulaRef = null;
            SourceFormulaAca = false;
            SourceFormulaBx = false;
            Value = CellValue.FromNumber(value);
            Raw = NumberFormatter.FormatGeneral(value);
        }

        public void SetError(string error)
        {
            Ast = null;
            AstReady = false;
            Dirty = true;
            Formula = null;
            FormulaDirty = false;
            CachedValue = CellValue.Empty;
            HasCachedValue = false;
            SourceFormulaType = null;
            SourceFormulaRef = null;
            SourceFormulaAca = false;
            SourceFormulaBx = false;
            Raw = string.IsNullOrEmpty(error) ? Errors.Error : error;
            Value = CellValue.FromError(Raw);
        }

        private static int PercentDecimals(string text)
        {
            int dot = text.LastIndexOf('.');
            int comma = text.LastIndexOf(',');
            int sep = Math.Max(dot, comma);
            if (sep < 0) return 0;
            int n = 0;
            for (int i = sep + 1; i < text.Length && char.IsDigit(text[i]); i++) n++;
            return Math.Max(0, Math.Min(10, n));
        }

        public void ClearContents()
        {
            Raw = null;
            Formula = null;
            Value = CellValue.Empty;
            CachedValue = CellValue.Empty;
            HasCachedValue = false;
            Ast = null;
            AstReady = false;
            Dirty = true;
            FormulaDirty = true;
            SourceFormulaType = null;
            SourceFormulaRef = null;
            SourceFormulaAca = false;
            SourceFormulaBx = false;
            Hyperlink = null;
            HyperlinkTooltip = null;
            HyperlinkInternal = false;
        }
    }

    public static class NumberParser
    {
        public static bool TryParse(string text, out double value)
        {
            value = 0;
            if (string.IsNullOrEmpty(text)) return false;
            string s = text.Trim();
            if (s.Length == 0) return false;
            bool negative = false;
            if (s.Length > 2 && s[0] == '(' && s[s.Length - 1] == ')')
            {
                negative = true;
                s = s.Substring(1, s.Length - 2).Trim();
            }
            bool percent = false;
            if (s.EndsWith("%", StringComparison.Ordinal))
            {
                percent = true;
                s = s.Substring(0, s.Length - 1).Trim();
            }
            s = s.Replace("$", "").Replace(" ", "").Replace("\u00A0", "").Replace("USD", "");
            if (s.Length == 0) return false;
            if (s[0] == '+') s = s.Substring(1);
            if (s.Length == 0) return false;

            int lastDot = s.LastIndexOf('.');
            int lastComma = s.LastIndexOf(',');
            bool hasDot = lastDot >= 0;
            bool hasComma = lastComma >= 0;
            string normalized;
            if (hasDot && hasComma)
            {
                if (lastComma > lastDot)
                {
                    normalized = s.Replace(".", "").Replace(',', '.');
                }
                else
                {
                    normalized = s.Replace(",", "");
                }
            }
            else if (hasComma)
            {
                normalized = s.Replace(',', '.');
            }
            else if (hasDot)
            {
                int dotCount = 0;
                for (int i = 0; i < s.Length; i++) if (s[i] == '.') dotCount++;
                if (dotCount > 1)
                {
                    normalized = s.Replace(".", "");
                }
                else
                {
                    int digitsAfter = s.Length - lastDot - 1;
                    int digitsBefore = lastDot;
                    bool looksGrouped = digitsAfter == 3 && digitsBefore >= 1 && digitsBefore <= 3 &&
                                        s[0] != '0';
                    normalized = looksGrouped ? s.Replace(".", "") : s;
                }
            }
            else
            {
                normalized = s;
            }
            if (!double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out value)) return false;
            if (negative) value = -value;
            if (percent) value /= 100.0;
            return !double.IsNaN(value) && !double.IsInfinity(value);
        }
    }

    public static class NumberFormatter
    {
        public static readonly CultureInfo Es = CultureInfo.GetCultureInfo("es-CO");

        public static string FormatGeneral(double d)
        {
            if (double.IsNaN(d) || double.IsInfinity(d)) return Errors.Num;
            double a = Math.Abs(d);
            if (a != 0 && (a >= 1e12 || a < 1e-9))
                return d.ToString("0.#####E+00", Es).Replace("E+", "E+").Replace("E-", "E-");
            string s = d.ToString("0.##########", Es);
            if (s == "-0") s = "0";
            return s;
        }

        public static string Format(CellValue v, CellStyle style)
        {
            if (v.Kind == ValueKind.Error) return v.Error;
            if (v.Kind == ValueKind.Empty) return "";
            if (v.Kind == ValueKind.Text) return v.Text;
            if (v.Kind == ValueKind.Bool) return v.Bool ? "VERDADERO" : "FALSO";
            if (v.Kind != ValueKind.Number) return "";
            if (style == null) style = CellStyle.Default;
            switch (style.Format)
            {
                case NumberFormatKind.Number:
                    return FormatNumber(v.Number, style.Decimals, true);
                case NumberFormatKind.Currency:
                    {
                        if (!string.IsNullOrEmpty(style.CustomFormat))
                            return ExcelNumberFormat.Format(v.Number, style.CustomFormat);
                        string zeros = new string('0', Math.Max(0, style.Decimals));
                        string body = "#,##0" + (style.Decimals > 0 ? "." + zeros : "");
                        string fmt = "$ " + body + ";-$ " + body;
                        return SafeFormat(v.Number, fmt);
                    }
                case NumberFormatKind.Percent:
                    {
                        string zeros = new string('0', Math.Max(0, style.Decimals));
                        string fmt = "0" + (style.Decimals > 0 ? "." + zeros : "") + "%";
                        return SafeFormat(v.Number, fmt);
                    }
                case NumberFormatKind.Date:
                    return string.IsNullOrEmpty(style.CustomFormat)
                        ? FormatDate(v.Number, false)
                        : ExcelNumberFormat.Format(v.Number, style.CustomFormat);
                case NumberFormatKind.Time:
                    return string.IsNullOrEmpty(style.CustomFormat)
                        ? FormatDate(v.Number, true)
                        : ExcelNumberFormat.Format(v.Number, style.CustomFormat);
                case NumberFormatKind.Text:
                    return FormatGeneral(v.Number);
                case NumberFormatKind.Custom:
                    return ExcelNumberFormat.Format(v.Number, style.CustomFormat);
                default:
                    return FormatGeneral(v.Number);
            }
        }

        public static string FormatNumber(double d, int decimals, bool grouping)
        {
            decimals = Math.Max(0, Math.Min(10, decimals));
            string zeros = decimals > 0 ? "." + new string('0', decimals) : "";
            string fmt = (grouping ? "#,##0" : "0") + zeros;
            return SafeFormat(d, fmt);
        }

        public static string SafeFormat(double d, string fmt)
        {
            try
            {
                string s = d.ToString(fmt, Es);
                if (s == "-0") s = "0";
                return s;
            }
            catch
            {
                return FormatGeneral(d);
            }
        }

        public static string FormatDate(double serial, bool timeOnly)
        {
            if (serial < 0 || serial > 2958465) return Errors.Num;
            try
            {
                DateTime dt = DateTime.FromOADate(serial);
                return timeOnly ? dt.ToString("HH:mm:ss", Es) : dt.ToString("dd/MM/yyyy", Es);
            }
            catch
            {
                return Errors.Num;
            }
        }

        public static bool IsRightAligned(CellValue v, CellStyle style)
        {
            if (style != null && style.Align >= 0) return style.Align == 2;
            return v.Kind == ValueKind.Number || v.Kind == ValueKind.Error;
        }

        public static int EffectiveAlign(CellValue v, CellStyle style)
        {
            if (style != null && style.Align >= 0) return style.Align;
            if (v.Kind == ValueKind.Number || v.Kind == ValueKind.Error) return 2;
            if (v.Kind == ValueKind.Bool) return 1;
            return 0;
        }
    }

    public static class ExcelNumberFormat
    {
        private static readonly CultureInfo Es = CultureInfo.GetCultureInfo("es-CO");

        public static string Format(double value, string code)
        {
            if (string.IsNullOrEmpty(code)) return NumberFormatter.FormatGeneral(value);
            if (code.Trim().Equals("General", StringComparison.OrdinalIgnoreCase))
                return NumberFormatter.FormatGeneral(value);
            string[] sections = SplitSections(code);
            int pick;
            double work = value;
            if (value > 0) pick = 0;
            else if (value < 0) { pick = sections.Length > 1 ? 1 : 0; work = Math.Abs(value); }
            else pick = sections.Length > 2 ? 2 : (sections.Length > 1 ? 0 : 0);
            if (pick >= sections.Length) pick = 0;
            string rawSection = sections[pick];
            string section = CleanSection(rawSection);
            if (section.Length == 0) return "";
            if (IsDateOrTime(section)) return FormatDate(work, rawSection);
            return FormatNumber(work, section, value < 0 && sections.Length == 1);
        }

        private static string[] SplitSections(string code)
        {
            List<string> parts = new List<string>();
            StringBuilder sb = new StringBuilder();
            bool quoted = false;
            int bracket = 0;
            for (int i = 0; i < code.Length; i++)
            {
                char c = code[i];
                if (c == '"')
                {
                    quoted = !quoted;
                    sb.Append(c);
                }
                else if (!quoted && c == '[') { bracket++; sb.Append(c); }
                else if (!quoted && c == ']' && bracket > 0) { bracket--; sb.Append(c); }
                else if (!quoted && bracket == 0 && c == ';')
                {
                    parts.Add(sb.ToString());
                    sb.Length = 0;
                }
                else sb.Append(c);
            }
            parts.Add(sb.ToString());
            return parts.ToArray();
        }

        private static string CleanSection(string s)
        {
            StringBuilder sb = new StringBuilder();
            bool quoted = false;
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (c == '"') { quoted = !quoted; continue; }
                if (quoted) { sb.Append(c); continue; }
                if (c == '[')
                {
                    int end = s.IndexOf(']', i + 1);
                    if (end > i)
                    {
                        string inner = s.Substring(i + 1, end - i - 1);
                        if (inner.StartsWith("$", StringComparison.Ordinal))
                        {
                            int dash = inner.IndexOf('-');
                            if (dash > 1) sb.Append(inner.Substring(1, dash - 1));
                        }
                        i = end;
                        continue;
                    }
                }
                if (c == '\\' && i + 1 < s.Length) { sb.Append(s[i + 1]); i++; continue; }
                if (c == '_' && i + 1 < s.Length) { sb.Append(' '); i++; continue; }
                if (c == '*' && i + 1 < s.Length) { i++; continue; }
                sb.Append(c);
            }
            return sb.ToString();
        }

        private static bool IsDateOrTime(string s)
        {
            string lower = s.ToLowerInvariant();
            if (lower.IndexOf("general", StringComparison.Ordinal) >= 0) return false;
            bool hasDate = lower.IndexOf('y') >= 0 || lower.IndexOf('d') >= 0 ||
                           lower.IndexOf("mmm", StringComparison.Ordinal) >= 0 ||
                           lower.IndexOf("m/", StringComparison.Ordinal) >= 0 ||
                           lower.IndexOf("/m", StringComparison.Ordinal) >= 0;
            bool hasTime = lower.IndexOf('h') >= 0 || lower.IndexOf('s') >= 0;
            return hasDate || hasTime;
        }

        private static string FormatDate(double serial, string code)
        {
            DateTime dt;
            try
            {
                if (serial < 0 || serial > 2958465) return NumberFormatter.FormatGeneral(serial);
                dt = DateTime.FromOADate(serial);
            }
            catch { return NumberFormatter.FormatGeneral(serial); }
            string lower = code.ToLowerInvariant();
            bool hasDate = lower.IndexOf('y') >= 0 || lower.IndexOf('d') >= 0 ||
                           lower.IndexOf("mmm", StringComparison.Ordinal) >= 0 ||
                           lower.IndexOf("m/", StringComparison.Ordinal) >= 0 ||
                           lower.IndexOf("/m", StringComparison.Ordinal) >= 0;
            bool hasTime = lower.IndexOf('h') >= 0 || lower.IndexOf('s') >= 0;
            if (hasDate && hasTime)
            {
                int split = code.IndexOf(' ');
                if (split > 0)
                {
                    string datePattern = DatePattern(code.Substring(0, split), false);
                    string timePattern = DatePattern(code.Substring(split + 1), true);
                    try { return dt.ToString(datePattern, Es) + " " + dt.ToString(timePattern, Es); }
                    catch { }
                }
            }
            string pattern = DatePattern(code, !hasDate && hasTime);
            try { return dt.ToString(pattern, Es); }
            catch { return dt.ToString("dd/MM/yyyy", Es); }
        }

        private static string DatePattern(string code, bool timeOnly)
        {
            StringBuilder sb = new StringBuilder();
            int i = 0;
            while (i < code.Length)
            {
                char c = code[i];
                if (c == '"')
                {
                    int end = code.IndexOf('"', i + 1);
                    if (end < 0) end = code.Length;
                    string literal = code.Substring(i + 1, Math.Max(0, end - i - 1));
                    AppendLiteral(sb, literal);
                    i = end < code.Length ? end + 1 : end;
                    continue;
                }
                if (c == '\\' && i + 1 < code.Length)
                {
                    AppendLiteral(sb, code[i + 1].ToString());
                    i += 2;
                    continue;
                }
                if (c == '_' || c == '*')
                {
                    i += 2;
                    continue;
                }
                if (c == '[')
                {
                    int end = code.IndexOf(']', i + 1);
                    if (end > i)
                    {
                        string inner = code.Substring(i + 1, end - i - 1).ToLowerInvariant();
                        if (inner == "h") sb.Append('H');
                        else if (inner == "m") sb.Append('m');
                        else if (inner == "s") sb.Append('s');
                        i = end + 1;
                        continue;
                    }
                }
                string rest = code.Substring(i);
                if (StartsWith(rest, "AM/PM") || StartsWith(rest, "am/pm"))
                {
                    sb.Append("tt"); i += 5; continue;
                }
                if (StartsWith(rest, "yyyy")) { sb.Append("yyyy"); i += 4; continue; }
                if (StartsWith(rest, "yy")) { sb.Append("yy"); i += 2; continue; }
                if (StartsWith(rest, "mmmm")) { sb.Append("MMMM"); i += 4; continue; }
                if (StartsWith(rest, "mmm")) { sb.Append("MMM"); i += 3; continue; }
                if (StartsWith(rest, "mm")) { sb.Append("MM"); i += 2; continue; }
                if (StartsWith(rest, "dddd")) { sb.Append("dddd"); i += 4; continue; }
                if (StartsWith(rest, "ddd")) { sb.Append("ddd"); i += 3; continue; }
                if (StartsWith(rest, "dd")) { sb.Append("dd"); i += 2; continue; }
                if (StartsWith(rest, "hh")) { sb.Append("HH"); i += 2; continue; }
                if (StartsWith(rest, "ss")) { sb.Append("ss"); i += 2; continue; }
                if (c == 'y') { sb.Append("%y"); i++; continue; }
                if (c == 'm') { sb.Append("%M"); i++; continue; }
                if (c == 'd') { sb.Append("%d"); i++; continue; }
                if (c == 'h') { sb.Append("%H"); i++; continue; }
                if (c == 's') { sb.Append("%s"); i++; continue; }
                AppendLiteral(sb, c.ToString());
                i++;
            }
            return sb.ToString();
        }

        private static bool StartsWith(string s, string value)
        {
            return s.StartsWith(value, StringComparison.OrdinalIgnoreCase);
        }

        private static void AppendLiteral(StringBuilder sb, string literal)
        {
            foreach (char c in literal)
            {
                if (char.IsLetter(c)) sb.Append('\\');
                sb.Append(c);
            }
        }

        private static string FormatNumber(double value, string section, bool addSign)
        {
            string pattern = section;
            if (pattern.IndexOf('@') >= 0) return NumberFormatter.FormatGeneral(value);
            if (pattern.IndexOf("General", StringComparison.OrdinalIgnoreCase) >= 0)
                return NumberFormatter.FormatGeneral(value);
            pattern = pattern.Replace("\\", "");
            pattern = pattern.Replace("?", "#");
            if (pattern.Length == 0) return NumberFormatter.FormatGeneral(value);
            bool hasPlaceholder = false;
            foreach (char c in pattern) if (c == '0' || c == '#') { hasPlaceholder = true; break; }
            if (!hasPlaceholder)
            {
                if (pattern.IndexOf('@') >= 0) return NumberFormatter.FormatGeneral(value);
                return pattern;
            }
            try
            {
                string result = value.ToString(pattern, Es);
                if (addSign && value < 0 && result.Length > 0 && result[0] != '-') result = "-" + result;
                return result;
            }
            catch { return NumberFormatter.FormatGeneral(value); }
        }
    }

    public sealed class DataValidationRule
    {
        public int Row1;
        public int Col1;
        public int Row2;
        public int Col2;
        public readonly List<string> Items = new List<string>();
        public string SourceRange;
        public string Message = "El valor no está en la lista.";
        public bool AllowBlank = true;

        public bool Contains(int row, int col)
        {
            return row >= Math.Min(Row1, Row2) && row <= Math.Max(Row1, Row2) &&
                   col >= Math.Min(Col1, Col2) && col <= Math.Max(Col1, Col2);
        }

        public DataValidationRule Clone()
        {
            DataValidationRule r = new DataValidationRule
            {
                Row1 = Row1, Col1 = Col1, Row2 = Row2, Col2 = Col2,
                SourceRange = SourceRange, Message = Message, AllowBlank = AllowBlank
            };
            r.Items.AddRange(Items);
            return r;
        }
    }

    public sealed class FilterSelection
    {
        public int Column;
        public readonly List<string> Values = new List<string>();
        public FilterSelection Clone()
        {
            FilterSelection f = new FilterSelection { Column = Column };
            f.Values.AddRange(Values);
            return f;
        }
    }

    public sealed class Sheet
    {
        public Workbook Workbook;
        public string Name;
        public readonly Dictionary<long, Cell> Cells = new Dictionary<long, Cell>();
        public readonly Dictionary<int, float> ColWidths = new Dictionary<int, float>();
        public readonly Dictionary<int, float> RowHeights = new Dictionary<int, float>();
        public readonly HashSet<int> HiddenRows = new HashSet<int>();
        public readonly HashSet<int> HiddenCols = new HashSet<int>();
        public readonly List<MergeRange> Merges = new List<MergeRange>();
        public readonly List<SheetChart> Charts = new List<SheetChart>();
        public readonly List<SheetImage> Images = new List<SheetImage>();
        public readonly List<SheetShape> Shapes = new List<SheetShape>();
        public readonly List<PivotRange> PivotRanges = new List<PivotRange>();
        public readonly List<SheetSlicer> Slicers = new List<SheetSlicer>();
        public bool HasSlicerPlaceholders;
        public bool HasPicturePlaceholders;
        public bool HasChartPlaceholders;
        public readonly List<ConditionalFormat> ConditionalFormats = new List<ConditionalFormat>();
        public readonly List<TableFormat> TableFormats = new List<TableFormat>();
        public int DataVersion;
        public bool ShowGridLines = true;
        public bool PrintGridLines;
        public bool PrintHeadings;
        public bool PrintLandscape;
        public int PrintPaperSize;
        public int PrintScalePercent = 100;
        public bool PrintFitToWidth;
        public bool PrintFitToHeight;
        public float PrintMarginMm = -1f;
        public string PrintArea;
        public string PrintHeader;
        public string PrintFooter;
        public string PrintTitleRows;
        public string PrintTitleCols;
        public bool IsProtected;
        public string ProtectionPasswordHash;

        public bool VerifyProtectionPassword(string password)
        {
            if (string.IsNullOrEmpty(ProtectionPasswordHash)) return true;
            return string.Equals(ProtectionPasswordHash, ExcelPasswordHash(password ?? ""),
                StringComparison.OrdinalIgnoreCase);
        }

        public static string ExcelPasswordHash(string password)
        {
            int hash = 0;
            for (int i = 0; i < password.Length; i++)
            {
                int value = password[i] << (i + 1);
                int rotated = value >> 15;
                value &= 0x7FFF;
                hash ^= (rotated | value);
            }
            hash ^= password.Length;
            hash ^= 0xCE4B;
            return hash.ToString("X4", CultureInfo.InvariantCulture);
        }
        public readonly List<int> PrintRowBreaks = new List<int>();
        public readonly List<SheetImage> DeletedImages = new List<SheetImage>();
        public bool ImagesDirty;
        public readonly List<DataValidationRule> Validations = new List<DataValidationRule>();
        private readonly Dictionary<long, MergeRange> _mergeLookup = new Dictionary<long, MergeRange>();
        public int FreezeRows;
        public int FreezeCols;
        public string AutoFilterRef;
        public readonly List<FilterSelection> FilterSelections = new List<FilterSelection>();
        public string SourcePart;
        public bool WidthsDirty;
        public bool RowsDirty;
        public bool MergesDirty;
        public bool HyperlinksDirty;
        public bool StructureDirty;
        public bool ChartsDirty;
        public int UsedMaxRow = -1;
        public int UsedMaxCol = -1;

        public Sheet(Workbook wb, string name)
        {
            Workbook = wb;
            Name = name;
        }

        public static long Key(int row, int col)
        {
            return ((long)row << 14) | (uint)col;
        }

        public Cell GetCell(int row, int col)
        {
            if (row < 0 || col < 0 || row >= Limits.MaxRows || col >= Limits.MaxCols) return null;
            Cell c;
            Cells.TryGetValue(Key(row, col), out c);
            return c;
        }

        public Cell GetOrCreate(int row, int col)
        {
            if (row < 0 || col < 0 || row >= Limits.MaxRows || col >= Limits.MaxCols)
                throw new ArgumentOutOfRangeException("row/col", "Referencia de celda fuera de rango.");
            long key = Key(row, col);
            Cell c;
            if (!Cells.TryGetValue(key, out c))
            {
                c = new Cell();
                c.Value = CellValue.Empty;
                Cells[key] = c;
            }
            if (row > UsedMaxRow) UsedMaxRow = row;
            if (col > UsedMaxCol) UsedMaxCol = col;
            return c;
        }

        public void Touch()
        {
            DataVersion++;
        }

        public void SetInput(int row, int col, string text)
        {
            DataVersion++;
            TableFormatting.ExpandForCell(this, row, col);
            Cell c = GetOrCreate(row, col);
            c.SetInput(text);
            if ((c.Raw == null) && c.Style == null)
            {
                Cells.Remove(Key(row, col));
            }
            if (row > UsedMaxRow) UsedMaxRow = row;
            if (col > UsedMaxCol) UsedMaxCol = col;
        }

        public void SetTextLiteral(int row, int col, string text)
        {
            DataVersion++;
            TableFormatting.ExpandForCell(this, row, col);
            Cell c = GetOrCreate(row, col);
            c.SetTextLiteral(text);
            if (c.Raw == null && c.Style == null) Cells.Remove(Key(row, col));
            if (row > UsedMaxRow) UsedMaxRow = row;
            if (col > UsedMaxCol) UsedMaxCol = col;
        }

        public void SetNumber(int row, int col, double value)
        {
            DataVersion++;
            TableFormatting.ExpandForCell(this, row, col);
            Cell c = GetOrCreate(row, col);
            c.SetNumber(value);
            if (row > UsedMaxRow) UsedMaxRow = row;
            if (col > UsedMaxCol) UsedMaxCol = col;
        }

        public void SetError(int row, int col, string error)
        {
            DataVersion++;
            TableFormatting.ExpandForCell(this, row, col);
            Cell c = GetOrCreate(row, col);
            c.SetError(error);
            if (row > UsedMaxRow) UsedMaxRow = row;
            if (col > UsedMaxCol) UsedMaxCol = col;
        }

        public void SetStyle(int row, int col, CellStyle style)
        {
            Cell c = GetOrCreate(row, col);
            if (style == null || style.IsDefault()) c.Style = null;
            else c.Style = style;
            c.StyleDirty = true;
            c.Dirty = true;
            if ((c.Raw == null) && c.Style == null)
                Cells.Remove(Key(row, col));
        }

        public void ClearContents(int row, int col)
        {
            DataVersion++;
            Cell c = GetCell(row, col);
            if (c == null) return;
            c.ClearContents();
            if (c.Style == null) Cells.Remove(Key(row, col));
        }

        public void ClearRange(int r1, int c1, int r2, int c2)
        {
            if (r1 < 0) r1 = 0;
            if (c1 < 0) c1 = 0;
            if (r2 >= Limits.MaxRows) r2 = Limits.MaxRows - 1;
            if (c2 >= Limits.MaxCols) c2 = Limits.MaxCols - 1;
            if (r1 > r2 || c1 > c2) return;
            long area = (long)(r2 - r1 + 1) * (c2 - c1 + 1);
            if (area > 200000)
            {
                List<long> keys = new List<long>();
                foreach (KeyValuePair<long, Cell> kv in Cells)
                {
                    int row = (int)(kv.Key >> 14);
                    int col = (int)(kv.Key & 0x3FFF);
                    if (row >= r1 && row <= r2 && col >= c1 && col <= c2) keys.Add(kv.Key);
                }
                foreach (long key in keys)
                {
                    Cell c;
                    if (!Cells.TryGetValue(key, out c)) continue;
                    c.ClearContents();
                    if (c.Style == null) Cells.Remove(key);
                }
                return;
            }
            for (int r = r1; r <= r2; r++)
                for (int c = c1; c <= c2; c++)
                    ClearContents(r, c);
        }

        public void RecalculateUsedRange()
        {
            UsedMaxRow = -1;
            UsedMaxCol = -1;
            foreach (KeyValuePair<long, Cell> kv in Cells)
            {
                int row = (int)(kv.Key >> 14);
                int col = (int)(kv.Key & 0x3FFF);
                if (row > UsedMaxRow) UsedMaxRow = row;
                if (col > UsedMaxCol) UsedMaxCol = col;
            }
            foreach (MergeRange m in Merges)
            {
                if (m.Row2 > UsedMaxRow) UsedMaxRow = m.Row2;
                if (m.Col2 > UsedMaxCol) UsedMaxCol = m.Col2;
            }
        }

        public int EffectiveMaxRow
        {
            get { return UsedMaxRow < 0 ? 0 : UsedMaxRow; }
        }

        public int EffectiveMaxCol
        {
            get { return UsedMaxCol < 0 ? 0 : UsedMaxCol; }
        }

        public float GetColWidth(int col)
        {
            float w;
            if (ColWidths.TryGetValue(col, out w)) return w;
            return 96f;
        }

        public void SetColWidth(int col, float w)
        {
            ColWidths[col] = Math.Max(28f, Math.Min(800f, w));
            WidthsDirty = true;
        }

        public void SetColWidthImported(int col, float w)
        {
            ColWidths[col] = Math.Max(28f, Math.Min(800f, w));
        }

        public float GetRowHeightPoints(int row)
        {
            float h;
            if (RowHeights.TryGetValue(row, out h)) return h;
            return 15.75f;
        }

        public void SetRowHeightPoints(int row, float points)
        {
            RowHeights[row] = Math.Max(0f, Math.Min(409f, points));
            RowsDirty = true;
        }

        public void SetRowHeightPointsImported(int row, float points)
        {
            RowHeights[row] = Math.Max(0f, Math.Min(409f, points));
        }

        public void SetColumnHidden(int col, bool hidden)
        {
            if (hidden) HiddenCols.Add(col);
            else HiddenCols.Remove(col);
            WidthsDirty = true;
        }

        public void SetColumnHiddenImported(int col, bool hidden)
        {
            if (hidden) HiddenCols.Add(col);
            else HiddenCols.Remove(col);
        }

        public void SetRowHidden(int row, bool hidden)
        {
            if (hidden) HiddenRows.Add(row);
            else HiddenRows.Remove(row);
            RowsDirty = true;
        }

        public void SetRowHiddenImported(int row, bool hidden)
        {
            if (hidden) HiddenRows.Add(row);
            else HiddenRows.Remove(row);
        }

        public MergeRange GetMergeAt(int row, int col)
        {
            MergeRange m;
            return _mergeLookup.TryGetValue(Key(row, col), out m) ? m : null;
        }

        public void RebuildMergeLookup()
        {
            _mergeLookup.Clear();
            foreach (MergeRange m in Merges)
            {
                if (m.Row2 - m.Row1 > 10000 || m.Col2 - m.Col1 > 1000) continue;
                for (int r = m.Row1; r <= m.Row2; r++)
                    for (int c = m.Col1; c <= m.Col2; c++)
                        _mergeLookup[Key(r, c)] = m;
            }
        }

        public void AddMerge(MergeRange m)
        {
            for (int i = Merges.Count - 1; i >= 0; i--)
                if (Merges[i].Intersects(m.Row1, m.Col1, m.Row2, m.Col2))
                    Merges.RemoveAt(i);
            Merges.Add(m);
            RebuildMergeLookup();
            MergesDirty = true;
            if (m.Row2 > UsedMaxRow) UsedMaxRow = m.Row2;
            if (m.Col2 > UsedMaxCol) UsedMaxCol = m.Col2;
        }

        public void RemoveMergeAt(int row, int col)
        {
            MergeRange m = GetMergeAt(row, col);
            if (m == null) return;
            Merges.Remove(m);
            RebuildMergeLookup();
            MergesDirty = true;
        }

        public void ClearMerges()
        {
            if (Merges.Count == 0) return;
            Merges.Clear();
            _mergeLookup.Clear();
            MergesDirty = true;
        }

        public void ShiftMergesForInsert(bool rows, int at, int count)
        {
            if (Merges.Count == 0) return;
            foreach (MergeRange m in Merges)
            {
                if (rows)
                {
                    if (m.Row1 >= at) m.Row1 += count;
                    if (m.Row2 >= at) m.Row2 += count;
                }
                else
                {
                    if (m.Col1 >= at) m.Col1 += count;
                    if (m.Col2 >= at) m.Col2 += count;
                }
            }
            RebuildMergeLookup();
            MergesDirty = true;
        }

        public void ShiftStructuredRangesForInsert(bool rows, int at, int count)
        {
            DataVersion++;
            if (count <= 0) return;
            foreach (TableFormat t in TableFormats)
            {
                if (rows)
                {
                    if (t.Row1 >= at) t.Row1 += count;
                    if (t.Row2 >= at) t.Row2 += count;
                }
                else
                {
                    if (t.Col1 >= at) t.Col1 += count;
                    if (t.Col2 >= at) t.Col2 += count;
                }
            }
            foreach (ConditionalFormat cf in ConditionalFormats)
            {
                if (rows)
                {
                    if (cf.Row1 >= at) cf.Row1 += count;
                    if (cf.Row2 >= at) cf.Row2 += count;
                }
                else
                {
                    if (cf.Col1 >= at) cf.Col1 += count;
                    if (cf.Col2 >= at) cf.Col2 += count;
                }
            }
            foreach (PivotRange p in PivotRanges)
            {
                if (rows)
                {
                    if (p.Row1 >= at) p.Row1 += count;
                    if (p.Row2 >= at) p.Row2 += count;
                }
                else
                {
                    if (p.Col1 >= at) p.Col1 += count;
                    if (p.Col2 >= at) p.Col2 += count;
                }
            }
            foreach (DataValidationRule v in Validations)
            {
                if (rows)
                {
                    if (v.Row1 >= at) v.Row1 += count;
                    if (v.Row2 >= at) v.Row2 += count;
                }
                else
                {
                    if (v.Col1 >= at) v.Col1 += count;
                    if (v.Col2 >= at) v.Col2 += count;
                }
            }
        }

        public void ShiftStructuredRangesForDelete(bool rows, int at, int count)
        {
            DataVersion++;
            if (count <= 0) return;
            int end = at + count - 1;
            for (int i = TableFormats.Count - 1; i >= 0; i--)
            {
                TableFormat t = TableFormats[i];
                if (rows) ShiftSpan(ref t.Row1, ref t.Row2, at, end, count);
                else ShiftSpan(ref t.Col1, ref t.Col2, at, end, count);
                if (t.Row2 < t.Row1 || t.Col2 < t.Col1) TableFormats.RemoveAt(i);
            }
            for (int i = ConditionalFormats.Count - 1; i >= 0; i--)
            {
                ConditionalFormat c = ConditionalFormats[i];
                if (rows) ShiftSpan(ref c.Row1, ref c.Row2, at, end, count);
                else ShiftSpan(ref c.Col1, ref c.Col2, at, end, count);
                if (c.Row2 < c.Row1 || c.Col2 < c.Col1) ConditionalFormats.RemoveAt(i);
            }
            for (int i = PivotRanges.Count - 1; i >= 0; i--)
            {
                PivotRange p = PivotRanges[i];
                if (rows) ShiftSpan(ref p.Row1, ref p.Row2, at, end, count);
                else ShiftSpan(ref p.Col1, ref p.Col2, at, end, count);
                if (p.Row2 < p.Row1 || p.Col2 < p.Col1) PivotRanges.RemoveAt(i);
            }
            for (int i = Validations.Count - 1; i >= 0; i--)
            {
                DataValidationRule v = Validations[i];
                if (rows) ShiftSpan(ref v.Row1, ref v.Row2, at, end, count);
                else ShiftSpan(ref v.Col1, ref v.Col2, at, end, count);
                if (v.Row2 < v.Row1 || v.Col2 < v.Col1) Validations.RemoveAt(i);
            }
        }

        private static void ShiftSpan(ref int a, ref int b, int at, int end, int count)
        {
            if (b < at) return;
            if (a > end) { a -= count; b -= count; return; }
            if (a < at) b = Math.Max(at - 1, b - count);
            else { a = at; b = Math.Max(at - 1, b - count); }
        }

        public void ShiftMergesForDelete(bool rows, int at, int count)
        {
            if (Merges.Count == 0) return;
            int end = at + count - 1;
            for (int i = Merges.Count - 1; i >= 0; i--)
            {
                MergeRange m = Merges[i];
                int a = rows ? m.Row1 : m.Col1;
                int b = rows ? m.Row2 : m.Col2;
                if (b < at) continue;
                if (a > end)
                {
                    if (rows) { m.Row1 -= count; m.Row2 -= count; }
                    else { m.Col1 -= count; m.Col2 -= count; }
                    continue;
                }
                if (a >= at && b <= end) { Merges.RemoveAt(i); continue; }
                if (a < at && b > end)
                {
                    if (rows) m.Row2 -= count; else m.Col2 -= count;
                }
                else if (a < at)
                {
                    if (rows) m.Row2 = at - 1; else m.Col2 = at - 1;
                }
                else
                {
                    if (rows) { m.Row1 = at; m.Row2 -= count; }
                    else { m.Col1 = at; m.Col2 -= count; }
                }
                if (m.Row2 < m.Row1 || m.Col2 < m.Col1) Merges.RemoveAt(i);
            }
            RebuildMergeLookup();
            MergesDirty = true;
        }

        public IEnumerable<Cell> EnumerateCells()
        {
            foreach (KeyValuePair<long, Cell> kv in Cells) yield return kv.Value;
        }

        public Sheet Clone(Workbook owner)
        {
            Sheet s = new Sheet(owner, Name);
            s.UsedMaxRow = UsedMaxRow;
            s.DataVersion = DataVersion;
            s.UsedMaxCol = UsedMaxCol;
            foreach (KeyValuePair<long, Cell> kv in Cells)
                s.Cells[kv.Key] = kv.Value.Clone();
            foreach (KeyValuePair<int, float> kv in ColWidths)
                s.ColWidths[kv.Key] = kv.Value;
            foreach (KeyValuePair<int, float> kv in RowHeights)
                s.RowHeights[kv.Key] = kv.Value;
            foreach (int r in HiddenRows) s.HiddenRows.Add(r);
            foreach (int c in HiddenCols) s.HiddenCols.Add(c);
            foreach (MergeRange m in Merges) s.Merges.Add(m.Clone());
            foreach (SheetChart ch in Charts) s.Charts.Add(ch.Clone());
            foreach (SheetImage img in Images) s.Images.Add(img.Clone());
            foreach (SheetShape sh in Shapes) s.Shapes.Add(sh.Clone());
            foreach (PivotRange pr in PivotRanges) s.PivotRanges.Add(pr.Clone());
            foreach (SheetSlicer sl in Slicers) s.Slicers.Add(sl.Clone());
            foreach (SheetImage img in DeletedImages) s.DeletedImages.Add(img.Clone());
            s.ImagesDirty = ImagesDirty;
            s.HasSlicerPlaceholders = HasSlicerPlaceholders;
            s.HasPicturePlaceholders = HasPicturePlaceholders;
            s.HasChartPlaceholders = HasChartPlaceholders;
            s.ShowGridLines = ShowGridLines;
            s.PrintGridLines = PrintGridLines;
            s.PrintHeadings = PrintHeadings;
            s.PrintLandscape = PrintLandscape;
            s.PrintPaperSize = PrintPaperSize;
            s.PrintScalePercent = PrintScalePercent;
            s.PrintFitToWidth = PrintFitToWidth;
            s.PrintFitToHeight = PrintFitToHeight;
            s.PrintMarginMm = PrintMarginMm;
            s.PrintArea = PrintArea;
            s.PrintHeader = PrintHeader;
            s.PrintFooter = PrintFooter;
            s.PrintTitleRows = PrintTitleRows;
            s.PrintTitleCols = PrintTitleCols;
            s.IsProtected = IsProtected;
            s.ProtectionPasswordHash = ProtectionPasswordHash;
            foreach (int br in PrintRowBreaks) s.PrintRowBreaks.Add(br);
            foreach (ConditionalFormat cf in ConditionalFormats) s.ConditionalFormats.Add(cf.Clone());
            foreach (TableFormat tf in TableFormats)
            {
                TableFormat copy = tf.Clone();
                copy.Sheet = s;
                s.TableFormats.Add(copy);
            }
            foreach (DataValidationRule rule in Validations) s.Validations.Add(rule.Clone());
            s.RebuildMergeLookup();
            s.FreezeRows = FreezeRows;
            s.FreezeCols = FreezeCols;
            s.AutoFilterRef = AutoFilterRef;
            foreach (FilterSelection filter in FilterSelections) s.FilterSelections.Add(filter.Clone());
            s.SourcePart = SourcePart;
            s.WidthsDirty = WidthsDirty;
            s.RowsDirty = RowsDirty;
            s.MergesDirty = MergesDirty;
            s.HyperlinksDirty = HyperlinksDirty;
            s.StructureDirty = StructureDirty;
            s.ChartsDirty = ChartsDirty;
            return s;
        }
    }

    public sealed class Workbook
    {
        public readonly List<Sheet> Sheets = new List<Sheet>();
        public readonly List<DefinedName> DefinedNames = new List<DefinedName>();
        public readonly List<string> LoadWarnings = new List<string>();
        public readonly HashSet<string> UnsupportedFunctions =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public int ActiveSheetIndex;
        public XlsxSourceInfo XlsxSource;
        public string LegacyXlsPath;
        public bool Date1904;
        public System.Drawing.Color[] ThemePalette;

        public Workbook()
        {
            AddSheet("Hoja1");
        }

        public Sheet ActiveSheet
        {
            get
            {
                if (Sheets.Count == 0) AddSheet("Hoja1");
                if (ActiveSheetIndex < 0) ActiveSheetIndex = 0;
                if (ActiveSheetIndex >= Sheets.Count) ActiveSheetIndex = Sheets.Count - 1;
                return Sheets[ActiveSheetIndex];
            }
        }

        public void AddLoadWarning(string warning)
        {
            if (string.IsNullOrEmpty(warning)) return;
            if (!LoadWarnings.Contains(warning)) LoadWarnings.Add(warning);
        }

        public void AddUnsupportedFunction(string name)
        {
            if (string.IsNullOrEmpty(name)) return;
            if (UnsupportedFunctions.Count < 40) UnsupportedFunctions.Add(name.Trim());
        }

        public Sheet AddSheet(string name)
        {
            if (string.IsNullOrEmpty(name)) name = UniqueSheetName();
            Sheet s = new Sheet(this, name);
            Sheets.Add(s);
            return s;
        }

        public string UniqueSheetName()
        {
            int n = Sheets.Count + 1;
            while (true)
            {
                string name = "Hoja" + n.ToString(CultureInfo.InvariantCulture);
                bool exists = false;
                foreach (Sheet sh in Sheets) if (string.Equals(sh.Name, name, StringComparison.OrdinalIgnoreCase)) { exists = true; break; }
                if (!exists) return name;
                n++;
            }
        }

        public Sheet FindSheet(string name)
        {
            if (string.IsNullOrEmpty(name)) return null;
            foreach (Sheet s in Sheets)
                if (string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase)) return s;
            return null;
        }

        public DefinedName FindDefinedName(string name, int sheetIndex)
        {
            if (string.IsNullOrEmpty(name)) return null;
            DefinedName global = null;
            foreach (DefinedName d in DefinedNames)
            {
                if (!string.Equals(d.Name, name, StringComparison.OrdinalIgnoreCase)) continue;
                if (d.LocalSheetId == sheetIndex) return d;
                if (d.LocalSheetId < 0) global = d;
            }
            return global;
        }

        public Workbook Clone()
        {
            Workbook wb = new Workbook();
            wb.Sheets.Clear();
            foreach (Sheet s in Sheets) wb.Sheets.Add(s.Clone(wb));
            wb.ActiveSheetIndex = ActiveSheetIndex;
            foreach (DefinedName d in DefinedNames) wb.DefinedNames.Add(d.Clone());
            wb.XlsxSource = XlsxSource == null ? null : XlsxSource.Clone();
            wb.LegacyXlsPath = LegacyXlsPath;
            wb.Date1904 = Date1904;
            wb.ThemePalette = ThemePalette == null ? null : (System.Drawing.Color[])ThemePalette.Clone();
            foreach (string warning in LoadWarnings) wb.LoadWarnings.Add(warning);
            foreach (string fn in UnsupportedFunctions) wb.UnsupportedFunctions.Add(fn);
            return wb;
        }

        public void Recalculate()
        {
            FormulaEngine.Recalculate(this);
        }
    }

    public sealed class SelectionStats
    {
        public int NumericCount;
        public int NonEmptyCount;
        public double Sum;
        public double Min;
        public double Max;
        public bool HasNumbers { get { return NumericCount > 0; } }
        public double Average { get { return NumericCount == 0 ? 0 : Sum / NumericCount; } }
    }

    public static class SheetStats
    {
        public static SelectionStats Compute(Sheet sheet, int r1, int c1, int r2, int c2)
        {
            SelectionStats st = new SelectionStats();
            st.Min = double.MaxValue;
            st.Max = double.MinValue;
            if (sheet == null) return st;
            foreach (KeyValuePair<long, Cell> kv in sheet.Cells)
            {
                int row = (int)(kv.Key >> 14);
                int col = (int)(kv.Key & 0x3FFF);
                if (row < r1 || row > r2 || col < c1 || col > c2) continue;
                Cell cell = kv.Value;
                if (cell.Value.Kind == ValueKind.Empty) continue;
                st.NonEmptyCount++;
                if (cell.Value.Kind == ValueKind.Number)
                {
                    st.NumericCount++;
                    st.Sum += cell.Value.Number;
                    if (cell.Value.Number < st.Min) st.Min = cell.Value.Number;
                    if (cell.Value.Number > st.Max) st.Max = cell.Value.Number;
                }
            }
            return st;
        }
    }

    public sealed class UndoManager
    {
        private readonly List<Workbook> _undo = new List<Workbook>();
        private readonly List<Workbook> _redo = new List<Workbook>();
        public int Limit = 50;

        public bool CanUndo { get { return _undo.Count > 0; } }
        public bool CanRedo { get { return _redo.Count > 0; } }

        public void Push(Workbook current)
        {
            long cells = 0;
            foreach (Sheet s in current.Sheets) cells += s.Cells.Count;
            if (cells > 2000000)
            {
                // Un snapshot completo de un libro gigantesco consume más memoria
                // que el propio libro; es preferible no conservar undo a agotarla.
                _undo.Clear();
                _redo.Clear();
                return;
            }
            // Presupuesto de celdas copiadas: evita que un libro grande
            // multiplique su memoria por el número de instantáneas.
            const long MaxCopiedCells = 3000000;
            long perSnapshot = Math.Max(1, cells);
            Limit = (int)Math.Max(1, Math.Min(50, MaxCopiedCells / perSnapshot));
            if (cells > 250000) Limit = Math.Min(Limit, 5);
            _undo.Add(current.Clone());
            if (_undo.Count > Limit) _undo.RemoveAt(0);
            _redo.Clear();
        }

        public Workbook Undo(Workbook current)
        {
            if (_undo.Count == 0) return null;
            Workbook prev = _undo[_undo.Count - 1];
            _undo.RemoveAt(_undo.Count - 1);
            _redo.Add(current.Clone());
            return prev;
        }

        public Workbook Redo(Workbook current)
        {
            if (_redo.Count == 0) return null;
            Workbook next = _redo[_redo.Count - 1];
            _redo.RemoveAt(_redo.Count - 1);
            _undo.Add(current.Clone());
            return next;
        }

        public void Clear()
        {
            _undo.Clear();
            _redo.Clear();
        }
    }
}
