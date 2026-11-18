using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace CalcRapido
{
    // Traductor de la pila RPN de BIFF8 a una fórmula A1 comprensible por
    // el motor interno. No intenta cubrir todas las funciones arcanas de
    // Excel: si una fórmula no se entiende se conserva el valor cacheado.
    internal static class XlsFormula
    {
        private static readonly Dictionary<int, string> Names = BuildNames();
        private static readonly Dictionary<int, int> FixedArity = BuildArity();

        public static string Read(byte[] data, int pos, int length, List<string> sheetNames)
        {
            return Read(data, pos, length, sheetNames, null, 0, 0);
        }

        public static string Read(byte[] data, int pos, int length, List<string> sheetNames,
                                  List<string> definedNames)
        {
            return Read(data, pos, length, sheetNames, definedNames, 0, 0);
        }

        public static string Read(byte[] data, int pos, int length, List<string> sheetNames,
                                  List<string> definedNames, int baseRow, int baseCol)
        {
            if (length <= 0) return null;
            List<string> stack = new List<string>();
            int end = pos + length;
            int p = pos;
            while (p < end)
            {
                int id = data[p++];
                int baseId = id & 0x1F;
                if (id >= 0x03 && id <= 0x11)
                {
                    if (stack.Count < 2) return null;
                    string b = stack[stack.Count - 1];
                    string a = stack[stack.Count - 2];
                    stack.RemoveAt(stack.Count - 1);
                    stack.RemoveAt(stack.Count - 1);
                    string op;
                    switch (id)
                    {
                        case 0x03: op = "+"; break;
                        case 0x04: op = "-"; break;
                        case 0x05: op = "*"; break;
                        case 0x06: op = "/"; break;
                        case 0x07: op = "^"; break;
                        case 0x08: op = "&"; break;
                        case 0x09: op = "<"; break;
                        case 0x0A: op = "<="; break;
                        case 0x0B: op = "="; break;
                        case 0x0C: op = ">="; break;
                        case 0x0D: op = ">"; break;
                        case 0x0E: op = "<>"; break;
                        case 0x0F: op = " "; break;
                        case 0x10: op = ";"; break;
                        case 0x11: op = ":"; break;
                        default: return null;
                    }
                    stack.Add("(" + a + op + b + ")");
                    continue;
                }
                switch (id)
                {
                    case 0x12: // uplus
                    case 0x13: // uminus
                        {
                            if (stack.Count < 1) return null;
                            string a = stack[stack.Count - 1];
                            stack[stack.Count - 1] = (id == 0x13 ? "-(" : "+(") + a + ")";
                            break;
                        }
                    case 0x14: // percent
                        {
                            if (stack.Count < 1) return null;
                            string a = stack[stack.Count - 1];
                            stack[stack.Count - 1] = "(" + a + "%)";
                            break;
                        }
                    case 0x15: // paren
                        if (stack.Count < 1) return null;
                        stack[stack.Count - 1] = "(" + stack[stack.Count - 1] + ")";
                        break;
                    case 0x16: // missing arg
                        stack.Add("");
                        break;
                    case 0x17: // string
                        {
                            int cch = data[p++];
                            int flags = data[p++];
                            bool high = (flags & 1) != 0;
                            StringBuilder sb = new StringBuilder(cch);
                            for (int i = 0; i < cch && p < end; i++)
                            {
                                if (high) { sb.Append((char)U16(data, p)); p += 2; }
                                else sb.Append((char)data[p++]);
                            }
                            stack.Add("\"" + sb.ToString().Replace("\"", "\"\"") + "\"");
                            break;
                        }
                    case 0x19: // attr: grbit (1 byte) + data (2 bytes)
                        {
                            if (p + 3 > end) return null;
                            int grbit = data[p++];
                            int extra = U16(data, p);
                            p += 2;
                            if ((grbit & 0x10) != 0 && stack.Count > 0)
                                stack[stack.Count - 1] = "SUMA(" + stack[stack.Count - 1] + ")";
                            if ((grbit & 0x04) != 0) p += (extra + 1) * 2;
                            break;
                        }
                    case 0x1C: // error
                        stack.Add(ErrorText(data[p++]));
                        break;
                    case 0x1D: // bool
                        stack.Add(data[p++] != 0 ? "VERDADERO" : "FALSO");
                        break;
                    case 0x1E: // int
                        stack.Add(U16(data, p).ToString(CultureInfo.InvariantCulture));
                        p += 2;
                        break;
                    case 0x1F: // num
                        stack.Add(BitConverter.ToDouble(data, p).ToString("G17", CultureInfo.InvariantCulture));
                        p += 8;
                        break;
                    case 0x01: // shared formula reference (ptgExp)
                    case 0x02: // table (ptgTbl)
                    case 0x60: // array constant (ptgArray)
                        return null;
                    case 0x26: // ptgMemArea: marcador, no altera la pila
                    case 0x46:
                        if (p + 6 > end) return null;
                        p += 6;
                        break;
                    case 0x27: // ptgMemErr
                    case 0x47:
                        if (p + 6 > end) return null;
                        p += 6;
                        break;
                    case 0x29: // ptgMemFunc
                    case 0x49:
                        {
                            if (p + 2 > end) return null;
                            int nb = U16(data, p); p += 2;
                            if (nb < 0 || p + nb > end) return null;
                            p += nb;
                            break;
                        }
                    case 0x41: // function fixed (ptgFunc)
                        {
                            if (p + 2 > end) return null;
                            int fn = U16(data, p); p += 2;
                            string name;
                            int arity;
                            if (!Names.TryGetValue(fn, out name) || !FixedArity.TryGetValue(fn, out arity)) return null;
                            if (stack.Count < arity) return null;
                            string[] args = new string[arity];
                            for (int i = arity - 1; i >= 0; i--)
                            {
                                args[i] = stack[stack.Count - 1];
                                stack.RemoveAt(stack.Count - 1);
                            }
                            stack.Add(name + "(" + string.Join(";", args) + ")");
                            break;
                        }
                    case 0x42: // function variable (ptgFuncVar)
                        {
                            if (p + 3 > end) return null;
                            int argc = data[p++];
                            int fn = U16(data, p); p += 2;
                            string name;
                            if (!Names.TryGetValue(fn, out name)) return null;
                            if (stack.Count < argc) return null;
                            string[] args = new string[argc];
                            for (int i = argc - 1; i >= 0; i--)
                            {
                                args[i] = stack[stack.Count - 1];
                                stack.RemoveAt(stack.Count - 1);
                            }
                            stack.Add(name + "(" + string.Join(";", args) + ")");
                            break;
                        }
                    case 0x23: // name (ptgName, BIFF8)
                    case 0x43: // alias usado por algunos escritores
                        {
                            if (p + 4 > end || definedNames == null) return null;
                            int index = (int)U32(data, p) - 1;
                            p += 4;
                            if (index < 0 || index >= definedNames.Count) return null;
                            stack.Add(definedNames[index]);
                            break;
                        }
                    case 0x39: // name X (ptgNameX, BIFF8): ixti + nameindex
                    case 0x59:
                        {
                            if (p + 6 > end || definedNames == null) return null;
                            p += 2; // ixti
                            int index = (int)U32(data, p) - 1;
                            p += 4;
                            if (index < 0 || index >= definedNames.Count) return null;
                            stack.Add(definedNames[index]);
                            break;
                        }
                    default:
                        if (baseId == 0x04) // ref
                        {
                            string r = ReadRef(data, ref p, sheetNames, false, baseRow, baseCol);
                            if (r == null) return null;
                            stack.Add(r);
                        }
                        else if (baseId == 0x05) // area
                        {
                            string r = ReadArea(data, ref p, sheetNames, false, baseRow, baseCol);
                            if (r == null) return null;
                            stack.Add(r);
                        }
                        else if (baseId == 0x0A) { stack.Add("#REF!"); p += 4; }
                        else if (baseId == 0x0B) { stack.Add("#REF!"); p += 8; }
                        else if (baseId == 0x0C) { string r = ReadRef(data, ref p, sheetNames, true, baseRow, baseCol); if (r == null) return null; stack.Add(r); }
                        else if (baseId == 0x0D) { string r = ReadArea(data, ref p, sheetNames, true, baseRow, baseCol); if (r == null) return null; stack.Add(r); }
                        else if (baseId == 0x1A || baseId == 0x1C) // ref3d / refErr3d
                        {
                            if (p + 2 > end) return null;
                            int ixti = U16(data, p); p += 2;
                            string r = ReadRef(data, ref p, sheetNames, false, baseRow, baseCol);
                            if (r == null) return null;
                            stack.Add(Qualify(sheetNames, ixti, r));
                        }
                        else if (baseId == 0x1B || baseId == 0x1D) // area3d / areaErr3d
                        {
                            if (p + 2 > end) return null;
                            int ixti = U16(data, p); p += 2;
                            string r = ReadArea(data, ref p, sheetNames, false, baseRow, baseCol);
                            if (r == null) return null;
                            stack.Add(Qualify(sheetNames, ixti, r));
                        }
                        else return null;
                        break;
                }
            }
            if (stack.Count == 0) return null;
            return stack[stack.Count - 1];
        }

        private static string Qualify(List<string> names, int ixti, string reference)
        {
            int idx = ixti;
            if (idx < 0 || idx >= names.Count || names[idx] == null) return reference;
            string name = names[idx].Replace("'", "''");
            return "'" + name + "'!" + reference;
        }

        private static string ReadRef(byte[] data, ref int p, List<string> names, bool normalized,
                                      int baseRow, int baseCol)
        {
            if (p + 4 > data.Length) return null;
            int row = U16(data, p);
            int col = U16(data, p + 2);
            p += 4;
            bool rowRel = (col & 0x8000) != 0;
            bool colRel = (col & 0x4000) != 0;
            int r = row, c = col & 0x00FF;
            if (normalized)
            {
                if (rowRel && r >= 32768) r -= 65536;
                if (colRel && c >= 128) c -= 256;
                if (rowRel) r += baseRow;
                if (colRel) c += baseCol;
            }
            if (r < 0 || c < 0 || r >= Limits.MaxRows || c >= Limits.MaxCols) return "#REF!";
            return (colRel ? "" : "$") + AddressUtil.ColumnName(c) + (rowRel ? "" : "$") + (r + 1);
        }

        private static string ReadArea(byte[] data, ref int p, List<string> names, bool normalized,
                                       int baseRow, int baseCol)
        {
            if (p + 8 > data.Length) return null;
            int r1 = U16(data, p);
            int r2 = U16(data, p + 2);
            int c1 = U16(data, p + 4);
            int c2 = U16(data, p + 6);
            p += 8;
            bool rr1 = (c1 & 0x8000) != 0, rr2 = (c2 & 0x8000) != 0;
            bool cc1 = (c1 & 0x4000) != 0, cc2 = (c2 & 0x4000) != 0;
            int cc1v = c1 & 0x00FF, cc2v = c2 & 0x00FF;
            int rr1v = r1, rr2v = r2;
            if (normalized)
            {
                if (rr1 && rr1v >= 32768) rr1v -= 65536;
                if (rr2 && rr2v >= 32768) rr2v -= 65536;
                if (cc1 && cc1v >= 128) cc1v -= 256;
                if (cc2 && cc2v >= 128) cc2v -= 256;
                if (rr1) rr1v += baseRow;
                if (rr2) rr2v += baseRow;
                if (cc1) cc1v += baseCol;
                if (cc2) cc2v += baseCol;
            }
            if (rr1v < 0 || rr2v < 0 || cc1v < 0 || cc2v < 0 ||
                rr1v >= Limits.MaxRows || rr2v >= Limits.MaxRows ||
                cc1v >= Limits.MaxCols || cc2v >= Limits.MaxCols) return "#REF!";
            string a = (cc1 ? "" : "$") + AddressUtil.ColumnName(cc1v) + (rr1 ? "" : "$") + (rr1v + 1);
            string b = (cc2 ? "" : "$") + AddressUtil.ColumnName(cc2v) + (rr2 ? "" : "$") + (rr2v + 1);
            return a + ":" + b;
        }

        private static string ErrorText(int code)
        {
            switch (code)
            {
                case 0x00: return "#NULL!";
                case 0x07: return "#DIV/0!";
                case 0x0F: return "#VALUE!";
                case 0x17: return "#REF!";
                case 0x1D: return "#NAME?";
                case 0x24: return "#NUM!";
                case 0x2A: return "#N/A";
                default: return "#VALUE!";
            }
        }

        public static bool IsDateStyle(CellStyle st)
        {
            if (st == null) return false;
            if (st.Format == NumberFormatKind.Date || st.Format == NumberFormatKind.Time) return true;
            string code = (st.CustomFormat ?? "").ToLowerInvariant();
            return code.IndexOf("yy", StringComparison.Ordinal) >= 0 ||
                   code.IndexOf("dd", StringComparison.Ordinal) >= 0 ||
                   code.IndexOf("hh", StringComparison.Ordinal) >= 0 ||
                   code.IndexOf("ss", StringComparison.Ordinal) >= 0;
        }

        private static Dictionary<int, string> BuildNames()
        {
            Dictionary<int, string> d = new Dictionary<int, string>();
            d[0] = "COUNT"; d[1] = "IF"; d[2] = "ISNA"; d[3] = "ISERROR";
            d[4] = "SUM"; d[5] = "AVERAGE"; d[6] = "MIN"; d[7] = "MAX";
            d[8] = "ROW"; d[9] = "COLUMN"; d[10] = "NA"; d[11] = "NPV";
            d[12] = "STDEV"; d[13] = "DOLLAR"; d[14] = "FIXED"; d[15] = "SIN";
            d[16] = "COS"; d[17] = "TAN"; d[18] = "ATAN"; d[19] = "PI";
            d[20] = "SQRT"; d[21] = "EXP"; d[22] = "LN"; d[23] = "LOG10";
            d[24] = "ABS"; d[25] = "INT"; d[26] = "SIGN"; d[27] = "ROUND";
            d[28] = "LOOKUP"; d[29] = "INDEX"; d[30] = "REPT"; d[31] = "MID";
            d[32] = "LEN"; d[33] = "VALUE"; d[34] = "TRUE"; d[35] = "FALSE";
            d[36] = "AND"; d[37] = "OR"; d[38] = "NOT"; d[39] = "MOD";
            d[63] = "RAND"; d[64] = "MATCH"; d[65] = "DATE"; d[66] = "TIME";
            d[67] = "DAY"; d[68] = "MONTH"; d[69] = "YEAR"; d[70] = "WEEKDAY";
            d[71] = "HOUR"; d[72] = "MINUTE"; d[73] = "SECOND"; d[74] = "NOW";
            d[75] = "AREAS"; d[76] = "ROWS"; d[77] = "COLUMNS"; d[78] = "OFFSET";
            d[82] = "SEARCH"; d[97] = "ATAN2"; d[98] = "ASIN"; d[99] = "ACOS";
            d[100] = "CHOOSE"; d[101] = "HLOOKUP"; d[102] = "VLOOKUP";
            d[109] = "LOG"; d[111] = "CHAR"; d[112] = "LOWER"; d[113] = "UPPER";
            d[114] = "PROPER"; d[115] = "LEFT"; d[116] = "RIGHT"; d[117] = "EXACT";
            d[118] = "TRIM"; d[119] = "REPLACE"; d[120] = "SUBSTITUTE";
            d[121] = "CODE"; d[124] = "FIND"; d[127] = "ISTEXT"; d[128] = "ISNUMBER";
            d[129] = "ISBLANK"; d[130] = "T"; d[131] = "N";
            d[140] = "DATEVALUE"; d[141] = "TIMEVALUE"; d[148] = "INDIRECT";
            d[162] = "CLEAN"; d[169] = "COUNTA"; d[183] = "PRODUCT";
            d[190] = "ISNONTEXT"; d[193] = "STDEVP"; d[194] = "VARP"; d[197] = "TRUNC";
            d[212] = "ROUNDUP"; d[213] = "ROUNDDOWN"; d[216] = "RANK";
            d[219] = "ADDRESS"; d[220] = "DAYS360"; d[221] = "TODAY";
            d[227] = "MEDIAN"; d[228] = "SUMPRODUCT"; d[229] = "SINH";
            d[230] = "COSH"; d[231] = "TANH"; d[232] = "ASINH"; d[233] = "ACOSH";
            d[234] = "ATANH"; d[244] = "INFO"; d[247] = "DB"; d[252] = "FREQUENCY";
            d[261] = "ERROR.TYPE"; d[269] = "AVEDEV"; d[279] = "EVEN";
            d[285] = "FLOOR"; d[288] = "CEILING"; d[298] = "ODD";
            d[299] = "PERMUT"; d[303] = "SUMXMY2"; d[304] = "SUMX2MY2";
            d[305] = "SUMX2PY2"; d[306] = "CHITEST"; d[307] = "CORREL";
            d[308] = "COVAR"; d[309] = "FORECAST"; d[310] = "FTEST";
            d[311] = "INTERCEPT"; d[312] = "PEARSON"; d[313] = "RSQ";
            d[314] = "STEYX"; d[315] = "SLOPE"; d[316] = "TTEST"; d[317] = "PROB";
            d[318] = "DEVSQ"; d[319] = "GEOMEAN"; d[320] = "HARMEAN";
            d[321] = "SUMSQ"; d[322] = "KURT"; d[323] = "SKEW"; d[324] = "ZTEST";
            d[325] = "LARGE"; d[326] = "SMALL"; d[327] = "QUARTILE";
            d[328] = "PERCENTILE"; d[329] = "PERCENTRANK"; d[330] = "MODE";
            d[276] = "COMBIN"; d[278] = "CRITBINOM"; d[279] = "EVEN";
            d[280] = "EXPONDIST"; d[281] = "FDIST"; d[282] = "FINV";
            d[283] = "FISHER"; d[284] = "FISHERINV"; d[285] = "FLOOR";
            d[286] = "GAMMADIST"; d[287] = "GAMMAINV"; d[288] = "CEILING";
            d[289] = "HYPGEOMDIST"; d[290] = "LOGNORMDIST"; d[291] = "LOGINV";
            d[292] = "NEGBINOMDIST"; d[293] = "NORMDIST"; d[294] = "NORMSDIST";
            d[295] = "NORMINV";
            d[331] = "TRIMMEAN"; d[336] = "CONCATENATE"; d[337] = "POWER";
            d[342] = "RADIANS"; d[343] = "DEGREES"; d[344] = "SUBTOTAL";
            d[345] = "SUMIF"; d[346] = "COUNTIF"; d[347] = "COUNTBLANK";
            d[348] = "ISPMT"; d[349] = "DATEDIF"; d[352] = "ROMAN";
            d[353] = "GETPIVOTDATA"; d[354] = "HYPERLINK"; d[355] = "PHONETIC";
            d[356] = "AVERAGEA"; d[357] = "MAXA"; d[358] = "MINA";
            d[359] = "STDEVPA"; d[360] = "VARPA"; d[361] = "STDEVA"; d[362] = "VARA";
            return d;
        }

        private static Dictionary<int, int> BuildArity()
        {
            Dictionary<int, int> d = new Dictionary<int, int>();
            int[] zero = new int[] { 10, 19, 34, 35, 63, 74, 221 };
            int[] one = new int[] { 2, 3, 8, 9, 15, 16, 17, 18, 20, 21, 22, 23, 24, 25, 26,
                30, 32, 33, 67, 68, 69, 70, 71, 72, 73, 75, 76, 77, 98, 99, 111, 112,
                113, 114, 118, 121, 127, 128, 129, 130, 131, 140, 141, 148, 162, 190,
                229, 230, 231, 232, 233, 234, 261, 269, 279, 283, 284, 294, 298, 342, 343, 347 };
            int[] two = new int[] { 11, 27, 36, 37, 39, 64, 97, 100, 117, 119, 212, 213,
                220, 276, 285, 288, 303, 304, 305, 306, 307, 308, 309, 310, 311, 312,
                313, 314, 315, 316, 317, 318, 319, 320, 321, 322, 323, 324, 325, 326,
                327, 328, 329, 330, 331, 337, 345, 346 };
            int[] three = new int[] { 14, 31, 120, 124, 348 };
            int[] four = new int[] { 13, 28, 29, 119, 349 };
            foreach (int i in zero) d[i] = 0;
            foreach (int i in one) d[i] = 1;
            foreach (int i in two) d[i] = 2;
            foreach (int i in three) d[i] = 3;
            foreach (int i in four) d[i] = 4;
            d[4] = 1; d[5] = 1; d[0] = 1; d[1] = 3; d[65] = 3; d[66] = 3;
            d[78] = 3; d[101] = 4; d[102] = 4; d[12] = 1; d[6] = 1; d[7] = 1;
            return d;
        }

        private static ushort U16(byte[] b, int p)
        {
            if (p < 0 || p + 2 > b.Length) return 0;
            return (ushort)(b[p] | (b[p + 1] << 8));
        }

        private static uint U32(byte[] b, int p)
        {
            if (p < 0 || p + 4 > b.Length) return 0;
            return (uint)(b[p] | (b[p + 1] << 8) | (b[p + 2] << 16) | (b[p + 3] << 24));
        }
    }
}
