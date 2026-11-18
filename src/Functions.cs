using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace CalcRapido
{
    internal static class Builtins
    {
        private static readonly Random Rng = new Random(Environment.TickCount);
        private static readonly Dictionary<string, string> DisplayNames =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private static readonly Dictionary<string, string> Aliases = BuildAliases();

        private static void AddAlias(Dictionary<string, string> d, string canonical, params string[] aliases)
        {
            d[canonical] = canonical;
            foreach (string a in aliases) d[a] = canonical;
            DisplayNames[canonical] = aliases.Length > 0 ? aliases[0] : canonical;
        }

        public static string DisplayFormula(string formula)
        {
            if (string.IsNullOrEmpty(formula)) return formula;
            StringBuilder sb = new StringBuilder(formula.Length + 8);
            int i = 0;
            while (i < formula.Length)
            {
                char ch = formula[i];
                if (ch == '"')
                {
                    int j = i + 1;
                    while (j < formula.Length)
                    {
                        if (formula[j] == '"')
                        {
                            if (j + 1 < formula.Length && formula[j + 1] == '"') { j += 2; continue; }
                            j++;
                            break;
                        }
                        j++;
                    }
                    int end = Math.Min(j, formula.Length);
                    sb.Append(formula.Substring(i, end - i));
                    i = end;
                    continue;
                }
                if (ch == '\'')
                {
                    int j = i + 1;
                    while (j < formula.Length)
                    {
                        if (formula[j] == '\'')
                        {
                            if (j + 1 < formula.Length && formula[j + 1] == '\'') { j += 2; continue; }
                            j++;
                            break;
                        }
                        j++;
                    }
                    int end = Math.Min(j, formula.Length);
                    sb.Append(formula.Substring(i, end - i));
                    i = end;
                    continue;
                }
                if (ch == '[')
                {
                    int depth = 0;
                    int j = i;
                    while (j < formula.Length)
                    {
                        if (formula[j] == '[') depth++;
                        else if (formula[j] == ']')
                        {
                            depth--;
                            if (depth == 0) { j++; break; }
                        }
                        j++;
                    }
                    sb.Append(formula.Substring(i, j - i));
                    i = j;
                    continue;
                }
                if (char.IsLetter(ch) || ch == '_')
                {
                    int start = i;
                    i++;
                    while (i < formula.Length && (char.IsLetterOrDigit(formula[i]) || formula[i] == '_' || formula[i] == '.'))
                        i++;
                    string token = formula.Substring(start, i - start);
                    if (token.StartsWith("_XLFN.", StringComparison.OrdinalIgnoreCase))
                        token = token.Substring(6);
                    else if (token.StartsWith("_XLWS.", StringComparison.OrdinalIgnoreCase))
                        token = token.Substring(6);
                    int k = i;
                    while (k < formula.Length && formula[k] == ' ') k++;
                    if (k < formula.Length && formula[k] == '(')
                    {
                        string display;
                        if (DisplayNames.TryGetValue(token, out display)) sb.Append(display);
                        else sb.Append(token);
                    }
                    else sb.Append(token);
                    continue;
                }
                if (ch == ',') sb.Append(';');
                else sb.Append(ch);
                i++;
            }
            return sb.ToString();
        }

        private static Dictionary<string, string> BuildAliases()
        {
            Dictionary<string, string> d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            AddAlias(d, "SUM", "SUMA");
            AddAlias(d, "AVERAGE", "PROMEDIO", "MEDIA");
            AddAlias(d, "MIN");
            AddAlias(d, "MAX");
            AddAlias(d, "COUNT", "CONTAR");
            AddAlias(d, "COUNTA", "CONTARA");
            AddAlias(d, "COUNTBLANK", "CONTAR.BLANCO");
            AddAlias(d, "MEDIAN", "MEDIANA");
            AddAlias(d, "STDEV", "DESVEST", "DESVEST.M", "DESVEST.MUESTRAL", "STDEV.S");
            AddAlias(d, "STDEVP", "DESVESTP", "DESVEST.P", "DESVEST.POBLACIONAL", "STDEV.P");
            AddAlias(d, "VAR", "VAR.M", "VAR.MUESTRAL", "VAR.S");
            AddAlias(d, "VARP", "VAR.P", "VAR.POBLACIONAL");
            AddAlias(d, "ROUND", "REDONDEAR");
            AddAlias(d, "ROUNDUP", "REDONDEAR.MAS", "REDONDEAR.ARRIBA");
            AddAlias(d, "ROUNDDOWN", "REDONDEAR.MENOS", "REDONDEAR.ABAJO");
            AddAlias(d, "FIXED", "DECIMAL");
            AddAlias(d, "MROUND", "REDONDEAR.MULTIPLO");
            AddAlias(d, "COMBIN", "COMBINAT");
            AddAlias(d, "EVEN", "REDONDEAR.PAR");
            AddAlias(d, "ODD", "REDONDEAR.IMPAR");
            AddAlias(d, "INT", "ENTERO");
            AddAlias(d, "TRUNC", "TRUNCAR");
            AddAlias(d, "ABS");
            AddAlias(d, "MOD", "RESIDUO");
            AddAlias(d, "SQRT", "RAIZ");
            AddAlias(d, "POWER", "POTENCIA");
            AddAlias(d, "EXP");
            AddAlias(d, "LN");
            AddAlias(d, "LOG");
            AddAlias(d, "LOG10");
            AddAlias(d, "PI");
            AddAlias(d, "RAND", "ALEATORIO");
            AddAlias(d, "RANDBETWEEN", "ALEATORIO.ENTRE");
            AddAlias(d, "SIGN", "SIGNO");
            AddAlias(d, "CEILING", "MULTIPLO.SUPERIOR");
            AddAlias(d, "FLOOR", "MULTIPLO.INFERIOR");
            AddAlias(d, "SUMIF", "SUMAR.SI");
            AddAlias(d, "COUNTIF", "CONTAR.SI");
            AddAlias(d, "AVERAGEIF", "PROMEDIO.SI");
            AddAlias(d, "SUMIFS", "SUMAR.SI.CONJUNTO");
            AddAlias(d, "COUNTIFS", "CONTAR.SI.CONJUNTO");
            AddAlias(d, "AVERAGEIFS", "PROMEDIO.SI.CONJUNTO");
            AddAlias(d, "MAXIFS", "MAX.SI.CONJUNTO");
            AddAlias(d, "MINIFS", "MIN.SI.CONJUNTO");
            AddAlias(d, "SUBTOTAL", "SUBTOTALES");
            AddAlias(d, "AGGREGATE", "AGREGAR");
            AddAlias(d, "SUMPRODUCT", "SUMAPRODUCTO");
            AddAlias(d, "IF", "SI");
            AddAlias(d, "IFS", "SI.CONJUNTO");
            AddAlias(d, "SWITCH", "CAMBIAR");
            AddAlias(d, "AND", "Y");
            AddAlias(d, "OR", "O");
            AddAlias(d, "XOR", "XO");
            AddAlias(d, "NOT", "NO");
            AddAlias(d, "IFERROR", "SI.ERROR");
            AddAlias(d, "IFNA", "SI.ND");
            AddAlias(d, "TRUE", "VERDADERO");
            AddAlias(d, "FALSE", "FALSO");
            AddAlias(d, "ISBLANK", "ESBLANCO");
            AddAlias(d, "ISNUMBER", "ESNUMERO");
            AddAlias(d, "ISTEXT", "ESTEXTO");
            AddAlias(d, "ISERROR", "ESERROR");
            AddAlias(d, "ISNA", "ESNOD");
            AddAlias(d, "ISERR", "ESERR");
            AddAlias(d, "ISEVEN", "ES.PAR");
            AddAlias(d, "ISODD", "ES.IMPAR");
            AddAlias(d, "NA", "NOD");
            AddAlias(d, "CONCATENATE", "CONCATENAR");
            AddAlias(d, "CONCAT");
            AddAlias(d, "TEXTJOIN", "UNIRTEXTO");
            AddAlias(d, "TEXTBEFORE", "TEXTOANTES");
            AddAlias(d, "TEXTAFTER", "TEXTODESPUES");
            AddAlias(d, "LEN", "LARGO");
            AddAlias(d, "LEFT", "IZQUIERDA");
            AddAlias(d, "RIGHT", "DERECHA");
            AddAlias(d, "MID", "EXTRAE");
            AddAlias(d, "UPPER", "MAYUSC", "MAYUSCULAS");
            AddAlias(d, "LOWER", "MINUSC", "MINUSCULAS");
            AddAlias(d, "TRIM", "ESPACIOS");
            AddAlias(d, "PROPER", "NOMPROPIO");
            AddAlias(d, "TEXT", "TEXTO");
            AddAlias(d, "FORMULATEXT", "FORMULATEXTO");
            AddAlias(d, "VALUE", "VALOR");
            AddAlias(d, "REPT", "REPETIR");
            AddAlias(d, "SUBSTITUTE", "SUSTITUIR");
            AddAlias(d, "FIND", "ENCONTRAR");
            AddAlias(d, "SEARCH", "HALLAR");
            AddAlias(d, "EXACT", "IGUAL");
            AddAlias(d, "CHAR", "CARACTER");
            AddAlias(d, "CODE", "CODIGO");
            AddAlias(d, "TODAY", "HOY");
            AddAlias(d, "NOW", "AHORA");
            AddAlias(d, "DATE", "FECHA");
            AddAlias(d, "TIME", "NSHORA");
            AddAlias(d, "TIMEVALUE", "HORA.NUMERO");
            AddAlias(d, "DAYS", "DIAS");
            AddAlias(d, "DATEDIF", "SIFECHA");
            AddAlias(d, "NETWORKDAYS", "DIAS.LAB");
            AddAlias(d, "WORKDAY", "DIA.LAB");
            AddAlias(d, "YEARFRAC", "FRAC.AÑO", "FRAC.ANO");
            AddAlias(d, "RANK", "JERARQUIA");
            AddAlias(d, "RANK.EQ", "JERARQUIA.EQ");
            AddAlias(d, "PERCENTILE", "PERCENTIL");
            AddAlias(d, "PERCENTILE.INC", "PERCENTIL.INC");
            AddAlias(d, "QUARTILE", "CUARTIL");
            AddAlias(d, "QUARTILE.INC", "CUARTIL.INC");
            AddAlias(d, "CEILING.MATH", "MULTIPLO.SUPERIOR.MAT");
            AddAlias(d, "FLOOR.MATH", "MULTIPLO.INFERIOR.MAT");
            AddAlias(d, "YEAR", "A\u00D1O", "ANO");
            AddAlias(d, "MONTH", "MES");
            AddAlias(d, "DAY", "DIA");
            AddAlias(d, "HOUR", "HORA");
            AddAlias(d, "MINUTE", "MINUTO");
            AddAlias(d, "SECOND", "SEGUNDO");
            AddAlias(d, "WEEKDAY", "DIASEM");
            AddAlias(d, "DATEVALUE", "FECHANUMERO");
            AddAlias(d, "EDATE", "FECHA.MES");
            AddAlias(d, "EOMONTH", "FIN.MES");
            AddAlias(d, "VLOOKUP", "BUSCARV");
            AddAlias(d, "HLOOKUP", "BUSCARH");
            AddAlias(d, "XLOOKUP", "BUSCARX");
            AddAlias(d, "XMATCH", "COINCIDIRX");
            AddAlias(d, "INDEX", "INDICE");
            AddAlias(d, "MATCH", "COINCIDIR");
            AddAlias(d, "CHOOSE", "ELEGIR");
            AddAlias(d, "ROW", "FILA");
            AddAlias(d, "COLUMN", "COLUMNA");
            AddAlias(d, "ROWS", "FILAS");
            AddAlias(d, "COLUMNS", "COLUMNAS");
            AddAlias(d, "FILTER", "FILTRAR");
            AddAlias(d, "SORT", "ORDENAR");
            AddAlias(d, "UNIQUE", "UNICOS", "UNICAS");
            AddAlias(d, "SEQUENCE", "SECUENCIA");
            AddAlias(d, "LET", "LET");
            AddAlias(d, "TAKE", "TOMAR");
            AddAlias(d, "DROP", "DESCARTAR");
            AddAlias(d, "HSTACK", "APILARH");
            AddAlias(d, "VSTACK", "APILARV");
            return d;
        }

        public static string Canonical(string name)
        {
            if (string.IsNullOrEmpty(name)) return null;
            string n = name.Trim().ToUpperInvariant();
            if (n.StartsWith("_XLFN.", StringComparison.Ordinal)) n = n.Substring(6);
            if (n.StartsWith("_XLWS.", StringComparison.Ordinal)) n = n.Substring(6);
            string c;
            return Aliases.TryGetValue(n, out c) ? c : null;
        }

        private static CellValue Arg(FunctionNode n, int i, EvalContext ctx)
        {
            if (i < 0 || i >= n.Args.Count) return CellValue.Empty;
            return n.Args[i].Eval(ctx);
        }

        private static bool WrongArgs(FunctionNode n, int min, int max)
        {
            return n.Args.Count < min || (max >= 0 && n.Args.Count > max);
        }

        private const long MaxRangeCells = 2000000;

        private static RangeRef ClampRange(RangeRef r)
        {
            int r1 = r.Row1, c1 = r.Col1, r2 = r.Row2, c2 = r.Col2;
            long area = (long)(r2 - r1 + 1) * (c2 - c1 + 1);
            if (area > 200000)
            {
                r2 = Math.Min(r2, Math.Max(r1, r.Sheet.UsedMaxRow));
                c2 = Math.Min(c2, Math.Max(c1, r.Sheet.UsedMaxCol));
                area = (long)(r2 - r1 + 1) * (c2 - c1 + 1);
            }
            if (area > MaxRangeCells)
                throw new InvalidOperationException("El rango de la fórmula es demasiado grande.");
            return new RangeRef(r.Sheet, r1, c1, r2, c2);
        }

        private static IEnumerable<CellValue> RangeValues(RangeRef range, EvalContext ctx)
        {
            RangeRef r = ClampRange(range);
            for (int row = r.Row1; row <= r.Row2; row++)
                for (int col = r.Col1; col <= r.Col2; col++)
                    yield return ctx.GetCell(r.Sheet, row, col);
        }

        private static bool CollectNumbers(FunctionNode n, EvalContext ctx, bool includeDirectText,
                                           out List<double> numbers, out CellValue error)
        {
            numbers = new List<double>();
            error = CellValue.Empty;
            for (int i = 0; i < n.Args.Count; i++)
            {
                CellValue v = n.Args[i].Eval(ctx);
                if (v.Kind == ValueKind.Range)
                {
                    foreach (CellValue cv in RangeValues(v.Range, ctx))
                    {
                        if (cv.IsError) { error = cv; return false; }
                        if (cv.Kind == ValueKind.Number) numbers.Add(cv.Number);
                    }
                }
                else
                {
                    if (v.IsError) { error = v; return false; }
                    if (v.Kind == ValueKind.Number) numbers.Add(v.Number);
                    else if (v.Kind == ValueKind.Bool) numbers.Add(v.Bool ? 1.0 : 0.0);
                    else if (v.Kind == ValueKind.Text && includeDirectText)
                    {
                        double d;
                        if (NumberParser.TryParse(v.Text, out d)) numbers.Add(d);
                    }
                }
            }
            return true;
        }

        private static bool GetNumber(FunctionNode n, int index, EvalContext ctx, out double value, out CellValue error)
        {
            CellValue v = Arg(n, index, ctx);
            return EvalUtil.TryNumber(v, ctx, out value, out error);
        }

        private static bool GetText(FunctionNode n, int index, EvalContext ctx, out string value, out CellValue error)
        {
            CellValue v = Arg(n, index, ctx);
            value = EvalUtil.ToText(v, ctx, out error);
            return !error.IsError;
        }

        private static bool GetBool(FunctionNode n, int index, EvalContext ctx, out bool value, out CellValue error)
        {
            CellValue v = Arg(n, index, ctx);
            return EvalUtil.TryBool(v, ctx, out value, out error);
        }

        internal static CellValue Invoke(FunctionNode n, EvalContext ctx)
        {
            string name = Canonical(n.Name);
                        if (name == null) return CellValue.FromError(Errors.Name);

            switch (name)
            {
                case "IF": return FnIf(n, ctx);
                case "IFS": return FnIfs(n, ctx);
                case "SWITCH": return FnSwitch(n, ctx);
                case "FILTER": return FnFilter(n, ctx);
                case "SORT": return FnSort(n, ctx);
                case "UNIQUE": return FnUnique(n, ctx);
                case "SEQUENCE": return FnSequence(n, ctx);
                case "LET": return FnLet(n, ctx);
                case "TAKE": return FnTakeDrop(n, ctx, true);
                case "DROP": return FnTakeDrop(n, ctx, false);
                case "HSTACK": return FnStack(n, ctx, false);
                case "VSTACK": return FnStack(n, ctx, true);
                case "IFERROR": return FnIfError(n, ctx, false);
                case "IFNA": return FnIfError(n, ctx, true);
                case "AND": return FnAndOr(n, ctx, true);
                case "OR": return FnAndOr(n, ctx, false);
                case "XOR": return FnXor(n, ctx);
                case "NOT":
                    {
                        bool b; CellValue err = CellValue.Empty;
                        if (WrongArgs(n, 1, 1) || !GetBool(n, 0, ctx, out b, out err)) return err.IsError ? err : CellValue.FromError(Errors.Value);
                        return CellValue.FromBool(!b);
                    }
                case "TRUE": return CellValue.FromBool(true);
                case "FALSE": return CellValue.FromBool(false);
                case "ISBLANK":
                    {
                        if (WrongArgs(n, 1, 1)) return CellValue.FromError(Errors.Value);
                        CellValue v = Arg(n, 0, ctx);
                        return CellValue.FromBool(v.Kind == ValueKind.Empty);
                    }
                case "ISNUMBER":
                    {
                        if (WrongArgs(n, 1, 1)) return CellValue.FromError(Errors.Value);
                        CellValue v = Arg(n, 0, ctx);
                        v = EvalUtil.Deref(v, ctx);
                        return CellValue.FromBool(v.Kind == ValueKind.Number);
                    }
                case "ISTEXT":
                    {
                        if (WrongArgs(n, 1, 1)) return CellValue.FromError(Errors.Value);
                        CellValue v = Arg(n, 0, ctx);
                        v = EvalUtil.Deref(v, ctx);
                        return CellValue.FromBool(v.Kind == ValueKind.Text);
                    }
                case "ISNA":
                    {
                        if (WrongArgs(n, 1, 1)) return CellValue.FromError(Errors.Value);
                        CellValue v = Arg(n, 0, ctx);
                        return CellValue.FromBool(v.IsError && v.Error == Errors.NA);
                    }
                case "ISERR":
                    {
                        if (WrongArgs(n, 1, 1)) return CellValue.FromError(Errors.Value);
                        CellValue v = Arg(n, 0, ctx);
                        return CellValue.FromBool(v.IsError && v.Error != Errors.NA);
                    }
                case "ISERROR":
                    {
                        if (WrongArgs(n, 1, 1)) return CellValue.FromError(Errors.Value);
                        CellValue v = Arg(n, 0, ctx);
                        return CellValue.FromBool(v.IsError);
                    }
                case "ISEVEN":
                case "ISODD":
                    {
                        double d; CellValue err = CellValue.Empty;
                        if (WrongArgs(n, 1, 1) || !GetNumber(n, 0, ctx, out d, out err)) return err.IsError ? err : CellValue.FromError(Errors.Value);
                        bool even = Math.Abs(Math.Truncate(d) % 2) < 0.5;
                        return CellValue.FromBool(name == "ISEVEN" ? even : !even);
                    }
                case "NA": return CellValue.FromError(Errors.NA);
                case "CHOOSE": return FnChoose(n, ctx);
                case "ROW":
                case "COLUMN": return FnRowColumn(n, ctx, name == "ROW");
                case "ROWS":
                case "COLUMNS": return FnRowsColumns(n, ctx, name == "ROWS");
                case "SUM": return FnSum(n, ctx);
                case "AVERAGE": return FnAverage(n, ctx);
                case "MIN":
                case "MAX": return FnMinMax(n, ctx, name == "MAX");
                case "COUNT": return FnCount(n, ctx, false);
                case "COUNTA": return FnCount(n, ctx, true);
                case "COUNTBLANK": return FnCountBlank(n, ctx);
                case "MEDIAN": return FnMedian(n, ctx);
                case "STDEV":
                case "STDEVP":
                case "VAR":
                case "VARP": return FnVariance(n, ctx, name);
                case "ROUND":
                case "ROUNDUP":
                case "ROUNDDOWN": return FnRound(n, ctx, name);
                case "FIXED": return FnFixed(n, ctx);
                case "MROUND": return FnMRound(n, ctx);
                case "COMBIN": return FnCombin(n, ctx);
                case "EVEN":
                case "ODD": return FnEvenOdd(n, ctx, name == "EVEN");
                case "INT": return FnUnaryMath(n, ctx, "INT");
                case "TRUNC": return FnTrunc(n, ctx);
                case "ABS": return FnUnaryMath(n, ctx, "ABS");
                case "SIGN": return FnUnaryMath(n, ctx, "SIGN");
                case "SQRT": return FnUnaryMath(n, ctx, "SQRT");
                case "EXP": return FnUnaryMath(n, ctx, "EXP");
                case "LN": return FnUnaryMath(n, ctx, "LN");
                case "LOG10": return FnUnaryMath(n, ctx, "LOG10");
                case "LOG": return FnLog(n, ctx);
                case "MOD": return FnMod(n, ctx);
                case "POWER": return FnPower(n, ctx);
                case "PI": return CellValue.FromNumber(Math.PI);
                case "RAND": return CellValue.FromNumber(Rng.NextDouble());
                case "RANDBETWEEN": return FnRandBetween(n, ctx);
                case "CEILING":
                case "FLOOR": return FnCeilingFloor(n, ctx, name == "CEILING");
                case "SUMPRODUCT": return FnSumProduct(n, ctx);
                case "SUMIF": return FnConditionalAggregate(n, ctx, 0);
                case "COUNTIF": return FnConditionalAggregate(n, ctx, 1);
                case "AVERAGEIF": return FnConditionalAggregate(n, ctx, 2);
                case "SUMIFS": return FnMultiAggregate(n, ctx, 0);
                case "COUNTIFS": return FnMultiAggregate(n, ctx, 1);
                case "AVERAGEIFS": return FnMultiAggregate(n, ctx, 2);
                case "MAXIFS": return FnMultiAggregate(n, ctx, 3);
                case "MINIFS": return FnMultiAggregate(n, ctx, 4);
                case "SUBTOTAL": return FnSubtotal(n, ctx);
                case "AGGREGATE": return FnAggregate(n, ctx);
                case "CONCATENATE":
                case "CONCAT": return FnConcat(n, ctx);
                case "TEXTJOIN": return FnTextJoin(n, ctx);
                case "TEXTBEFORE": return FnTextBeforeAfter(n, ctx, true);
                case "TEXTAFTER": return FnTextBeforeAfter(n, ctx, false);
                case "LEN": return FnTextLength(n, ctx);
                case "LEFT":
                case "RIGHT": return FnLeftRight(n, ctx, name == "LEFT");
                case "MID": return FnMid(n, ctx);
                case "UPPER":
                case "LOWER": return FnUpperLower(n, ctx, name == "UPPER");
                case "TRIM": return FnTrim(n, ctx);
                case "PROPER": return FnProper(n, ctx);
                case "TEXT": return FnText(n, ctx);
                case "FORMULATEXT": return FnFormulaText(n, ctx);
                case "VALUE": return FnValue(n, ctx);
                case "REPT": return FnRept(n, ctx);
                case "SUBSTITUTE": return FnSubstitute(n, ctx);
                case "FIND": return FnFind(n, ctx, true);
                case "SEARCH": return FnFind(n, ctx, false);
                case "EXACT": return FnExact(n, ctx);
                case "CHAR": return FnChar(n, ctx);
                case "CODE": return FnCode(n, ctx);
                case "TODAY": return CellValue.FromNumber(DateTime.Today.ToOADate());
                case "NOW": return CellValue.FromNumber(DateTime.Now.ToOADate());
                case "DATE": return FnDate(n, ctx);
                case "TIME": return FnTime(n, ctx);
                case "TIMEVALUE": return FnTimeValue(n, ctx);
                case "DAYS": return FnDays(n, ctx);
                case "DATEDIF": return FnDateDif(n, ctx);
                case "NETWORKDAYS": return FnNetworkDays(n, ctx);
                case "WORKDAY": return FnWorkday(n, ctx);
                case "YEARFRAC": return FnYearFrac(n, ctx);
                case "RANK":
                case "RANK.EQ": return FnRank(n, ctx);
                case "PERCENTILE":
                case "PERCENTILE.INC": return FnPercentile(n, ctx);
                case "QUARTILE":
                case "QUARTILE.INC": return FnQuartile(n, ctx);
                case "CEILING.MATH":
                case "FLOOR.MATH": return FnCeilingMath(n, ctx, name == "CEILING.MATH");
                case "YEAR":
                case "MONTH":
                case "DAY":
                case "HOUR":
                case "MINUTE":
                case "SECOND": return FnDatePart(n, ctx, name);
                case "WEEKDAY": return FnWeekday(n, ctx);
                case "DATEVALUE": return FnDateValue(n, ctx);
                case "EDATE": return FnEDate(n, ctx);
                case "EOMONTH": return FnEOMonth(n, ctx);
                case "VLOOKUP": return FnVLookup(n, ctx, true);
                case "HLOOKUP": return FnVLookup(n, ctx, false);
                case "XLOOKUP": return FnXLookup(n, ctx);
                case "XMATCH": return FnXMatch(n, ctx);
                case "INDEX": return FnIndex(n, ctx);
                case "MATCH": return FnMatch(n, ctx);
                default: return CellValue.FromError(Errors.Name);
            }
        }

        private static CellValue FnIf(FunctionNode n, EvalContext ctx)
        {
            if (WrongArgs(n, 2, 3)) return CellValue.FromError(Errors.Value);
            bool cond; CellValue err = CellValue.Empty;
            if (!GetBool(n, 0, ctx, out cond, out err)) return err;
            if (cond) return Arg(n, 1, ctx);
            if (n.Args.Count >= 3) return Arg(n, 2, ctx);
            return CellValue.FromBool(false);
        }

        private static CellValue FnIfs(FunctionNode n, EvalContext ctx)
        {
            if (n.Args.Count < 2 || (n.Args.Count % 2) != 0) return CellValue.FromError(Errors.Value);
            for (int i = 0; i < n.Args.Count; i += 2)
            {
                bool cond; CellValue err = CellValue.Empty;
                if (!GetBool(n, i, ctx, out cond, out err)) return err;
                if (cond) return Arg(n, i + 1, ctx);
            }
            return CellValue.FromError(Errors.NA);
        }

        private static CellValue FnSwitch(FunctionNode n, EvalContext ctx)
        {
            if (n.Args.Count < 3) return CellValue.FromError(Errors.Value);
            CellValue expr = Arg(n, 0, ctx);
            if (expr.IsError) return expr;
            int i = 1;
            while (i + 1 < n.Args.Count)
            {
                CellValue candidate = Arg(n, i, ctx);
                if (!candidate.IsError && ValuesEqual(expr, candidate)) return Arg(n, i + 1, ctx);
                i += 2;
            }
            if (i < n.Args.Count) return Arg(n, i, ctx);
            return CellValue.FromError(Errors.NA);
        }

        private static CellValue FnIfError(FunctionNode n, EvalContext ctx, bool onlyNA)
        {
            if (WrongArgs(n, 2, 2)) return CellValue.FromError(Errors.Value);
            CellValue v = Arg(n, 0, ctx);
            if (v.IsError && (!onlyNA || v.Error == Errors.NA))
                return Arg(n, 1, ctx);
            return v;
        }

        private static CellValue FnAndOr(FunctionNode n, EvalContext ctx, bool isAnd)
        {
            if (n.Args.Count == 0) return CellValue.FromError(Errors.Value);
            bool any = false;
            bool result = isAnd;
            for (int i = 0; i < n.Args.Count; i++)
            {
                CellValue v = n.Args[i].Eval(ctx);
                if (v.Kind == ValueKind.Range)
                {
                    foreach (CellValue cv in RangeValues(v.Range, ctx))
                    {
                        if (cv.IsError) return cv;
                        if (cv.Kind == ValueKind.Bool) { any = true; result = isAnd ? (result && cv.Bool) : (result || cv.Bool); }
                        else if (cv.Kind == ValueKind.Number) { any = true; result = isAnd ? (result && cv.Number != 0) : (result || cv.Number != 0); }
                    }
                }
                else
                {
                    if (v.IsError) return v;
                    if (v.Kind == ValueKind.Empty) continue;
                    bool b; CellValue err = CellValue.Empty;
                    if (!EvalUtil.TryBool(v, ctx, out b, out err)) return err;
                    any = true;
                    result = isAnd ? (result && b) : (result || b);
                }
            }
            if (!any) return CellValue.FromBool(false);
            return CellValue.FromBool(result);
        }

        private static CellValue FnXor(FunctionNode n, EvalContext ctx)
        {
            if (n.Args.Count == 0) return CellValue.FromError(Errors.Value);
            bool any = false;
            bool result = false;
            for (int i = 0; i < n.Args.Count; i++)
            {
                CellValue v = n.Args[i].Eval(ctx);
                if (v.Kind == ValueKind.Range)
                {
                    foreach (CellValue cv in RangeValues(v.Range, ctx))
                    {
                        if (cv.IsError) return cv;
                        if (cv.Kind == ValueKind.Bool) { any = true; result ^= cv.Bool; }
                        else if (cv.Kind == ValueKind.Number) { any = true; result ^= cv.Number != 0; }
                    }
                }
                else
                {
                    if (v.IsError) return v;
                    if (v.Kind == ValueKind.Empty) continue;
                    bool b;
                    CellValue err = CellValue.Empty;
                    if (!EvalUtil.TryBool(v, ctx, out b, out err)) return err;
                    any = true;
                    result ^= b;
                }
            }
            return CellValue.FromBool(any && result);
        }

        private sealed class Matrix
        {
            public int Rows;
            public int Cols;
            public List<CellValue> Cells = new List<CellValue>();
            public CellValue At(int row, int col)
            {
                int i = row * Cols + col;
                return i >= 0 && i < Cells.Count ? Cells[i] : CellValue.Empty;
            }
        }

        private static Matrix ToMatrix(CellValue value, EvalContext ctx)
        {
            Matrix m = new Matrix();
            if (value.Kind == ValueKind.Range)
            {
                RangeRef r = ClampRange(value.Range);
                m.Rows = r.Rows;
                m.Cols = r.Cols;
                foreach (CellValue v in RangeValues(r, ctx)) m.Cells.Add(v);
                return m;
            }
            if (value.Kind == ValueKind.Array)
            {
                m.Rows = value.ArrayRows > 0 ? value.ArrayRows : 1;
                m.Cols = value.ArrayCols > 0 ? value.ArrayCols : Math.Max(1, value.Values == null ? 0 : value.Values.Count);
                if (value.Values != null) m.Cells.AddRange(value.Values);
                return m;
            }
            m.Rows = 1; m.Cols = 1;
            m.Cells.Add(value);
            return m;
        }

        private static CellValue FromMatrix(Matrix m)
        {
            return CellValue.FromArray(m.Cells, m.Rows, m.Cols);
        }

        private static CellValue FnFilter(FunctionNode n, EvalContext ctx)
        {
            if (WrongArgs(n, 2, 3)) return CellValue.FromError(Errors.Value);
            CellValue source = Arg(n, 0, ctx);
            if (source.IsError) return source;
            Matrix m = ToMatrix(source, ctx);
            CellValue maskValue = Arg(n, 1, ctx);
            if (maskValue.IsError) return maskValue;
            Matrix mask = ToMatrix(maskValue, ctx);
            bool filterRows = mask.Rows == m.Rows || mask.Cells.Count == m.Rows;
            bool filterCols = !filterRows && mask.Rows == 1 && mask.Cols == m.Cols;
            if (!filterRows && !filterCols) return CellValue.FromError(Errors.Value);
            List<int> selected = new List<int>();
            int count = filterRows ? m.Rows : m.Cols;
            for (int i = 0; i < count; i++)
            {
                CellValue v = filterRows ? mask.At(i, 0) : mask.At(0, i);
                bool b; CellValue err = CellValue.Empty;
                if (!EvalUtil.TryBool(v, ctx, out b, out err)) return err;
                if (b) selected.Add(i);
            }
            if (selected.Count == 0)
            {
                if (n.Args.Count >= 3) return Arg(n, 2, ctx);
                return CellValue.FromError(Errors.NA);
            }
            Matrix result = new Matrix();
            if (filterRows)
            {
                result.Rows = selected.Count; result.Cols = m.Cols;
                foreach (int r in selected)
                    for (int c = 0; c < m.Cols; c++) result.Cells.Add(m.At(r, c));
            }
            else
            {
                result.Rows = m.Rows; result.Cols = selected.Count;
                for (int r = 0; r < m.Rows; r++)
                    foreach (int c in selected) result.Cells.Add(m.At(r, c));
            }
            return FromMatrix(result);
        }

        private static CellValue FnSort(FunctionNode n, EvalContext ctx)
        {
            if (WrongArgs(n, 1, 4)) return CellValue.FromError(Errors.Value);
            CellValue source = Arg(n, 0, ctx);
            if (source.IsError) return source;
            Matrix m = ToMatrix(source, ctx);
            double indexD = 1, orderD = 1;
            CellValue err = CellValue.Empty;
            if (n.Args.Count >= 2 && !GetNumber(n, 1, ctx, out indexD, out err)) return err;
            if (n.Args.Count >= 3 && !GetNumber(n, 2, ctx, out orderD, out err)) return err;
            int index = (int)Math.Truncate(indexD);
            int order = orderD < 0 ? -1 : 1;
            if (index < 1 || index > m.Cols) return CellValue.FromError(Errors.Value);
            List<int> rows = new List<int>();
            for (int r = 0; r < m.Rows; r++) rows.Add(r);
            rows.Sort(delegate(int a, int b)
            {
                int cmp = CompareCells(m.At(a, index - 1), m.At(b, index - 1), ctx);
                return order * cmp;
            });
            Matrix result = new Matrix { Rows = m.Rows, Cols = m.Cols };
            foreach (int r in rows)
                for (int c = 0; c < m.Cols; c++) result.Cells.Add(m.At(r, c));
            return FromMatrix(result);
        }

        private static int CompareCells(CellValue a, CellValue b, EvalContext ctx)
        {
            if (a.Kind == ValueKind.Number && b.Kind == ValueKind.Number)
                return a.Number.CompareTo(b.Number);
            bool ba, bb; CellValue err = CellValue.Empty;
            if (EvalUtil.TryBool(a, ctx, out ba, out err) && EvalUtil.TryBool(b, ctx, out bb, out err))
                return ba.CompareTo(bb);
            return string.Compare(a.ToDisplayString(), b.ToDisplayString(),
                StringComparison.CurrentCultureIgnoreCase);
        }

        private static CellValue FnUnique(FunctionNode n, EvalContext ctx)
        {
            if (WrongArgs(n, 1, 2)) return CellValue.FromError(Errors.Value);
            CellValue source = Arg(n, 0, ctx);
            if (source.IsError) return source;
            Matrix m = ToMatrix(source, ctx);
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            Matrix result = new Matrix { Rows = 0, Cols = m.Cols };
            for (int r = 0; r < m.Rows; r++)
            {
                StringBuilder key = new StringBuilder();
                for (int c = 0; c < m.Cols; c++)
                {
                    key.Append(m.At(r, c).Kind).Append(':')
                        .Append(m.At(r, c).ToDisplayString()).Append('\u001f');
                }
                if (!seen.Add(key.ToString())) continue;
                result.Rows++;
                for (int c = 0; c < m.Cols; c++) result.Cells.Add(m.At(r, c));
            }
            return FromMatrix(result);
        }

        private static CellValue FnSequence(FunctionNode n, EvalContext ctx)
        {
            if (WrongArgs(n, 1, 4)) return CellValue.FromError(Errors.Value);
            double rowsD, colsD = 1, startD = 1, stepD = 1;
            CellValue err = CellValue.Empty;
            if (!GetNumber(n, 0, ctx, out rowsD, out err)) return err;
            if (n.Args.Count >= 2 && !GetNumber(n, 1, ctx, out colsD, out err)) return err;
            if (n.Args.Count >= 3 && !GetNumber(n, 2, ctx, out startD, out err)) return err;
            if (n.Args.Count >= 4 && !GetNumber(n, 3, ctx, out stepD, out err)) return err;
            int rows = (int)Math.Truncate(rowsD);
            int cols = (int)Math.Truncate(colsD);
            if (rows < 1 || cols < 1 || (long)rows * cols > 100000) return CellValue.FromError(Errors.Value);
            Matrix m = new Matrix { Rows = rows, Cols = cols };
            double value = startD;
            for (int r = 0; r < rows; r++)
                for (int c = 0; c < cols; c++)
                {
                    m.Cells.Add(CellValue.FromNumber(value));
                    value += stepD;
                }
            return FromMatrix(m);
        }

        private static CellValue FnLet(FunctionNode n, EvalContext ctx)
        {
            if (n.Args.Count < 3 || n.Args.Count % 2 == 0) return CellValue.FromError(Errors.Value);
            int pairs = (n.Args.Count - 1) / 2;
            Dictionary<string, CellValue> previous = new Dictionary<string, CellValue>(StringComparer.OrdinalIgnoreCase);
            List<string> assigned = new List<string>();
            for (int i = 0; i < pairs; i++)
            {
                NameNode nameNode = n.Args[i * 2] as NameNode;
                if (nameNode == null) return CellValue.FromError(Errors.Name);
                string name = nameNode.Name;
                CellValue value = n.Args[i * 2 + 1].Eval(ctx);
                if (value.IsError) return value;
                if (!previous.ContainsKey(name)) previous[name] = ctx.Variables.ContainsKey(name)
                    ? ctx.Variables[name] : CellValue.Empty;
                ctx.Variables[name] = value;
                assigned.Add(name);
            }
            CellValue result;
            try { result = n.Args[n.Args.Count - 1].Eval(ctx); }
            finally
            {
                foreach (string name in assigned)
                    if (previous.ContainsKey(name)) ctx.Variables[name] = previous[name];
            }
            return result;
        }

        private static CellValue FnTakeDrop(FunctionNode n, EvalContext ctx, bool take)
        {
            if (WrongArgs(n, 2, 3)) return CellValue.FromError(Errors.Value);
            CellValue source = Arg(n, 0, ctx);
            if (source.IsError) return source;
            Matrix m = ToMatrix(source, ctx);
            double countD; CellValue err = CellValue.Empty;
            if (!GetNumber(n, 1, ctx, out countD, out err)) return err;
            int count = (int)Math.Truncate(countD);
            int start = 0, end = m.Rows;
            if (take)
            {
                if (count >= 0) end = Math.Min(m.Rows, count);
                else start = Math.Max(0, m.Rows + count);
            }
            else
            {
                if (count >= 0) start = Math.Min(m.Rows, count);
                else end = Math.Max(start, m.Rows + count);
            }
            if (end < start) end = start;
            Matrix result = new Matrix { Rows = end - start, Cols = m.Cols };
            for (int r = start; r < end; r++)
                for (int c = 0; c < m.Cols; c++) result.Cells.Add(m.At(r, c));
            return FromMatrix(result);
        }

        private static CellValue FnStack(FunctionNode n, EvalContext ctx, bool vertical)
        {
                        if (n.Args.Count == 0) return CellValue.FromError(Errors.Value);
            List<Matrix> all = new List<Matrix>();
            foreach (FormulaNode arg in n.Args)
            {
                CellValue v = arg.Eval(ctx);
                if (v.IsError) return v;
                all.Add(ToMatrix(v, ctx));
            }
            int rows = 0, cols = 0;
            if (vertical)
            {
                foreach (Matrix m in all) { rows += m.Rows; cols = Math.Max(cols, m.Cols); }
            }
            else
            {
                foreach (Matrix m in all) { cols += m.Cols; rows = Math.Max(rows, m.Rows); }
            }
            if ((long)rows * cols > 100000) return CellValue.FromError(Errors.Value);
            Matrix result = new Matrix { Rows = rows, Cols = cols };
            for (int i = 0; i < result.Rows * result.Cols; i++) result.Cells.Add(CellValue.Empty);
            int rowOffset = 0, colOffset = 0;
            foreach (Matrix m in all)
            {
                int rr = vertical ? rowOffset : 0;
                int cc = vertical ? 0 : colOffset;
                for (int r = 0; r < m.Rows; r++)
                    for (int c = 0; c < m.Cols; c++)
                        result.Cells[(rr + r) * cols + (cc + c)] = m.At(r, c);
                if (vertical) rowOffset += m.Rows; else colOffset += m.Cols;
            }
            return FromMatrix(result);
        }

        private static CellValue FnChoose(FunctionNode n, EvalContext ctx)
        {
            if (n.Args.Count < 2) return CellValue.FromError(Errors.Value);
            double d; CellValue err = CellValue.Empty;
            if (!GetNumber(n, 0, ctx, out d, out err)) return err;
            int idx = (int)Math.Truncate(d);
            if (idx < 1 || idx >= n.Args.Count) return CellValue.FromError(Errors.Value);
            return Arg(n, idx, ctx);
        }

        private static CellValue FnRowColumn(FunctionNode n, EvalContext ctx, bool row)
        {
            if (n.Args.Count == 0)
                return CellValue.FromNumber(row ? ctx.CurrentRow + 1 : ctx.CurrentCol + 1);
            FunctionNode fn = n.Args[0] as FunctionNode;
            RefNode rn = n.Args[0] as RefNode;
            RangeNode rng = n.Args[0] as RangeNode;
            if (rn != null) return CellValue.FromNumber(row ? rn.Row + 1 : rn.Col + 1);
            if (rng != null) return CellValue.FromNumber(row ? rng.R1 + 1 : rng.C1 + 1);
            CellValue v = n.Args[0].Eval(ctx);
            if (v.Kind == ValueKind.Range)
                return CellValue.FromNumber(row ? v.Range.Row1 + 1 : v.Range.Col1 + 1);
            return CellValue.FromError(Errors.Value);
        }

        private static CellValue FnRowsColumns(FunctionNode n, EvalContext ctx, bool rows)
        {
            if (WrongArgs(n, 1, 1)) return CellValue.FromError(Errors.Value);
            CellValue v = Arg(n, 0, ctx);
            if (v.Kind == ValueKind.Range)
                return CellValue.FromNumber(rows ? v.Range.Rows : v.Range.Cols);
            return CellValue.FromNumber(1);
        }

        private static CellValue FnSum(FunctionNode n, EvalContext ctx)
        {
            List<double> nums; CellValue err = CellValue.Empty;
            if (!CollectNumbers(n, ctx, true, out nums, out err)) return err;
            double sum = 0;
            foreach (double d in nums) sum += d;
            return CellValue.FromNumber(sum);
        }

        private static CellValue FnAverage(FunctionNode n, EvalContext ctx)
        {
            List<double> nums; CellValue err = CellValue.Empty;
            if (!CollectNumbers(n, ctx, true, out nums, out err)) return err;
            if (nums.Count == 0) return CellValue.FromError(Errors.Div0);
            double sum = 0;
            foreach (double d in nums) sum += d;
            return CellValue.FromNumber(sum / nums.Count);
        }

        private static CellValue FnMinMax(FunctionNode n, EvalContext ctx, bool max)
        {
            List<double> nums; CellValue err = CellValue.Empty;
            if (!CollectNumbers(n, ctx, true, out nums, out err)) return err;
            if (nums.Count == 0) return CellValue.FromNumber(0);
            double v = nums[0];
            for (int i = 1; i < nums.Count; i++)
            {
                if (max ? nums[i] > v : nums[i] < v) v = nums[i];
            }
            return CellValue.FromNumber(v);
        }

        private static CellValue FnCount(FunctionNode n, EvalContext ctx, bool countAll)
        {
            int count = 0;
            for (int i = 0; i < n.Args.Count; i++)
            {
                CellValue v = n.Args[i].Eval(ctx);
                if (v.Kind == ValueKind.Range)
                {
                    foreach (CellValue cv in RangeValues(v.Range, ctx))
                    {
                        // COUNT/COUNTA no propagan errores: COUNTA cuenta
                        // errores como celdas no vacías y COUNT los ignora.
                        if (countAll)
                        {
                            if (cv.Kind != ValueKind.Empty) count++;
                        }
                        else if (cv.Kind == ValueKind.Number) count++;
                    }
                }
                else
                {
                    if (countAll)
                    {
                        if (v.Kind != ValueKind.Empty) count++;
                    }
                    else
                    {
                        if (v.Kind == ValueKind.Number) count++;
                        else if (v.Kind == ValueKind.Text)
                        {
                            double d;
                            if (NumberParser.TryParse(v.Text, out d) &&
                                !double.IsNaN(d) && !double.IsInfinity(d)) count++;
                        }
                    }
                }
            }
            return CellValue.FromNumber(count);
        }

        private static CellValue FnCountBlank(FunctionNode n, EvalContext ctx)
        {
            if (WrongArgs(n, 1, 1)) return CellValue.FromError(Errors.Value);
            CellValue v = Arg(n, 0, ctx);
            if (v.Kind != ValueKind.Range)
                return CellValue.FromNumber(v.Kind == ValueKind.Empty ||
                    (v.Kind == ValueKind.Text && (v.Text ?? "").Length == 0) ? 1 : 0);
            RangeRef r = ClampRange(v.Range);
            int count = 0;
            for (int row = r.Row1; row <= r.Row2; row++)
                for (int col = r.Col1; col <= r.Col2; col++)
                {
                    CellValue cv = ctx.GetCell(r.Sheet, row, col);
                    if (cv.Kind == ValueKind.Empty ||
                        (cv.Kind == ValueKind.Text && (cv.Text ?? "").Length == 0)) count++;
                }
            return CellValue.FromNumber(count);
        }

        private static CellValue FnMedian(FunctionNode n, EvalContext ctx)
        {
            List<double> nums; CellValue err = CellValue.Empty;
            if (!CollectNumbers(n, ctx, true, out nums, out err)) return err;
            if (nums.Count == 0) return CellValue.FromError(Errors.Num);
            nums.Sort();
            int mid = nums.Count / 2;
            if ((nums.Count % 2) == 1) return CellValue.FromNumber(nums[mid]);
            return CellValue.FromNumber((nums[mid - 1] + nums[mid]) / 2.0);
        }

        private static CellValue FnVariance(FunctionNode n, EvalContext ctx, string kind)
        {
            List<double> nums; CellValue err = CellValue.Empty;
            if (!CollectNumbers(n, ctx, true, out nums, out err)) return err;
            bool population = kind == "STDEVP" || kind == "VARP";
            if (nums.Count < (population ? 1 : 2)) return CellValue.FromError(Errors.Div0);
            double mean = 0;
            foreach (double d in nums) mean += d;
            mean /= nums.Count;
            double sum = 0;
            foreach (double d in nums) { double diff = d - mean; sum += diff * diff; }
            double variance = sum / (population ? nums.Count : nums.Count - 1);
            if (kind.StartsWith("STDEV", StringComparison.Ordinal)) return CellValue.FromNumber(Math.Sqrt(variance));
            return CellValue.FromNumber(variance);
        }

        private static CellValue FnRound(FunctionNode n, EvalContext ctx, string kind)
        {
            if (WrongArgs(n, 1, 2)) return CellValue.FromError(Errors.Value);
            double d, digits = 0; CellValue err = CellValue.Empty;
            if (!GetNumber(n, 0, ctx, out d, out err)) return err;
            if (n.Args.Count >= 2 && !GetNumber(n, 1, ctx, out digits, out err)) return err;
            int dec = (int)Math.Truncate(digits);
            dec = Math.Max(-15, Math.Min(15, dec));
            double factor = Math.Pow(10, dec);
            double scaled = d * factor;
            double r;
            if (kind == "ROUND") r = Math.Round(scaled, 0, MidpointRounding.AwayFromZero);
            else if (kind == "ROUNDUP") r = scaled >= 0 ? Math.Ceiling(scaled) : Math.Floor(scaled);
            else r = scaled >= 0 ? Math.Floor(scaled) : Math.Ceiling(scaled);
            double result = r / factor;
            if (double.IsNaN(result) || double.IsInfinity(result)) return CellValue.FromError(Errors.Num);
            return CellValue.FromNumber(result);
        }

        private static CellValue FnMRound(FunctionNode n, EvalContext ctx)
        {
            double d, multiple; CellValue err = CellValue.Empty;
            if (WrongArgs(n, 2, 2) || !GetNumber(n, 0, ctx, out d, out err) || !GetNumber(n, 1, ctx, out multiple, out err))
                return err.IsError ? err : CellValue.FromError(Errors.Value);
            if (multiple == 0) return CellValue.FromError(Errors.Div0);
            if ((d > 0) != (multiple > 0)) return CellValue.FromError(Errors.Num);
            return CellValue.FromNumber(Math.Round(d / multiple, 0, MidpointRounding.AwayFromZero) * multiple);
        }

        private static CellValue FnCombin(FunctionNode n, EvalContext ctx)
        {
            double nd, kd; CellValue err = CellValue.Empty;
            if (WrongArgs(n, 2, 2) || !GetNumber(n, 0, ctx, out nd, out err) ||
                !GetNumber(n, 1, ctx, out kd, out err)) return err;
            long n0 = (long)Math.Truncate(nd), k0 = (long)Math.Truncate(kd);
            if (n0 < 0 || k0 < 0 || k0 > n0 || n0 > 1000) return CellValue.FromError(Errors.Num);
            double result = 1;
            for (long i = 1; i <= k0; i++)
            {
                result = result * (n0 - k0 + i) / i;
                if (double.IsNaN(result) || double.IsInfinity(result)) return CellValue.FromError(Errors.Num);
            }
            return CellValue.FromNumber(Math.Round(result, 0, MidpointRounding.AwayFromZero));
        }

        private static CellValue FnEvenOdd(FunctionNode n, EvalContext ctx, bool even)
        {
            double d; CellValue err = CellValue.Empty;
            if (WrongArgs(n, 1, 1) || !GetNumber(n, 0, ctx, out d, out err)) return err;
            double a = Math.Abs(d);
            double r = Math.Ceiling(a / 2.0) * 2.0;
            if (even)
            {
                if (Math.Abs(a - Math.Truncate(a)) < 1e-12 && ((long)Math.Truncate(a) % 2) == 0) r = a;
            }
            else
            {
                r = Math.Ceiling((a - 1) / 2.0) * 2.0 + 1.0;
                if (Math.Abs(a - Math.Truncate(a)) < 1e-12 && ((long)Math.Truncate(a) % 2) == 1) r = a;
            }
            return CellValue.FromNumber(d < 0 ? -r : r);
        }

        private static CellValue FnFixed(FunctionNode n, EvalContext ctx)
        {
            double d, digits = 2;
            CellValue err = CellValue.Empty;
            if (WrongArgs(n, 1, 3) || !GetNumber(n, 0, ctx, out d, out err)) return err;
            if (n.Args.Count >= 2 && !GetNumber(n, 1, ctx, out digits, out err)) return err;
            bool noCommas = false;
            if (n.Args.Count >= 3 && !GetBool(n, 2, ctx, out noCommas, out err)) return err;
            int dec = Math.Max(-15, Math.Min(15, (int)Math.Truncate(digits)));
            double rounded = d;
            if (dec >= 0) rounded = Math.Round(d, dec, MidpointRounding.AwayFromZero);
            else
            {
                double factor = Math.Pow(10, -dec);
                rounded = Math.Round(d / factor, 0, MidpointRounding.AwayFromZero) * factor;
                dec = 0;
            }
            if (double.IsNaN(rounded) || double.IsInfinity(rounded))
                return CellValue.FromError(Errors.Num);
            return CellValue.FromText(NumberFormatter.FormatNumber(rounded, dec, !noCommas));
        }

        private static CellValue FnUnaryMath(FunctionNode n, EvalContext ctx, string kind)
        {
            double d; CellValue err = CellValue.Empty;
            if (WrongArgs(n, 1, 1) || !GetNumber(n, 0, ctx, out d, out err))
                return err.IsError ? err : CellValue.FromError(Errors.Value);
            double r;
            switch (kind)
            {
                case "INT": r = Math.Floor(d); break;
                case "ABS": r = Math.Abs(d); break;
                case "SIGN": r = Math.Sign(d); break;
                case "SQRT":
                    if (d < 0) return CellValue.FromError(Errors.Num);
                    r = Math.Sqrt(d); break;
                case "EXP": r = Math.Exp(d); break;
                case "LN":
                    if (d <= 0) return CellValue.FromError(Errors.Num);
                    r = Math.Log(d); break;
                case "LOG10":
                    if (d <= 0) return CellValue.FromError(Errors.Num);
                    r = Math.Log10(d); break;
                default: return CellValue.FromError(Errors.Value);
            }
            if (double.IsNaN(r) || double.IsInfinity(r)) return CellValue.FromError(Errors.Num);
            return CellValue.FromNumber(r);
        }

        private static CellValue FnTrunc(FunctionNode n, EvalContext ctx)
        {
            double d, digits = 0; CellValue err = CellValue.Empty;
            if (WrongArgs(n, 1, 2) || !GetNumber(n, 0, ctx, out d, out err)) return err.IsError ? err : CellValue.FromError(Errors.Value);
            if (n.Args.Count >= 2 && !GetNumber(n, 1, ctx, out digits, out err)) return err;
            int dec = Math.Max(-15, Math.Min(15, (int)Math.Truncate(digits)));
            double factor = Math.Pow(10, dec);
            double result = Math.Truncate(d * factor) / factor;
            if (double.IsNaN(result) || double.IsInfinity(result)) return CellValue.FromError(Errors.Num);
            return CellValue.FromNumber(result);
        }

        private static CellValue FnLog(FunctionNode n, EvalContext ctx)
        {
            double d, baseValue = 10; CellValue err = CellValue.Empty;
            if (WrongArgs(n, 1, 2) || !GetNumber(n, 0, ctx, out d, out err)) return err.IsError ? err : CellValue.FromError(Errors.Value);
            if (n.Args.Count >= 2 && !GetNumber(n, 1, ctx, out baseValue, out err)) return err;
            if (d <= 0 || baseValue <= 0 || baseValue == 1) return CellValue.FromError(Errors.Num);
            return CellValue.FromNumber(Math.Log(d, baseValue));
        }

        private static CellValue FnMod(FunctionNode n, EvalContext ctx)
        {
            double a, b; CellValue err = CellValue.Empty;
            if (WrongArgs(n, 2, 2) || !GetNumber(n, 0, ctx, out a, out err) || !GetNumber(n, 1, ctx, out b, out err))
                return err.IsError ? err : CellValue.FromError(Errors.Value);
            if (b == 0) return CellValue.FromError(Errors.Div0);
            return CellValue.FromNumber(a - b * Math.Floor(a / b));
        }

        private static CellValue FnPower(FunctionNode n, EvalContext ctx)
        {
            double a, b; CellValue err = CellValue.Empty;
            if (WrongArgs(n, 2, 2) || !GetNumber(n, 0, ctx, out a, out err) || !GetNumber(n, 1, ctx, out b, out err))
                return err.IsError ? err : CellValue.FromError(Errors.Value);
            double r = Math.Pow(a, b);
            if (double.IsNaN(r) || double.IsInfinity(r)) return CellValue.FromError(Errors.Num);
            return CellValue.FromNumber(r);
        }

        private static CellValue FnRandBetween(FunctionNode n, EvalContext ctx)
        {
            double a, b; CellValue err = CellValue.Empty;
            if (WrongArgs(n, 2, 2) || !GetNumber(n, 0, ctx, out a, out err) || !GetNumber(n, 1, ctx, out b, out err))
                return err.IsError ? err : CellValue.FromError(Errors.Value);
            double loD = Math.Ceiling(a), hiD = Math.Floor(b);
            if (loD > hiD || loD < int.MinValue || hiD > int.MaxValue || hiD - loD > int.MaxValue)
                return CellValue.FromError(Errors.Num);
            long lo = (long)loD, hi = (long)hiD;
            if (lo == hi) return CellValue.FromNumber(lo);
            double span = (double)(hi - lo + 1);
            return CellValue.FromNumber(lo + Math.Floor(Rng.NextDouble() * span));
        }

        private static CellValue FnCeilingFloor(FunctionNode n, EvalContext ctx, bool ceiling)
        {
            double d, sig; CellValue err = CellValue.Empty;
            if (WrongArgs(n, 2, 2) || !GetNumber(n, 0, ctx, out d, out err) || !GetNumber(n, 1, ctx, out sig, out err))
                return err.IsError ? err : CellValue.FromError(Errors.Value);
            if (sig == 0) return CellValue.FromNumber(0);
            double r = ceiling ? Math.Ceiling(d / sig) * sig : Math.Floor(d / sig) * sig;
            return CellValue.FromNumber(r);
        }

        private static CellValue FnSumProduct(FunctionNode n, EvalContext ctx)
        {
            if (n.Args.Count == 0) return CellValue.FromError(Errors.Value);
            List<CellValue> vals = new List<CellValue>();
            foreach (FormulaNode node in n.Args) vals.Add(node.Eval(ctx));
            bool hasRange = false;
            int rows = 1, cols = 1;
            foreach (CellValue v in vals)
            {
                if (v.IsError) return v;
                if (v.Kind == ValueKind.Range)
                {
                    RangeRef rr = ClampRange(v.Range);
                    hasRange = true; rows = rr.Rows; cols = rr.Cols; break;
                }
            }
            if (!hasRange)
            {
                double product = 1;
                foreach (CellValue v in vals)
                {
                    double d; CellValue err = CellValue.Empty;
                    if (!EvalUtil.TryNumber(v, ctx, out d, out err)) return err;
                    product *= d;
                }
                return CellValue.FromNumber(product);
            }
            double sum = 0;
            for (int row = 0; row < rows; row++)
            {
                for (int col = 0; col < cols; col++)
                {
                    double product = 1;
                    bool bad = false;
                    foreach (CellValue v in vals)
                    {
                        double d = 0;
                        if (v.Kind == ValueKind.Range)
                        {
                            RangeRef rr = ClampRange(v.Range);
                            if (rr.Rows != rows || rr.Cols != cols) return CellValue.FromError(Errors.Value);
                            CellValue cv = ctx.GetCell(rr.Sheet, rr.Row1 + row, rr.Col1 + col);
                            if (cv.IsError) return cv;
                            if (cv.Kind == ValueKind.Number) d = cv.Number;
                        }
                        else
                        {
                            CellValue err = CellValue.Empty;
                            if (!EvalUtil.TryNumber(v, ctx, out d, out err)) return err;
                        }
                        product *= d;
                    }
                    if (!bad) sum += product;
                }
            }
            return CellValue.FromNumber(sum);
        }

        private static CellValue FnConditionalAggregate(FunctionNode n, EvalContext ctx, int kind)
        {
            int min = kind == 1 ? 2 : 2;
            int max = kind == 1 ? 2 : 3;
            if (WrongArgs(n, min, max)) return CellValue.FromError(Errors.Value);
            CellValue critRangeValue = Arg(n, 0, ctx);
            if (critRangeValue.IsError) return critRangeValue;
            if (critRangeValue.Kind != ValueKind.Range) return CellValue.FromError(Errors.Value);
            CellValue criteriaValue = Arg(n, 1, ctx);
            Criteria criteria = new Criteria(criteriaValue);
            RangeRef cr = ClampRange(critRangeValue.Range);
            RangeRef sr = cr;
            if (kind != 1 && n.Args.Count >= 3)
            {
                CellValue sv = Arg(n, 2, ctx);
                if (sv.Kind != ValueKind.Range) return CellValue.FromError(Errors.Value);
                sr = ClampRange(sv.Range);
            }
            double sum = 0;
            int matches = 0;
            int numeric = 0;
            for (int row = cr.Row1; row <= cr.Row2; row++)
            {
                for (int col = cr.Col1; col <= cr.Col2; col++)
                {
                    CellValue cv = ctx.GetCell(cr.Sheet, row, col);
                    if (cv.IsError) continue;
                    if (!criteria.Matches(cv)) continue;
                    matches++;
                    if (kind == 1) continue;
                    int srRow = sr.Row1 + (row - cr.Row1);
                    int srCol = sr.Col1 + (col - cr.Col1);
                    if (srRow > sr.Row2 || srCol > sr.Col2) continue;
                    CellValue sv = ctx.GetCell(sr.Sheet, srRow, srCol);
                    if (sv.IsError) return sv;
                    if (sv.Kind == ValueKind.Number) { sum += sv.Number; numeric++; }
                }
            }
            if (kind == 1) return CellValue.FromNumber(matches);
            if (kind == 2)
            {
                if (numeric == 0) return CellValue.FromError(Errors.Div0);
                return CellValue.FromNumber(sum / numeric);
            }
            return CellValue.FromNumber(sum);
        }

        private static CellValue FnMultiAggregate(FunctionNode n, EvalContext ctx, int kind)
        {
            bool countOnly = kind == 1;
            int min = countOnly ? 2 : 3;
            if (WrongArgs(n, min, -1)) return CellValue.FromError(Errors.Value);
            int criteriaStart = countOnly ? 0 : 1;
            int pairs = n.Args.Count - criteriaStart;
            if (pairs < 2 || (pairs % 2) != 0) return CellValue.FromError(Errors.Value);

            RangeRef sumRange = default(RangeRef);
            if (!countOnly)
            {
                CellValue sv = Arg(n, 0, ctx);
                if (sv.IsError) return sv;
                if (sv.Kind != ValueKind.Range) return CellValue.FromError(Errors.Value);
                sumRange = ClampRange(sv.Range);
            }

            List<RangeRef> ranges = new List<RangeRef>();
            List<Criteria> criteria = new List<Criteria>();
            for (int i = criteriaStart; i < n.Args.Count; i += 2)
            {
                CellValue rv = Arg(n, i, ctx);
                if (rv.IsError) return rv;
                if (rv.Kind != ValueKind.Range) return CellValue.FromError(Errors.Value);
                RangeRef rr = ClampRange(rv.Range);
                if (ranges.Count > 0 && (rr.Rows != ranges[0].Rows || rr.Cols != ranges[0].Cols))
                    return CellValue.FromError(Errors.Value);
                ranges.Add(rr);
                criteria.Add(new Criteria(Arg(n, i + 1, ctx)));
            }
            if (!countOnly && (sumRange.Rows != ranges[0].Rows || sumRange.Cols != ranges[0].Cols))
                return CellValue.FromError(Errors.Value);

            double sum = 0, minValue = double.MaxValue, maxValue = double.MinValue;
            int matches = 0, numeric = 0;
            RangeRef cr = ranges[0];
            for (int row = 0; row < cr.Rows; row++)
            {
                for (int col = 0; col < cr.Cols; col++)
                {
                    bool all = true;
                    for (int k = 0; k < ranges.Count; k++)
                    {
                        RangeRef rr = ranges[k];
                        CellValue cv = ctx.GetCell(rr.Sheet, rr.Row1 + row, rr.Col1 + col);
                        if (cv.IsError || !criteria[k].Matches(cv)) { all = false; break; }
                    }
                    if (!all) continue;
                    matches++;
                    if (countOnly) continue;
                    CellValue sv = ctx.GetCell(sumRange.Sheet, sumRange.Row1 + row, sumRange.Col1 + col);
                    if (sv.IsError) return sv;
                    if (sv.Kind != ValueKind.Number) continue;
                    numeric++;
                    sum += sv.Number;
                    if (sv.Number < minValue) minValue = sv.Number;
                    if (sv.Number > maxValue) maxValue = sv.Number;
                }
            }
            if (kind == 1) return CellValue.FromNumber(matches);
            if (kind == 2)
            {
                if (numeric == 0) return CellValue.FromError(Errors.Div0);
                return CellValue.FromNumber(sum / numeric);
            }
            if (kind == 3) return CellValue.FromNumber(numeric == 0 ? 0 : maxValue);
            if (kind == 4) return CellValue.FromNumber(numeric == 0 ? 0 : minValue);
            return CellValue.FromNumber(sum);
        }

        private static CellValue FnSubtotal(FunctionNode n, EvalContext ctx)
        {
            if (n.Args.Count < 2) return CellValue.FromError(Errors.Value);
            double codeD;
            CellValue err = CellValue.Empty;
            if (!GetNumber(n, 0, ctx, out codeD, out err)) return err;
            int code = (int)Math.Truncate(codeD);
            int fn = code >= 100 ? code - 100 : code;
            if (fn < 1 || fn > 11) return CellValue.FromError(Errors.Value);
            List<double> numbers = new List<double>();
            int nonEmpty = 0;
            for (int i = 1; i < n.Args.Count; i++)
            {
                CellValue v = Arg(n, i, ctx);
                if (v.Kind != ValueKind.Range)
                {
                    if (v.IsError) return v;
                    if (v.Kind != ValueKind.Empty) nonEmpty++;
                    double d;
                    if (EvalUtil.TryNumber(v, ctx, out d, out err)) numbers.Add(d);
                    continue;
                }
                RangeRef r = ClampRange(v.Range);
                for (int row = r.Row1; row <= r.Row2; row++)
                {
                    if ((code >= 100) && r.Sheet.HiddenRows.Contains(row)) continue;
                    for (int col = r.Col1; col <= r.Col2; col++)
                    {
                        CellValue cv = ctx.GetCell(r.Sheet, row, col);
                        if (cv.IsError) continue;
                        if (cv.Kind != ValueKind.Empty) nonEmpty++;
                        if (cv.Kind == ValueKind.Number) numbers.Add(cv.Number);
                    }
                }
            }
            switch (fn)
            {
                case 1:
                    if (numbers.Count == 0) return CellValue.FromError(Errors.Div0);
                    double avg = 0; foreach (double d in numbers) avg += d;
                    return CellValue.FromNumber(avg / numbers.Count);
                case 2: return CellValue.FromNumber(numbers.Count);
                case 3: return CellValue.FromNumber(nonEmpty);
                case 4:
                case 5:
                    if (numbers.Count == 0) return CellValue.FromNumber(0);
                    double best = numbers[0];
                    for (int i = 1; i < numbers.Count; i++)
                        best = fn == 4 ? Math.Max(best, numbers[i]) : Math.Min(best, numbers[i]);
                    return CellValue.FromNumber(best);
                case 6:
                    double product = 1; foreach (double d in numbers) product *= d;
                    return CellValue.FromNumber(numbers.Count == 0 ? 0 : product);
                case 7:
                case 8:
                case 9:
                case 10:
                case 11:
                    {
                        if (fn == 9)
                        {
                            double total = 0; foreach (double d in numbers) total += d;
                            return CellValue.FromNumber(total);
                        }
                        bool sample = fn == 7 || fn == 10;
                        if (numbers.Count < (sample ? 2 : 1)) return CellValue.FromError(Errors.Div0);
                        double mean = 0; foreach (double d in numbers) mean += d;
                        mean /= numbers.Count;
                        double ss = 0; foreach (double d in numbers) { double q = d - mean; ss += q * q; }
                        double variance = ss / (sample ? numbers.Count - 1 : numbers.Count);
                        bool stdev = fn == 7 || fn == 8;
                        return CellValue.FromNumber(stdev ? Math.Sqrt(variance) : variance);
                    }
            }
            return CellValue.FromError(Errors.Value);
        }

        private static CellValue FnAggregate(FunctionNode n, EvalContext ctx)
        {
            if (n.Args.Count < 3) return CellValue.FromError(Errors.Value);
            double fnD, optD; CellValue err = CellValue.Empty;
            if (!GetNumber(n, 0, ctx, out fnD, out err) || !GetNumber(n, 1, ctx, out optD, out err)) return err;
            int fn = (int)Math.Truncate(fnD), options = (int)Math.Truncate(optD);
            if (options < 0 || options > 7) return CellValue.FromError(Errors.Value);
            bool ignoreErrors = (options & 4) != 0;
            bool ignoreHidden = (options & 1) != 0;
            List<double> nums = new List<double>();
            int nonEmpty = 0;
            for (int i = 2; i < n.Args.Count; i++)
            {
                CellValue v = Arg(n, i, ctx);
                if (v.Kind != ValueKind.Range)
                {
                    if (v.IsError && !ignoreErrors) return v;
                    if (v.Kind != ValueKind.Empty) nonEmpty++;
                    if (v.Kind == ValueKind.Number) nums.Add(v.Number);
                    continue;
                }
                RangeRef r = ClampRange(v.Range);
                for (int row = r.Row1; row <= r.Row2; row++)
                {
                    if (ignoreHidden && r.Sheet.HiddenRows.Contains(row)) continue;
                    for (int col = r.Col1; col <= r.Col2; col++)
                    {
                        CellValue cv = ctx.GetCell(r.Sheet, row, col);
                        if (cv.IsError) { if (ignoreErrors) continue; return cv; }
                        if (cv.Kind != ValueKind.Empty) nonEmpty++;
                        if (cv.Kind == ValueKind.Number) nums.Add(cv.Number);
                    }
                }
            }
            switch (fn)
            {
                case 1: return nums.Count == 0 ? CellValue.FromError(Errors.Div0)
                    : CellValue.FromNumber(nums.Sum() / nums.Count);
                case 2: return CellValue.FromNumber(nums.Count);
                case 3: return CellValue.FromNumber(nonEmpty);
                case 4: return nums.Count == 0 ? CellValue.FromNumber(0) : CellValue.FromNumber(nums.Max());
                case 5: return nums.Count == 0 ? CellValue.FromNumber(0) : CellValue.FromNumber(nums.Min());
                case 6:
                    double product = 1; foreach (double d in nums) product *= d;
                    return CellValue.FromNumber(nums.Count == 0 ? 0 : product);
                case 7:
                case 8:
                case 10:
                case 11:
                    {
                        if (nums.Count < ((fn == 7 || fn == 10) ? 2 : 1))
                            return CellValue.FromError(Errors.Div0);
                        double mean = nums.Sum() / nums.Count;
                        double ss = 0; foreach (double d in nums) { double q = d - mean; ss += q * q; }
                        double variance = ss / ((fn == 7 || fn == 10) ? nums.Count - 1 : nums.Count);
                        bool stdev = fn == 7 || fn == 8;
                        return CellValue.FromNumber(stdev ? Math.Sqrt(variance) : variance);
                    }
                case 9: return CellValue.FromNumber(nums.Sum());
                case 12:
                    if (nums.Count == 0) return CellValue.FromError(Errors.Num);
                    nums.Sort();
                    return CellValue.FromNumber(nums.Count % 2 == 1 ? nums[nums.Count / 2]
                        : (nums[nums.Count / 2 - 1] + nums[nums.Count / 2]) / 2.0);
                default: return CellValue.FromError(Errors.Value);
            }
        }

        private static CellValue FnConcat(FunctionNode n, EvalContext ctx)
        {
            StringBuilder sb = new StringBuilder();
            foreach (FormulaNode node in n.Args)
            {
                CellValue v = node.Eval(ctx);
                if (v.Kind == ValueKind.Range)
                {
                    foreach (CellValue cv in RangeValues(v.Range, ctx))
                    {
                        if (cv.IsError) return cv;
                        if (cv.Kind == ValueKind.Empty) continue;
                        sb.Append(cv.ToDisplayString());
                        if (sb.Length > 32767) return CellValue.FromError(Errors.Value);
                    }
                }
                else
                {
                    if (v.IsError) return v;
                    if (v.Kind == ValueKind.Empty) continue;
                    CellValue err = CellValue.Empty;
                    string s = EvalUtil.ToText(v, ctx, out err);
                    if (err.IsError) return err;
                    sb.Append(s);
                    if (sb.Length > 32767) return CellValue.FromError(Errors.Value);
                }
            }
            return CellValue.FromText(sb.ToString());
        }

        private static CellValue FnTextJoin(FunctionNode n, EvalContext ctx)
        {
            if (n.Args.Count < 3) return CellValue.FromError(Errors.Value);
            string delim; CellValue err = CellValue.Empty;
            if (!GetText(n, 0, ctx, out delim, out err)) return err;
            bool ignoreEmpty;
            if (!GetBool(n, 1, ctx, out ignoreEmpty, out err)) return err;
            StringBuilder sb = new StringBuilder();
            bool first = true;
            for (int i = 2; i < n.Args.Count; i++)
            {
                CellValue v = Arg(n, i, ctx);
                if (v.Kind == ValueKind.Range)
                {
                    foreach (CellValue cv in RangeValues(v.Range, ctx))
                    {
                        if (cv.IsError) return cv;
                        if (cv.Kind == ValueKind.Empty) { if (ignoreEmpty) continue; }
                        string s = cv.ToDisplayString();
                        if (ignoreEmpty && s.Length == 0) continue;
                        if (!first) sb.Append(delim);
                        sb.Append(s); first = false;
                        if (sb.Length > 32767) return CellValue.FromError(Errors.Value);
                    }
                }
                else
                {
                    if (v.IsError) return v;
                    if (v.Kind == ValueKind.Empty && ignoreEmpty) continue;
                    string s; if (!GetText(n, i, ctx, out s, out err)) return err;
                    if (ignoreEmpty && s.Length == 0) continue;
                    if (!first) sb.Append(delim);
                    sb.Append(s); first = false;
                    if (sb.Length > 32767) return CellValue.FromError(Errors.Value);
                }
            }
            return CellValue.FromText(sb.ToString());
        }

        private static CellValue FnTextBeforeAfter(FunctionNode n, EvalContext ctx, bool before)
        {
            if (WrongArgs(n, 2, 6)) return CellValue.FromError(Errors.Value);
            string text, delimiter; CellValue err = CellValue.Empty;
            if (!GetText(n, 0, ctx, out text, out err)) return err;
            if (!GetText(n, 1, ctx, out delimiter, out err)) return err;
            int instance = 1;
            if (n.Args.Count >= 3)
            {
                double d;
                if (!GetNumber(n, 2, ctx, out d, out err)) return err;
                instance = (int)Math.Truncate(d);
            }
            bool caseInsensitive = false;
            if (n.Args.Count >= 4)
            {
                double d;
                if (!GetNumber(n, 3, ctx, out d, out err)) return err;
                caseInsensitive = ((int)Math.Truncate(d) & 1) == 0;
            }
            StringComparison cmp = caseInsensitive ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            int pos = -1, start = 0;
            int count = Math.Max(1, Math.Abs(instance));
            for (int i = 0; i < count; i++)
            {
                pos = text.IndexOf(delimiter, start, cmp);
                if (pos < 0) break;
                start = pos + Math.Max(1, delimiter.Length);
            }
            if (pos < 0)
            {
                if (n.Args.Count >= 6) return Arg(n, 5, ctx);
                return CellValue.FromError(Errors.NA);
            }
            return CellValue.FromText(before ? text.Substring(0, pos)
                                             : text.Substring(pos + delimiter.Length));
        }

        private static CellValue FnTextLength(FunctionNode n, EvalContext ctx)
        {
            string s; CellValue err = CellValue.Empty;
            if (WrongArgs(n, 1, 1) || !GetText(n, 0, ctx, out s, out err)) return err.IsError ? err : CellValue.FromError(Errors.Value);
            return CellValue.FromNumber(s.Length);
        }

        private static CellValue FnLeftRight(FunctionNode n, EvalContext ctx, bool left)
        {
            if (WrongArgs(n, 1, 2)) return CellValue.FromError(Errors.Value);
            string s; CellValue err = CellValue.Empty;
            if (!GetText(n, 0, ctx, out s, out err)) return err;
            double count = 1;
            if (n.Args.Count >= 2 && !GetNumber(n, 1, ctx, out count, out err)) return err;
            int c = (int)Math.Truncate(count);
            if (c < 0) return CellValue.FromError(Errors.Value);
            if (c >= s.Length) return CellValue.FromText(s);
            return CellValue.FromText(left ? s.Substring(0, c) : s.Substring(s.Length - c));
        }

        private static CellValue FnMid(FunctionNode n, EvalContext ctx)
        {
            string s; CellValue err = CellValue.Empty; double start, count;
            if (WrongArgs(n, 3, 3) || !GetText(n, 0, ctx, out s, out err) ||
                !GetNumber(n, 1, ctx, out start, out err) || !GetNumber(n, 2, ctx, out count, out err))
                return err.IsError ? err : CellValue.FromError(Errors.Value);
            int st = (int)Math.Truncate(start);
            int ct = (int)Math.Truncate(count);
            if (st < 1 || ct < 0) return CellValue.FromError(Errors.Value);
            if (st > s.Length) return CellValue.FromText("");
            int len = Math.Min(ct, s.Length - (st - 1));
            return CellValue.FromText(s.Substring(st - 1, len));
        }

        private static CellValue FnUpperLower(FunctionNode n, EvalContext ctx, bool upper)
        {
            string s; CellValue err = CellValue.Empty;
            if (WrongArgs(n, 1, 1) || !GetText(n, 0, ctx, out s, out err)) return err.IsError ? err : CellValue.FromError(Errors.Value);
            return CellValue.FromText(upper ? s.ToUpper(CultureInfo.CurrentCulture) : s.ToLower(CultureInfo.CurrentCulture));
        }

        private static CellValue FnTrim(FunctionNode n, EvalContext ctx)
        {
            string s; CellValue err = CellValue.Empty;
            if (WrongArgs(n, 1, 1) || !GetText(n, 0, ctx, out s, out err)) return err.IsError ? err : CellValue.FromError(Errors.Value);
            StringBuilder sb = new StringBuilder();
            bool lastSpace = false;
            foreach (char ch in s.Trim())
            {
                if (ch == ' ')
                {
                    if (!lastSpace) sb.Append(ch);
                    lastSpace = true;
                }
                else
                {
                    sb.Append(ch);
                    lastSpace = false;
                }
            }
            return CellValue.FromText(sb.ToString());
        }

        private static CellValue FnProper(FunctionNode n, EvalContext ctx)
        {
            string s; CellValue err = CellValue.Empty;
            if (WrongArgs(n, 1, 1) || !GetText(n, 0, ctx, out s, out err)) return err.IsError ? err : CellValue.FromError(Errors.Value);
            StringBuilder sb = new StringBuilder(s.Length);
            bool newWord = true;
            foreach (char ch in s)
            {
                if (char.IsLetter(ch))
                {
                    sb.Append(newWord ? char.ToUpper(ch, CultureInfo.CurrentCulture) : char.ToLower(ch, CultureInfo.CurrentCulture));
                    newWord = false;
                }
                else
                {
                    sb.Append(ch);
                    newWord = true;
                }
            }
            return CellValue.FromText(sb.ToString());
        }

        private static CellValue FnFormulaText(FunctionNode n, EvalContext ctx)
        {
            if (WrongArgs(n, 1, 1)) return CellValue.FromError(Errors.Value);
            Cell cell = null;
            RefNode rn = n.Args.Count > 0 ? n.Args[0] as RefNode : null;
            if (rn != null)
            {
                Sheet sh = ctx.ResolveSheet(rn.SheetName);
                if (sh != null) cell = sh.GetCell(rn.Row, rn.Col);
            }
            if (cell == null)
            {
                CellValue v = Arg(n, 0, ctx);
                if (v.Kind == ValueKind.Range && v.Range != null &&
                    v.Range.Row1 == v.Range.Row2 && v.Range.Col1 == v.Range.Col2)
                    cell = v.Range.Sheet.GetCell(v.Range.Row1, v.Range.Col1);
            }
            if (cell == null || !cell.IsFormula || string.IsNullOrEmpty(cell.Raw))
                return CellValue.FromError(Errors.NA);
            return CellValue.FromText(cell.Raw.StartsWith("=", StringComparison.Ordinal)
                ? cell.Raw : "=" + cell.Raw);
        }

        private static CellValue FnText(FunctionNode n, EvalContext ctx)
        {
            double d; string fmt; CellValue err = CellValue.Empty;
            if (WrongArgs(n, 2, 2) || !GetNumber(n, 0, ctx, out d, out err) || !GetText(n, 1, ctx, out fmt, out err))
                return err.IsError ? err : CellValue.FromError(Errors.Value);
            string lower = fmt.ToLowerInvariant();
            CellStyle st = new CellStyle();
            bool dateLike = lower.IndexOf("yy", StringComparison.Ordinal) >= 0 ||
                            lower.IndexOf("dd", StringComparison.Ordinal) >= 0 ||
                            lower.IndexOf("ddd", StringComparison.Ordinal) >= 0 ||
                            lower.IndexOf("mmm", StringComparison.Ordinal) >= 0 ||
                            lower.IndexOf("d/", StringComparison.Ordinal) >= 0 ||
                            lower.IndexOf("d-", StringComparison.Ordinal) >= 0 ||
                            lower.Trim() == "d";
            if (dateLike)
                return CellValue.FromText(ExcelNumberFormat.Format(d, fmt));
            if (lower.IndexOf("hh:") >= 0 || lower.IndexOf("mm:") >= 0)
            {
                st.Format = NumberFormatKind.Time;
                return CellValue.FromText(NumberFormatter.Format(CellValue.FromNumber(d), st));
            }
            int decimals = 0;
            int dot = lower.LastIndexOf('.');
            if (dot >= 0)
                for (int i = dot + 1; i < lower.Length && (lower[i] == '0' || lower[i] == '#'); i++) decimals++;
            st.Decimals = decimals;
            if (lower.IndexOf('%') >= 0) st.Format = NumberFormatKind.Percent;
            else if (lower.IndexOf('$') >= 0) st.Format = NumberFormatKind.Currency;
            else if (lower.IndexOf('#') >= 0 || lower.IndexOf('0') >= 0) st.Format = NumberFormatKind.Number;
            else return CellValue.FromText(NumberFormatter.FormatGeneral(d));
            return CellValue.FromText(NumberFormatter.Format(CellValue.FromNumber(d), st));
        }

        private static CellValue FnValue(FunctionNode n, EvalContext ctx)
        {
            string s; CellValue err = CellValue.Empty;
            if (WrongArgs(n, 1, 1) || !GetText(n, 0, ctx, out s, out err)) return err.IsError ? err : CellValue.FromError(Errors.Value);
            double d;
            if (NumberParser.TryParse(s, out d)) return CellValue.FromNumber(d);
            DateTime dt;
            if (DateTime.TryParse(s, CultureInfo.GetCultureInfo("es-CO"), DateTimeStyles.None, out dt) ||
                DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                return CellValue.FromNumber(dt.ToOADate());
            return CellValue.FromError(Errors.Value);
        }

        private static CellValue FnRept(FunctionNode n, EvalContext ctx)
        {
            string s; CellValue err = CellValue.Empty; double times;
            if (WrongArgs(n, 2, 2) || !GetText(n, 0, ctx, out s, out err) || !GetNumber(n, 1, ctx, out times, out err))
                return err.IsError ? err : CellValue.FromError(Errors.Value);
            if (times < 0 || times > 32767) return CellValue.FromError(Errors.Value);
            if (s.Length == 0) return CellValue.FromText("");
            int count = (int)Math.Truncate(times);
            if ((long)s.Length * count > 32767) return CellValue.FromError(Errors.Value);
            StringBuilder sb = new StringBuilder(s.Length * count);
            for (int i = 0; i < count; i++) sb.Append(s);
            return CellValue.FromText(sb.ToString());
        }

        private static CellValue FnSubstitute(FunctionNode n, EvalContext ctx)
        {
            string text, oldText, newText; CellValue err = CellValue.Empty;
            if (WrongArgs(n, 3, 4) || !GetText(n, 0, ctx, out text, out err) ||
                !GetText(n, 1, ctx, out oldText, out err) || !GetText(n, 2, ctx, out newText, out err))
                return err.IsError ? err : CellValue.FromError(Errors.Value);
            if (oldText.Length == 0) return CellValue.FromText(text);
            if (n.Args.Count == 4)
            {
                double inst; 
                if (!GetNumber(n, 3, ctx, out inst, out err)) return err;
                int instance = (int)Math.Truncate(inst);
                if (instance < 1) return CellValue.FromError(Errors.Value);
                int pos = -1, seen = 0;
                for (int i = 0; i <= text.Length - oldText.Length; i++)
                {
                    if (string.CompareOrdinal(text, i, oldText, 0, oldText.Length) == 0)
                    {
                        seen++;
                        if (seen == instance) { pos = i; break; }
                        i += oldText.Length - 1;
                    }
                }
                if (pos < 0) return CellValue.FromText(text);
                return CellValue.FromText(text.Substring(0, pos) + newText + text.Substring(pos + oldText.Length));
            }
            return CellValue.FromText(text.Replace(oldText, newText));
        }

        private static CellValue FnFind(FunctionNode n, EvalContext ctx, bool caseSensitive)
        {
            string find, within; CellValue err = CellValue.Empty; double start = 1;
            if (WrongArgs(n, 2, 3) || !GetText(n, 0, ctx, out find, out err) ||
                !GetText(n, 1, ctx, out within, out err))
                return err.IsError ? err : CellValue.FromError(Errors.Value);
            if (n.Args.Count >= 3 && !GetNumber(n, 2, ctx, out start, out err)) return err;
            int st = (int)Math.Truncate(start);
            if (st < 1 || st > within.Length + 1) return CellValue.FromError(Errors.Value);
            int index = within.IndexOf(find, st - 1,
                caseSensitive ? StringComparison.Ordinal : StringComparison.CurrentCultureIgnoreCase);
            if (index < 0) return CellValue.FromError(Errors.Value);
            return CellValue.FromNumber(index + 1);
        }

        private static CellValue FnExact(FunctionNode n, EvalContext ctx)
        {
            string a, b; CellValue err = CellValue.Empty;
            if (WrongArgs(n, 2, 2) || !GetText(n, 0, ctx, out a, out err) || !GetText(n, 1, ctx, out b, out err))
                return err.IsError ? err : CellValue.FromError(Errors.Value);
            return CellValue.FromBool(string.Equals(a, b, StringComparison.Ordinal));
        }

        private static CellValue FnChar(FunctionNode n, EvalContext ctx)
        {
            double d; CellValue err = CellValue.Empty;
            if (WrongArgs(n, 1, 1) || !GetNumber(n, 0, ctx, out d, out err)) return err.IsError ? err : CellValue.FromError(Errors.Value);
            int code = (int)Math.Truncate(d);
            if (code < 1 || code > 65535) return CellValue.FromError(Errors.Value);
            return CellValue.FromText(((char)code).ToString());
        }

        private static CellValue FnCode(FunctionNode n, EvalContext ctx)
        {
            string s; CellValue err = CellValue.Empty;
            if (WrongArgs(n, 1, 1) || !GetText(n, 0, ctx, out s, out err)) return err.IsError ? err : CellValue.FromError(Errors.Value);
            if (s.Length == 0) return CellValue.FromError(Errors.Value);
            return CellValue.FromNumber((int)s[0]);
        }

        private static bool TryDate(double serial, out DateTime dt)
        {
            dt = DateTime.MinValue;
            if (serial < 0 || serial > 2958465) return false;
            try { dt = DateTime.FromOADate(serial); return true; }
            catch { return false; }
        }

        private static CellValue FnDate(FunctionNode n, EvalContext ctx)
        {
            double y, m, d; CellValue err = CellValue.Empty;
            if (WrongArgs(n, 3, 3) || !GetNumber(n, 0, ctx, out y, out err) ||
                !GetNumber(n, 1, ctx, out m, out err) || !GetNumber(n, 2, ctx, out d, out err))
                return err.IsError ? err : CellValue.FromError(Errors.Value);
            int year = (int)Math.Truncate(y);
            int month = (int)Math.Truncate(m);
            int day = (int)Math.Truncate(d);
            if (year < 0 || year > 9999) return CellValue.FromError(Errors.Num);
            if (year < 1900) year += 1900;
            try
            {
                DateTime dt = new DateTime(year, 1, 1).AddMonths(month - 1).AddDays(day - 1);
                return CellValue.FromNumber(dt.ToOADate());
            }
            catch { return CellValue.FromError(Errors.Num); }
        }

        private static CellValue FnTime(FunctionNode n, EvalContext ctx)
        {
            double h, m, s; CellValue err = CellValue.Empty;
            if (WrongArgs(n, 3, 3) || !GetNumber(n, 0, ctx, out h, out err) ||
                !GetNumber(n, 1, ctx, out m, out err) || !GetNumber(n, 2, ctx, out s, out err))
                return err.IsError ? err : CellValue.FromError(Errors.Value);
            double total = Math.Truncate(h) * 3600 + Math.Truncate(m) * 60 + Math.Truncate(s);
            if (total < 0 || double.IsNaN(total) || double.IsInfinity(total))
                return CellValue.FromError(Errors.Num);
            return CellValue.FromNumber((total % 86400.0) / 86400.0);
        }

        private static CellValue FnTimeValue(FunctionNode n, EvalContext ctx)
        {
            string s; CellValue err = CellValue.Empty;
            if (WrongArgs(n, 1, 1) || !GetText(n, 0, ctx, out s, out err)) return err;
            DateTime dt;
            string[] formats = new string[] { "H:mm", "H:mm:ss", "HH:mm", "HH:mm:ss", "h:mm tt", "h:mm:ss tt" };
            if (DateTime.TryParseExact(s.Trim(), formats, CultureInfo.GetCultureInfo("es-CO"),
                DateTimeStyles.None, out dt) ||
                DateTime.TryParse(s, CultureInfo.GetCultureInfo("es-CO"), DateTimeStyles.None, out dt) ||
                DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                return CellValue.FromNumber(dt.TimeOfDay.TotalDays);
            return CellValue.FromError(Errors.Value);
        }

        private static CellValue FnDays(FunctionNode n, EvalContext ctx)
        {
            double end, start; CellValue err = CellValue.Empty;
            if (WrongArgs(n, 2, 2) || !GetNumber(n, 0, ctx, out end, out err) ||
                !GetNumber(n, 1, ctx, out start, out err)) return err;
            DateTime a, b;
            if (!TryDate(end, out a) || !TryDate(start, out b)) return CellValue.FromError(Errors.Num);
            return CellValue.FromNumber((a.Date - b.Date).TotalDays);
        }

        private static CellValue FnDateDif(FunctionNode n, EvalContext ctx)
        {
            double start, end; string unit; CellValue err = CellValue.Empty;
            if (WrongArgs(n, 3, 3) || !GetNumber(n, 0, ctx, out start, out err) ||
                !GetNumber(n, 1, ctx, out end, out err) || !GetText(n, 2, ctx, out unit, out err))
                return err;
            DateTime a, b;
            if (!TryDate(start, out a) || !TryDate(end, out b) || b < a) return CellValue.FromError(Errors.Num);
            string u = (unit ?? "").Trim().ToUpperInvariant();
            int years = b.Year - a.Year;
            int months = years * 12 + b.Month - a.Month;
            if (b.Day < a.Day) months--;
            int days = (int)(b.Date - a.Date).TotalDays;
            if (u == "Y") return CellValue.FromNumber(months / 12);
            if (u == "M") return CellValue.FromNumber(months);
            if (u == "D") return CellValue.FromNumber(days);
            if (u == "YM") return CellValue.FromNumber(months % 12);
            if (u == "MD")
            {
                int d = b.Day - a.Day;
                if (d < 0)
                {
                    DateTime prev = b.AddMonths(-1);
                    d = (b.Date - new DateTime(prev.Year, prev.Month, a.Day)).Days;
                }
                return CellValue.FromNumber(d);
            }
            if (u == "YD")
            {
                DateTime anniversary = new DateTime(b.Year, a.Month, Math.Min(a.Day, DateTime.DaysInMonth(b.Year, a.Month)));
                if (anniversary > b) anniversary = anniversary.AddYears(-1);
                return CellValue.FromNumber((b.Date - anniversary).Days);
            }
            return CellValue.FromError(Errors.Num);
        }

        private static HashSet<int> HolidaySerials(FunctionNode n, int index, EvalContext ctx)
        {
            HashSet<int> holidays = new HashSet<int>();
            if (index >= n.Args.Count) return holidays;
            CellValue v = Arg(n, index, ctx);
            if (v.Kind == ValueKind.Range)
            {
                foreach (CellValue cv in RangeValues(v.Range, ctx))
                    if (cv.Kind == ValueKind.Number) holidays.Add((int)Math.Floor(cv.Number));
            }
            else if (v.Kind == ValueKind.Number) holidays.Add((int)Math.Floor(v.Number));
            return holidays;
        }

        private static CellValue FnNetworkDays(FunctionNode n, EvalContext ctx)
        {
            if (WrongArgs(n, 2, 3)) return CellValue.FromError(Errors.Value);
            double start, end; CellValue err = CellValue.Empty;
            if (!GetNumber(n, 0, ctx, out start, out err) || !GetNumber(n, 1, ctx, out end, out err)) return err;
            DateTime a, b;
            if (!TryDate(start, out a) || !TryDate(end, out b)) return CellValue.FromError(Errors.Num);
            if (b < a) { DateTime t = a; a = b; b = t; }
            HashSet<int> holidays = HolidaySerials(n, 2, ctx);
            int count = 0;
            for (DateTime d = a.Date; d <= b.Date; d = d.AddDays(1))
                if (d.DayOfWeek != DayOfWeek.Saturday && d.DayOfWeek != DayOfWeek.Sunday &&
                    !holidays.Contains((int)d.ToOADate())) count++;
            return CellValue.FromNumber(count);
        }

        private static CellValue FnWorkday(FunctionNode n, EvalContext ctx)
        {
            if (WrongArgs(n, 2, 3)) return CellValue.FromError(Errors.Value);
            double start, daysD; CellValue err = CellValue.Empty;
            if (!GetNumber(n, 0, ctx, out start, out err) || !GetNumber(n, 1, ctx, out daysD, out err)) return err;
            DateTime d;
            if (!TryDate(start, out d)) return CellValue.FromError(Errors.Num);
            int days = (int)Math.Truncate(daysD);
            HashSet<int> holidays = HolidaySerials(n, 2, ctx);
            if (days == 0) return CellValue.FromNumber(d.ToOADate());
            int step = days > 0 ? 1 : -1;
            int left = Math.Abs(days);
            for (int guard = 0; guard < 2000000 && left > 0; guard++)
            {
                d = d.AddDays(step);
                if (d.DayOfWeek == DayOfWeek.Saturday || d.DayOfWeek == DayOfWeek.Sunday) continue;
                if (holidays.Contains((int)d.ToOADate())) continue;
                left--;
            }
            if (left != 0) return CellValue.FromError(Errors.Num);
            return CellValue.FromNumber(d.ToOADate());
        }

        private static CellValue FnYearFrac(FunctionNode n, EvalContext ctx)
        {
            if (WrongArgs(n, 2, 3)) return CellValue.FromError(Errors.Value);
            double start, end; CellValue err = CellValue.Empty;
            if (!GetNumber(n, 0, ctx, out start, out err) || !GetNumber(n, 1, ctx, out end, out err)) return err;
            DateTime a, b;
            if (!TryDate(start, out a) || !TryDate(end, out b)) return CellValue.FromError(Errors.Num);
            double basis = 0;
            if (n.Args.Count >= 3 && !GetNumber(n, 2, ctx, out basis, out err)) return err;
            int bs = Math.Max(0, Math.Min(4, (int)Math.Truncate(basis)));
            if (a == b) return CellValue.FromNumber(0);
            double days = (b.Date - a.Date).TotalDays;
            if (bs == 0)
            {
                int years = b.Year - a.Year;
                if (years <= 0) return CellValue.FromNumber(days / 365.0);
                bool leap = DateTime.IsLeapYear(a.Year) || DateTime.IsLeapYear(b.Year);
                return CellValue.FromNumber(days / (leap ? 366.0 : 365.0));
            }
            if (bs == 1) return CellValue.FromNumber(days / 365.0);
            if (bs == 2) return CellValue.FromNumber(days / 360.0);
            if (bs == 3) return CellValue.FromNumber(days / 365.0);
            return CellValue.FromNumber(days / 360.0);
        }

        private static CellValue FnRank(FunctionNode n, EvalContext ctx)
        {
            if (WrongArgs(n, 2, 3)) return CellValue.FromError(Errors.Value);
            double value, order = 0; CellValue err = CellValue.Empty;
            if (!GetNumber(n, 0, ctx, out value, out err)) return err;
            if (n.Args.Count >= 3 && !GetNumber(n, 2, ctx, out order, out err)) return err;
            CellValue range = Arg(n, 1, ctx);
            if (range.Kind != ValueKind.Range) return CellValue.FromError(Errors.Value);
            List<double> nums = new List<double>();
            foreach (CellValue cv in RangeValues(range.Range, ctx))
                if (cv.Kind == ValueKind.Number) nums.Add(cv.Number);
            bool found = false;
            int rank = 1;
            foreach (double d in nums)
            {
                if (Math.Abs(d - value) < 1e-12) { found = true; continue; }
                if (order == 0 ? d > value : d < value) rank++;
            }
            if (!found) return CellValue.FromError(Errors.NA);
            return CellValue.FromNumber(rank);
        }

        private static List<double> NumericRange(FunctionNode n, int index, EvalContext ctx, out CellValue error)
        {
            List<double> nums = new List<double>();
            error = CellValue.Empty;
            CellValue v = Arg(n, index, ctx);
            if (v.Kind != ValueKind.Range) { error = CellValue.FromError(Errors.Value); return nums; }
            foreach (CellValue cv in RangeValues(v.Range, ctx))
            {
                if (cv.IsError) { error = cv; return nums; }
                if (cv.Kind == ValueKind.Number) nums.Add(cv.Number);
            }
            return nums;
        }

        private static CellValue FnPercentile(FunctionNode n, EvalContext ctx)
        {
            if (WrongArgs(n, 2, 2)) return CellValue.FromError(Errors.Value);
            CellValue err;
            List<double> nums = NumericRange(n, 0, ctx, out err);
            if (err.IsError) return err;
            double p; err = CellValue.Empty;
            if (!GetNumber(n, 1, ctx, out p, out err)) return err;
            if (p < 0 || p > 1 || nums.Count == 0) return CellValue.FromError(Errors.Num);
            nums.Sort();
            return CellValue.FromNumber(PercentileValue(nums, p));
        }

        private static double PercentileValue(List<double> sorted, double p)
        {
            double pos = p * (sorted.Count - 1);
            int lo = (int)Math.Floor(pos);
            int hi = (int)Math.Ceiling(pos);
            if (lo == hi) return sorted[lo];
            double frac = pos - lo;
            return sorted[lo] * (1 - frac) + sorted[hi] * frac;
        }

        private static CellValue FnQuartile(FunctionNode n, EvalContext ctx)
        {
            if (WrongArgs(n, 2, 2)) return CellValue.FromError(Errors.Value);
            double q; CellValue err = CellValue.Empty;
            if (!GetNumber(n, 1, ctx, out q, out err)) return err;
            int qi = (int)Math.Truncate(q);
            if (qi < 0 || qi > 4) return CellValue.FromError(Errors.Num);
            List<double> nums = NumericRange(n, 0, ctx, out err);
            if (err.IsError) return err;
            if (nums.Count == 0) return CellValue.FromError(Errors.Num);
            nums.Sort();
            return CellValue.FromNumber(PercentileValue(nums, qi / 4.0));
        }

        private static CellValue FnCeilingMath(FunctionNode n, EvalContext ctx, bool ceiling)
        {
            if (WrongArgs(n, 1, 3)) return CellValue.FromError(Errors.Value);
            double value, sig = 1; CellValue err = CellValue.Empty;
            if (!GetNumber(n, 0, ctx, out value, out err)) return err;
            if (n.Args.Count >= 2 && !GetNumber(n, 1, ctx, out sig, out err)) return err;
            sig = Math.Abs(sig);
            if (sig == 0) return CellValue.FromNumber(0);
            double q = value / sig;
            double r = ceiling ? Math.Ceiling(q) * sig : Math.Floor(q) * sig;
            return CellValue.FromNumber(r);
        }

        private static CellValue FnDatePart(FunctionNode n, EvalContext ctx, string part)
        {
            double serial; CellValue err = CellValue.Empty;
            if (WrongArgs(n, 1, 1) || !GetNumber(n, 0, ctx, out serial, out err)) return err.IsError ? err : CellValue.FromError(Errors.Value);
            DateTime dt;
            if (!TryDate(serial, out dt)) return CellValue.FromError(Errors.Num);
            switch (part)
            {
                case "YEAR": return CellValue.FromNumber(dt.Year);
                case "MONTH": return CellValue.FromNumber(dt.Month);
                case "DAY": return CellValue.FromNumber(dt.Day);
                case "HOUR": return CellValue.FromNumber(dt.Hour);
                case "MINUTE": return CellValue.FromNumber(dt.Minute);
                default: return CellValue.FromNumber(dt.Second);
            }
        }

        private static CellValue FnWeekday(FunctionNode n, EvalContext ctx)
        {
            double serial, type = 1; CellValue err = CellValue.Empty;
            if (WrongArgs(n, 1, 2) || !GetNumber(n, 0, ctx, out serial, out err)) return err.IsError ? err : CellValue.FromError(Errors.Value);
            if (n.Args.Count >= 2 && !GetNumber(n, 1, ctx, out type, out err)) return err;
            DateTime dt;
            if (!TryDate(serial, out dt)) return CellValue.FromError(Errors.Num);
            int dow = (int)dt.DayOfWeek;
            switch ((int)Math.Truncate(type))
            {
                case 1: return CellValue.FromNumber(dow + 1);
                case 2: return CellValue.FromNumber(((dow + 6) % 7) + 1);
                case 3: return CellValue.FromNumber((dow + 6) % 7);
                default: return CellValue.FromError(Errors.Num);
            }
        }

        private static CellValue FnDateValue(FunctionNode n, EvalContext ctx)
        {
            string s; CellValue err = CellValue.Empty;
            if (WrongArgs(n, 1, 1) || !GetText(n, 0, ctx, out s, out err)) return err.IsError ? err : CellValue.FromError(Errors.Value);
            DateTime dt;
            string[] formats = new string[] { "d/M/yyyy", "dd/MM/yyyy", "d-M-yyyy", "yyyy-M-d", "yyyy/MM/dd", "d/M/yy" };
            if (DateTime.TryParseExact(s, formats, CultureInfo.GetCultureInfo("es-CO"), DateTimeStyles.None, out dt) ||
                DateTime.TryParse(s, CultureInfo.GetCultureInfo("es-CO"), DateTimeStyles.None, out dt) ||
                DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                return CellValue.FromNumber(dt.Date.ToOADate());
            return CellValue.FromError(Errors.Value);
        }

        private static CellValue FnEDate(FunctionNode n, EvalContext ctx)
        {
            double serial, months; CellValue err = CellValue.Empty;
            if (WrongArgs(n, 2, 2) || !GetNumber(n, 0, ctx, out serial, out err) || !GetNumber(n, 1, ctx, out months, out err))
                return err.IsError ? err : CellValue.FromError(Errors.Value);
            DateTime dt;
            if (!TryDate(serial, out dt)) return CellValue.FromError(Errors.Num);
            try { return CellValue.FromNumber(dt.AddMonths((int)Math.Truncate(months)).ToOADate()); }
            catch { return CellValue.FromError(Errors.Num); }
        }

        private static CellValue FnEOMonth(FunctionNode n, EvalContext ctx)
        {
            double serial, months; CellValue err = CellValue.Empty;
            if (WrongArgs(n, 2, 2) || !GetNumber(n, 0, ctx, out serial, out err) || !GetNumber(n, 1, ctx, out months, out err))
                return err.IsError ? err : CellValue.FromError(Errors.Value);
            DateTime dt;
            if (!TryDate(serial, out dt)) return CellValue.FromError(Errors.Num);
            try
            {
                DateTime first = new DateTime(dt.Year, dt.Month, 1).AddMonths((int)Math.Truncate(months) + 1).AddDays(-1);
                return CellValue.FromNumber(first.ToOADate());
            }
            catch { return CellValue.FromError(Errors.Num); }
        }

        private static bool ValuesEqual(CellValue a, CellValue b)
        {
            if (a.IsError || b.IsError) return false;
            return EvalUtil.Compare(a, b) == 0;
        }

        private static CellValue FnXLookup(FunctionNode n, EvalContext ctx)
        {
            if (WrongArgs(n, 3, 6)) return CellValue.FromError(Errors.Value);
            CellValue lookup = Arg(n, 0, ctx);
            if (lookup.IsError) return lookup;
            CellValue lookupRange = Arg(n, 1, ctx);
            CellValue returnRange = Arg(n, 2, ctx);
            if (lookupRange.Kind != ValueKind.Range || returnRange.Kind != ValueKind.Range)
                return CellValue.FromError(Errors.Value);
            RangeRef lr = ClampRange(lookupRange.Range);
            RangeRef rr = ClampRange(returnRange.Range);
            bool vertical;
            int count;
            if (lr.Cols == 1 && (lr.Rows > 1 || rr.Rows > 1))
            {
                vertical = true; count = lr.Rows;
                if (rr.Rows != count || rr.Cols != 1) return CellValue.FromError(Errors.Value);
            }
            else if (lr.Rows == 1)
            {
                vertical = false; count = lr.Cols;
                if (rr.Cols != count || rr.Rows != 1) return CellValue.FromError(Errors.Value);
            }
            else return CellValue.FromError(Errors.Value);

            int matchMode = 0, searchMode = 1;
            CellValue err = CellValue.Empty;
            if (n.Args.Count >= 5)
            {
                double d;
                if (!GetNumber(n, 4, ctx, out d, out err)) return err;
                matchMode = (int)Math.Truncate(d);
            }
            if (n.Args.Count >= 6)
            {
                double d;
                if (!GetNumber(n, 5, ctx, out d, out err)) return err;
                searchMode = (int)Math.Truncate(d);
            }
            int best = -1;
            bool exact = false;
            int start = searchMode < 0 ? count - 1 : 0;
            int step = searchMode < 0 ? -1 : 1;
            for (int k = 0; k < count; k++)
            {
                int i = start + k * step;
                CellValue key = vertical ? ctx.GetCell(lr.Sheet, lr.Row1 + i, lr.Col1)
                                         : ctx.GetCell(lr.Sheet, lr.Row1, lr.Col1 + i);
                if (key.IsError) continue;
                int cmp = EvalUtil.Compare(key, lookup);
                if (cmp == 0) { best = i; exact = true; break; }
                if (matchMode == -1 && cmp < 0 && (best < 0 || cmp >= 0)) best = i;
                else if (matchMode == 1 && cmp > 0 && (best < 0 || cmp <= 0)) best = i;
            }
            if (best < 0)
            {
                if (n.Args.Count >= 4) return Arg(n, 3, ctx);
                return CellValue.FromError(Errors.NA);
            }
            if (!exact && matchMode == 0)
            {
                if (n.Args.Count >= 4) return Arg(n, 3, ctx);
                return CellValue.FromError(Errors.NA);
            }
            return vertical ? ctx.GetCell(rr.Sheet, rr.Row1 + best, rr.Col1)
                            : ctx.GetCell(rr.Sheet, rr.Row1, rr.Col1 + best);
        }

        private static CellValue FnXMatch(FunctionNode n, EvalContext ctx)
        {
            if (WrongArgs(n, 2, 4)) return CellValue.FromError(Errors.Value);
            CellValue lookup = Arg(n, 0, ctx);
            if (lookup.IsError) return lookup;
            CellValue rangeValue = Arg(n, 1, ctx);
            if (rangeValue.Kind != ValueKind.Range) return CellValue.FromError(Errors.Value);
            RangeRef r = ClampRange(rangeValue.Range);
            bool vertical = r.Cols == 1;
            if (!vertical && r.Rows != 1) return CellValue.FromError(Errors.Value);
            int count = vertical ? r.Rows : r.Cols;
            int matchMode = 0, searchMode = 1;
            CellValue err = CellValue.Empty;
            if (n.Args.Count >= 3)
            {
                double d;
                if (!GetNumber(n, 2, ctx, out d, out err)) return err;
                matchMode = (int)Math.Truncate(d);
            }
            if (n.Args.Count >= 4)
            {
                double d;
                if (!GetNumber(n, 3, ctx, out d, out err)) return err;
                searchMode = (int)Math.Truncate(d);
            }
            int best = -1;
            bool exact = false;
            int start = searchMode < 0 ? count - 1 : 0;
            int step = searchMode < 0 ? -1 : 1;
            for (int k = 0; k < count; k++)
            {
                int i = start + k * step;
                CellValue key = vertical ? ctx.GetCell(r.Sheet, r.Row1 + i, r.Col1)
                                         : ctx.GetCell(r.Sheet, r.Row1, r.Col1 + i);
                if (key.IsError) continue;
                int cmp = EvalUtil.Compare(key, lookup);
                if (cmp == 0) { best = i; exact = true; break; }
                if (matchMode == -1 && cmp < 0) best = i;
                else if (matchMode == 1 && cmp > 0 && (best < 0 || cmp <= 0)) best = i;
            }
            if (best < 0 || (!exact && matchMode == 0)) return CellValue.FromError(Errors.NA);
            return CellValue.FromNumber(best + 1);
        }

        private static CellValue FnVLookup(FunctionNode n, EvalContext ctx, bool vertical)
        {
            if (WrongArgs(n, 3, 4)) return CellValue.FromError(Errors.Value);
            CellValue lookup = Arg(n, 0, ctx);
            if (lookup.IsError) return lookup;
            CellValue tableValue = Arg(n, 1, ctx);
            if (tableValue.IsError) return tableValue;
            if (tableValue.Kind != ValueKind.Range) return CellValue.FromError(Errors.Value);
            double indexDouble; CellValue err = CellValue.Empty;
            if (!GetNumber(n, 2, ctx, out indexDouble, out err)) return err;
            int index = (int)Math.Truncate(indexDouble);
            if (index < 1) return CellValue.FromError(Errors.Value);
            bool approximate = true;
            if (n.Args.Count >= 4)
            {
                bool b;
                if (!GetBool(n, 3, ctx, out b, out err)) return err;
                approximate = b;
            }
            RangeRef r = ClampRange(tableValue.Range);
            int limit = vertical ? r.Rows : r.Cols;
            int indexLimit = vertical ? r.Cols : r.Rows;
            if (index > indexLimit) return CellValue.FromError(Errors.Ref);
            int best = -1;
            for (int k = 0; k < limit; k++)
            {
                CellValue key = vertical ? ctx.GetCell(r.Sheet, r.Row1 + k, r.Col1)
                                         : ctx.GetCell(r.Sheet, r.Row1, r.Col1 + k);
                if (key.IsError) return key;
                if (approximate)
                {
                    int cmp = EvalUtil.Compare(key, lookup);
                    if (cmp <= 0) best = k;
                    else break;
                }
                else
                {
                    if (ValuesEqual(key, lookup)) { best = k; break; }
                }
            }
            if (best < 0) return CellValue.FromError(Errors.NA);
            if (vertical)
                return ctx.GetCell(r.Sheet, r.Row1 + best, r.Col1 + index - 1);
            return ctx.GetCell(r.Sheet, r.Row1 + index - 1, r.Col1 + best);
        }

        private static CellValue FnIndex(FunctionNode n, EvalContext ctx)
        {
            if (WrongArgs(n, 2, 3)) return CellValue.FromError(Errors.Value);
            CellValue tableValue = Arg(n, 0, ctx);
            if (tableValue.IsError) return tableValue;
            double rowD, colD = 1; CellValue err = CellValue.Empty;
            if (!GetNumber(n, 1, ctx, out rowD, out err)) return err;
            int rowArg = (int)Math.Truncate(rowD);
            int colArg = 1;
            if (n.Args.Count >= 3 && !GetNumber(n, 2, ctx, out colD, out err)) return err;
            colArg = (int)Math.Truncate(colD);
            if (tableValue.Kind == ValueKind.Array)
            {
                List<CellValue> array = tableValue.Values ?? new List<CellValue>();
                int rows = tableValue.ArrayRows > 0 ? tableValue.ArrayRows : 1;
                int cols = tableValue.ArrayCols > 0 ? tableValue.ArrayCols : Math.Max(1, array.Count);
                if (n.Args.Count < 3)
                {
                    if (cols > 1 && rows == 1) { colArg = rowArg; rowArg = 1; }
                    else colArg = 1;
                }
                if (rowArg < 1 || colArg < 1 || rowArg > rows || colArg > cols)
                    return CellValue.FromError(Errors.Ref);
                int index = (rowArg - 1) * cols + (colArg - 1);
                if (index < 0 || index >= array.Count) return CellValue.FromError(Errors.Ref);
                return array[index];
            }
            if (tableValue.Kind != ValueKind.Range) return CellValue.FromError(Errors.Value);
            RangeRef r = ClampRange(tableValue.Range);
            if (n.Args.Count < 3)
            {
                if (r.Rows == 1 && r.Cols > 1) { colArg = rowArg; rowArg = 1; }
                else colArg = 1;
            }
            if (rowArg < 1 || colArg < 1) return CellValue.FromError(Errors.Value);
            if (rowArg > r.Rows || colArg > r.Cols) return CellValue.FromError(Errors.Ref);
            return ctx.GetCell(r.Sheet, r.Row1 + rowArg - 1, r.Col1 + colArg - 1);
        }

        private static CellValue FnMatch(FunctionNode n, EvalContext ctx)
        {
            if (WrongArgs(n, 2, 3)) return CellValue.FromError(Errors.Value);
            CellValue lookup = Arg(n, 0, ctx);
            if (lookup.IsError) return lookup;
            CellValue rangeValue = Arg(n, 1, ctx);
            if (rangeValue.IsError) return rangeValue;
            if (rangeValue.Kind != ValueKind.Range) return CellValue.FromError(Errors.Value);
            double typeD = 1; CellValue err = CellValue.Empty;
            if (n.Args.Count >= 3 && !GetNumber(n, 2, ctx, out typeD, out err)) return err;
            int type = (int)Math.Truncate(typeD);
            RangeRef r = ClampRange(rangeValue.Range);
            bool vertical = r.Cols == 1;
            if (!vertical && r.Rows != 1) return CellValue.FromError(Errors.NA);
            int limit = vertical ? r.Rows : r.Cols;
            int best = -1;
            for (int k = 0; k < limit; k++)
            {
                CellValue v = vertical ? ctx.GetCell(r.Sheet, r.Row1 + k, r.Col1)
                                       : ctx.GetCell(r.Sheet, r.Row1, r.Col1 + k);
                if (v.IsError) continue;
                if (type == 0)
                {
                    if (ValuesEqual(v, lookup)) return CellValue.FromNumber(k + 1);
                }
                else if (type > 0)
                {
                    if (EvalUtil.Compare(v, lookup) <= 0) best = k;
                    else break;
                }
                else
                {
                    if (EvalUtil.Compare(v, lookup) >= 0) best = k;
                    else break;
                }
            }
            if (type != 0 && best >= 0) return CellValue.FromNumber(best + 1);
            return CellValue.FromError(Errors.NA);
        }

        private sealed class Criteria
        {
            private readonly string _op;
            private readonly bool _numeric;
            private readonly double _number;
            private readonly string _text;
            private readonly bool _blank;
            private readonly Regex _regex;

            public Criteria(CellValue value)
            {
                _op = "=";
                _numeric = false; _number = 0; _text = ""; _blank = false;
                if (value.IsError) return;
                if (value.Kind == ValueKind.Number)
                {
                    _numeric = true; _number = value.Number;
                    return;
                }
                if (value.Kind == ValueKind.Bool)
                {
                    _text = value.Bool ? "VERDADERO" : "FALSO";
                    return;
                }
                string s = value.Kind == ValueKind.Empty ? "" : (value.Text ?? "");
                if (s.Length > 0 && (s[0] == '=' || s[0] == '<' || s[0] == '>'))
                {
                    if (s.StartsWith(">=", StringComparison.Ordinal) || s.StartsWith("<=", StringComparison.Ordinal) ||
                        s.StartsWith("<>", StringComparison.Ordinal))
                    {
                        _op = s.Substring(0, 2);
                        s = s.Substring(2);
                    }
                    else
                    {
                        _op = s.Substring(0, 1);
                        s = s.Substring(1);
                    }
                }
                s = s.Trim();
                if (s.Length == 0)
                {
                    _blank = true;
                    _text = "";
                }
                else
                {
                    double n;
                    if (NumberParser.TryParse(s, out n) && s.IndexOf('*') < 0 && s.IndexOf('?') < 0)
                    {
                        _numeric = true; _number = n;
                    }
                    else
                    {
                        _text = s;
                        if (s.IndexOf('*') >= 0 || s.IndexOf('?') >= 0)
                            if (s.Length <= 512)
                            {
                                _regex = new Regex("^" + WildcardToRegex(s) + "$",
                                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant,
                                    TimeSpan.FromMilliseconds(250));
                            }
                    }
                }
            }

            private static string WildcardToRegex(string s)
            {
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < s.Length; i++)
                {
                    char c = s[i];
                    if (c == '~' && i + 1 < s.Length) { i++; sb.Append(Regex.Escape(s[i].ToString())); }
                    else if (c == '*') sb.Append(".*");
                    else if (c == '?') sb.Append(".");
                    else sb.Append(Regex.Escape(c.ToString()));
                }
                return sb.ToString();
            }

            public bool Matches(CellValue v)
            {
                if (v.IsError) return false;
                if (_blank)
                {
                    bool blank = v.Kind == ValueKind.Empty || (v.Kind == ValueKind.Text && (v.Text ?? "").Length == 0);
                    return _op == "<>" ? !blank : blank;
                }
                if (_numeric)
                {
                    double d;
                    if (v.Kind == ValueKind.Number) d = v.Number;
                    else if (v.Kind == ValueKind.Empty) d = 0;
                    else if (v.Kind == ValueKind.Bool) d = v.Bool ? 1 : 0;
                    else return _op == "<>";
                    int cmp = d.CompareTo(_number);
                    return Apply(cmp);
                }
                string text;
                if (v.Kind == ValueKind.Text) text = v.Text ?? "";
                else if (v.Kind == ValueKind.Empty) text = "";
                else if (v.Kind == ValueKind.Number) text = NumberFormatter.FormatGeneral(v.Number);
                else text = v.Bool ? "VERDADERO" : "FALSO";
                if (_regex != null)
                {
                    try
                    {
                        bool m = _regex.IsMatch(text);
                        return _op == "<>" ? !m : m;
                    }
                    catch (RegexMatchTimeoutException)
                    {
                        return false;
                    }
                }
                int c = string.Compare(text, _text, StringComparison.CurrentCultureIgnoreCase);
                return Apply(c);
            }

            private bool Apply(int cmp)
            {
                switch (_op)
                {
                    case "=": return cmp == 0;
                    case "<>": return cmp != 0;
                    case "<": return cmp < 0;
                    case ">": return cmp > 0;
                    case "<=": return cmp <= 0;
                    default: return cmp >= 0;
                }
            }
        }
    }

    public sealed class FunctionInfo
    {
        public readonly string Name;
        public readonly string Signature;
        public readonly string Category;
        public readonly string Description;

        public FunctionInfo(string name, string signature, string category, string description)
        {
            Name = name; Signature = signature; Category = category; Description = description;
        }

        public override string ToString()
        {
            return Name + "   " + Signature;
        }
    }

    public static class FunctionCatalog
    {
        private static readonly List<FunctionInfo> _all = Build();

        public static List<FunctionInfo> All { get { return _all; } }

        private static void Add(List<FunctionInfo> list, string name, string sig, string cat, string desc)
        {
            list.Add(new FunctionInfo(name, sig, cat, desc));
        }

        private static List<FunctionInfo> Build()
        {
            List<FunctionInfo> l = new List<FunctionInfo>();
            Add(l, "SUMA", "SUMA(número1; [número2]; ...)", "Matemáticas", "Suma todos los números de un rango o lista.");
            Add(l, "PROMEDIO", "PROMEDIO(número1; [número2]; ...)", "Estadística", "Calcula la media aritmética de los números.");
            Add(l, "MIN", "MIN(número1; [número2]; ...)", "Estadística", "Devuelve el valor mínimo.");
            Add(l, "MAX", "MAX(número1; [número2]; ...)", "Estadística", "Devuelve el valor máximo.");
            Add(l, "CONTAR", "CONTAR(valor1; [valor2]; ...)", "Estadística", "Cuenta las celdas que contienen números.");
            Add(l, "CONTARA", "CONTARA(valor1; [valor2]; ...)", "Estadística", "Cuenta las celdas no vacías.");
            Add(l, "SI", "SI(prueba; valor_si_verdadero; [valor_si_falso])", "Lógica", "Devuelve un valor u otro según una condición.");
            Add(l, "Y", "Y(condición1; [condición2]; ...)", "Lógica", "Devuelve VERDADERO si todas las condiciones son verdaderas.");
            Add(l, "O", "O(condición1; [condición2]; ...)", "Lógica", "Devuelve VERDADERO si alguna condición es verdadera.");
            Add(l, "SI.ERROR", "SI.ERROR(valor; valor_si_error)", "Lógica", "Devuelve una alternativa cuando hay un error.");
            Add(l, "REDONDEAR", "REDONDEAR(número; núm_decimales)", "Matemáticas", "Redondea un número al número de decimales indicado.");
            Add(l, "ENTERO", "ENTERO(número)", "Matemáticas", "Redondea un número hacia abajo hasta el entero más próximo.");
            Add(l, "RESIDUO", "RESIDUO(número; divisor)", "Matemáticas", "Devuelve el resto de una división.");
            Add(l, "RAIZ", "RAIZ(número)", "Matemáticas", "Devuelve la raíz cuadrada positiva.");
            Add(l, "POTENCIA", "POTENCIA(número; exponente)", "Matemáticas", "Eleva un número a una potencia.");
            Add(l, "SUMAPRODUCTO", "SUMAPRODUCTO(matriz1; [matriz2]; ...)", "Matemáticas", "Multiplica los elementos correspondientes y suma los productos.");
            Add(l, "SUMAR.SI", "SUMAR.SI(rango; criterio; [rango_suma])", "Matemáticas", "Suma las celdas que cumplen un criterio.");
            Add(l, "CONTAR.SI", "CONTAR.SI(rango; criterio)", "Estadística", "Cuenta las celdas que cumplen un criterio.");
            Add(l, "BUSCARV", "BUSCARV(valor; tabla; índice_col; [aproximado])", "Búsqueda", "Busca un valor en la primera columna de una tabla.");
            Add(l, "BUSCARH", "BUSCARH(valor; tabla; índice_fila; [aproximado])", "Búsqueda", "Busca un valor en la primera fila de una tabla.");
            Add(l, "INDICE", "INDICE(matriz; fila; [columna])", "Búsqueda", "Devuelve el valor de una posición de una matriz.");
            Add(l, "COINCIDIR", "COINCIDIR(valor; matriz; [tipo])", "Búsqueda", "Devuelve la posición de un valor en una matriz.");
            Add(l, "ELEGIR", "ELEGIR(índice; valor1; [valor2]; ...)", "Búsqueda", "Elige un valor de una lista según su posición.");
            Add(l, "LARGO", "LARGO(texto)", "Texto", "Devuelve el número de caracteres de un texto.");
            Add(l, "IZQUIERDA", "IZQUIERDA(texto; [núm_caracteres])", "Texto", "Devuelve los primeros caracteres de un texto.");
            Add(l, "DERECHA", "DERECHA(texto; [núm_caracteres])", "Texto", "Devuelve los últimos caracteres de un texto.");
            Add(l, "EXTRAE", "EXTRAE(texto; posición; núm_caracteres)", "Texto", "Devuelve caracteres de un texto a partir de una posición.");
            Add(l, "MAYUSC", "MAYUSC(texto)", "Texto", "Convierte el texto a mayúsculas.");
            Add(l, "MINUSC", "MINUSC(texto)", "Texto", "Convierte el texto a minúsculas.");
            Add(l, "CONCATENAR", "CONCATENAR(texto1; [texto2]; ...)", "Texto", "Une varios textos en uno solo.");
            Add(l, "TEXTO", "TEXTO(valor; formato)", "Texto", "Convierte un valor en texto con formato.");
            Add(l, "HOY", "HOY()", "Fecha", "Devuelve la fecha actual.");
            Add(l, "AHORA", "AHORA()", "Fecha", "Devuelve la fecha y hora actuales.");
            Add(l, "FECHA", "FECHA(año; mes; día)", "Fecha", "Construye una fecha a partir de año, mes y día.");
            Add(l, "AÑO", "AÑO(fecha)", "Fecha", "Extrae el año de una fecha.");
            Add(l, "MES", "MES(fecha)", "Fecha", "Extrae el mes de una fecha.");
            Add(l, "DIA", "DIA(fecha)", "Fecha", "Extrae el día de una fecha.");
            return l;
        }
    }
}
