using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace CalcRapido
{
    internal enum TokenKind
    {
        Number, Text, Ident, QuotedSheet, ErrorValue, Operator,
        LParen, RParen, Comma, Colon, Bang, Percent, End, Array
    }

    internal sealed class Token
    {
        public TokenKind Kind;
        public string Text;
        public double Number;
        public int Pos;

        public Token(TokenKind kind, string text, int pos)
        {
            Kind = kind; Text = text; Pos = pos;
        }
    }

    public sealed class FormulaParseException : Exception
    {
        public int Position;
        public FormulaParseException(string message, int position) : base(message)
        {
            Position = position;
        }
    }

    internal sealed class FormulaLexer
    {
        private readonly string _s;
        private readonly bool _hasSemicolon;
        private int _i;
        private int _parenDepth;

        public FormulaLexer(string formula)
        {
            _s = formula ?? "";
            _hasSemicolon = ContainsArgumentSemicolon(_s);
        }

        private static bool ContainsArgumentSemicolon(string s)
        {
            bool dquote = false;
            bool squote = false;
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (c == '"' && !squote)
                {
                    if (dquote && i + 1 < s.Length && s[i + 1] == '"') { i++; continue; }
                    dquote = !dquote;
                }
                else if (c == '\'' && !dquote)
                {
                    if (squote && i + 1 < s.Length && s[i + 1] == '\'') { i++; continue; }
                    squote = !squote;
                }
                else if (c == ';' && !dquote && !squote) return true;
            }
            return false;
        }

        private char Cur { get { return _i < _s.Length ? _s[_i] : '\0'; } }
        private char Next { get { return _i + 1 < _s.Length ? _s[_i + 1] : '\0'; } }

        public List<Token> Tokenize()
        {
            List<Token> list = new List<Token>();
            while (true)
            {
                while (_i < _s.Length && char.IsWhiteSpace(_s[_i])) _i++;
                if (_i >= _s.Length)
                {
                    list.Add(new Token(TokenKind.End, "", _i));
                    break;
                }
                int start = _i;
                char ch = _s[_i];

                if (ch == '{')
                {
                    _i++;
                    int contentStart = _i;
                    bool quoted = false;
                    int close = -1;
                    while (_i < _s.Length)
                    {
                        if (_s[_i] == '"')
                        {
                            if (quoted && _i + 1 < _s.Length && _s[_i + 1] == '"') { _i += 2; continue; }
                            quoted = !quoted;
                        }
                        else if (_s[_i] == '}' && !quoted) { close = _i; break; }
                        _i++;
                    }
                    if (close < 0) throw new FormulaParseException("Falta cerrar una constante matricial.", start);
                    list.Add(new Token(TokenKind.Array,
                        _s.Substring(contentStart, close - contentStart), start));
                    _i = close + 1;
                    continue;
                }

                if (ch == '"')
                {
                    _i++;
                    StringBuilder sb = new StringBuilder();
                    bool closed = false;
                    while (_i < _s.Length)
                    {
                        if (_s[_i] == '"')
                        {
                            if (_i + 1 < _s.Length && _s[_i + 1] == '"')
                            {
                                sb.Append('"');
                                _i += 2;
                                continue;
                            }
                            _i++;
                            closed = true;
                            break;
                        }
                        sb.Append(_s[_i]);
                        _i++;
                    }
                    if (!closed) throw new FormulaParseException("Falta cerrar una cadena de texto.", start);
                    Token t = new Token(TokenKind.Text, sb.ToString(), start);
                    list.Add(t);
                    continue;
                }

                if (ch == '\'')
                {
                    _i++;
                    StringBuilder sb = new StringBuilder();
                    bool closed = false;
                    while (_i < _s.Length)
                    {
                        if (_s[_i] == '\'')
                        {
                            if (_i + 1 < _s.Length && _s[_i + 1] == '\'')
                            {
                                sb.Append('\'');
                                _i += 2;
                                continue;
                            }
                            _i++;
                            closed = true;
                            break;
                        }
                        sb.Append(_s[_i]);
                        _i++;
                    }
                    if (!closed) throw new FormulaParseException("Falta cerrar el nombre de la hoja.", start);
                    list.Add(new Token(TokenKind.QuotedSheet, sb.ToString(), start));
                    continue;
                }

                if (ch == '#')
                {
                    _i++;
                    while (_i < _s.Length)
                    {
                        char c = _s[_i];
                        if (char.IsWhiteSpace(c) || c == ',' || c == ';' || c == '(' || c == ')' ||
                            c == '+' || c == '-' || c == '*' || c == '^' || c == '&' || c == '=' ||
                            c == '<' || c == '>')
                            break;
                        _i++;
                    }
                    string raw = _s.Substring(start, _i - start);
                    list.Add(new Token(TokenKind.ErrorValue, NormalizeError(raw), start));
                    continue;
                }

                if (char.IsDigit(ch) || (ch == '.' && char.IsDigit(Next)))
                {
                    _i = start;
                    while (_i < _s.Length && char.IsDigit(_s[_i])) _i++;
                    if (_i < _s.Length && _s[_i] == '.' && _i + 1 < _s.Length && char.IsDigit(_s[_i + 1]))
                    {
                        _i++;
                        while (_i < _s.Length && char.IsDigit(_s[_i])) _i++;
                    }
                    else if ((_parenDepth == 0 || _hasSemicolon) &&
                             _i < _s.Length && _s[_i] == ',' && _i + 1 < _s.Length && char.IsDigit(_s[_i + 1]))
                    {
                        _i++;
                        while (_i < _s.Length && char.IsDigit(_s[_i])) _i++;
                    }
                    if (_i < _s.Length && (_s[_i] == 'e' || _s[_i] == 'E'))
                    {
                        int save = _i;
                        int j = _i + 1;
                        if (j < _s.Length && (_s[j] == '+' || _s[j] == '-')) j++;
                        if (j < _s.Length && char.IsDigit(_s[j]))
                        {
                            _i = j;
                            while (_i < _s.Length && char.IsDigit(_s[_i])) _i++;
                        }
                        else
                        {
                            _i = save;
                        }
                    }
                    string raw = _s.Substring(start, _i - start).Replace(',', '.');
                    double d;
                    if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out d))
                        throw new FormulaParseException("Número no válido: " + raw, start);
                    Token tok = new Token(TokenKind.Number, raw, start);
                    tok.Number = d;
                    list.Add(tok);
                    continue;
                }

                if (char.IsLetter(ch) || ch == '_' || ch == '$')
                {
                    _i++;
                    while (_i < _s.Length)
                    {
                        char c = _s[_i];
                        if (char.IsLetterOrDigit(c) || c == '_' || c == '.' || c == '$' || c > 127) _i++;
                        else break;
                    }
                    string word = _s.Substring(start, _i - start);
                    list.Add(new Token(TokenKind.Ident, word, start));
                    continue;
                }

                if (ch == '(') { _i++; _parenDepth++; list.Add(new Token(TokenKind.LParen, "(", start)); continue; }
                if (ch == ')')
                {
                    _i++;
                    if (_parenDepth > 0) _parenDepth--;
                    list.Add(new Token(TokenKind.RParen, ")", start));
                    continue;
                }
                if (ch == ',') { _i++; list.Add(new Token(TokenKind.Comma, ",", start)); continue; }
                if (ch == ';') { _i++; list.Add(new Token(TokenKind.Comma, ";", start)); continue; }
                if (ch == ':') { _i++; list.Add(new Token(TokenKind.Colon, ":", start)); continue; }
                if (ch == '!') { _i++; list.Add(new Token(TokenKind.Bang, "!", start)); continue; }
                if (ch == '%') { _i++; list.Add(new Token(TokenKind.Percent, "%", start)); continue; }

                if (ch == '<' || ch == '>' || ch == '=')
                {
                    if (_i + 1 < _s.Length)
                    {
                        char n = _s[_i + 1];
                        if ((ch == '<' && n == '>') || (ch == '<' && n == '=') || (ch == '>' && n == '='))
                        {
                            _i += 2;
                            list.Add(new Token(TokenKind.Operator, _s.Substring(start, 2), start));
                            continue;
                        }
                    }
                    _i++;
                    list.Add(new Token(TokenKind.Operator, ch.ToString(), start));
                    continue;
                }

                if (ch == '+' || ch == '-' || ch == '*' || ch == '/' || ch == '^' || ch == '&')
                {
                    _i++;
                    list.Add(new Token(TokenKind.Operator, ch.ToString(), start));
                    continue;
                }

                throw new FormulaParseException("Carácter no válido: " + ch, start);
            }
            return list;
        }

        private static string NormalizeError(string raw)
        {
            string u = raw.ToUpperInvariant().Replace("\u00A1", "").Replace("\u00BF", "");
            if (u.Contains("DIV/0")) return Errors.Div0;
            if (u.Contains("VALOR") || u.Contains("VALUE")) return Errors.Value;
            if (u.Contains("NOMBRE") || u.Contains("NAME")) return Errors.Name;
            if (u == "#REF" || u.Contains("REF!")) return Errors.Ref;
            if (u.Contains("N/A")) return Errors.NA;
            if (u.Contains("NUM")) return Errors.Num;
            if (u.Contains("CIRC")) return Errors.Circular;
            return Errors.Error;
        }
    }

    internal sealed class RefAtom
    {
        public string SheetName;
        public bool IsCell;
        public bool IsColumn;
        public bool IsRow;
        public int Row;
        public int Col;
        public bool AbsRow;
        public bool AbsCol;
    }

    internal sealed class FormulaParser
    {
        private readonly List<Token> _t;
        private int _i;

        private int _depth;
        private const int MaxNesting = 128;

        public FormulaParser(List<Token> tokens)
        {
            _t = tokens;
        }

        private Token Cur { get { return _t[_i]; } }
        private Token Peek(int n)
        {
            int k = _i + n;
            return k < _t.Count ? _t[k] : _t[_t.Count - 1];
        }

        private Token Advance()
        {
            Token t = _t[_i];
            if (_i < _t.Count - 1) _i++;
            return t;
        }

        private bool IsOp(string op)
        {
            return Cur.Kind == TokenKind.Operator && Cur.Text == op;
        }

        private void Expect(TokenKind kind, string message)
        {
            if (Cur.Kind != kind) throw new FormulaParseException(message, Cur.Pos);
            Advance();
        }

        public FormulaNode ParseFormula()
        {
            if (Cur.Kind == TokenKind.End) throw new FormulaParseException("La fórmula está vacía.", 0);
            FormulaNode n = ParseExpression();
            if (Cur.Kind != TokenKind.End)
                throw new FormulaParseException("Sobra contenido en la fórmula: " + Cur.Text, Cur.Pos);
            return n;
        }

        private FormulaNode ParseExpression()
        {
            if (++_depth > MaxNesting)
                throw new FormulaParseException("La fórmula tiene demasiados niveles de anidamiento.", Cur.Pos);
            try
            {
                return ParseComparison();
            }
            finally
            {
                _depth--;
            }
        }

        private FormulaNode ParseComparison()
        {
            FormulaNode left = ParseConcat();
            while (Cur.Kind == TokenKind.Operator &&
                   (Cur.Text == "=" || Cur.Text == "<>" || Cur.Text == "<" || Cur.Text == ">" ||
                    Cur.Text == "<=" || Cur.Text == ">="))
            {
                string op = Advance().Text;
                FormulaNode right = ParseConcat();
                left = new BinaryNode(op, left, right);
            }
            return left;
        }

        private FormulaNode ParseConcat()
        {
            FormulaNode left = ParseAdd();
            while (IsOp("&"))
            {
                Advance();
                FormulaNode right = ParseAdd();
                left = new BinaryNode("&", left, right);
            }
            return left;
        }

        private FormulaNode ParseAdd()
        {
            FormulaNode left = ParseMul();
            while (IsOp("+") || IsOp("-"))
            {
                string op = Advance().Text;
                FormulaNode right = ParseMul();
                left = new BinaryNode(op, left, right);
            }
            return left;
        }

        private FormulaNode ParseMul()
        {
            FormulaNode left = ParseUnary();
            while (IsOp("*") || IsOp("/"))
            {
                string op = Advance().Text;
                FormulaNode right = ParseUnary();
                left = new BinaryNode(op, left, right);
            }
            return left;
        }

        private FormulaNode ParseUnary()
        {
            if (IsOp("+") || IsOp("-"))
            {
                string op = Advance().Text;
                if (IsOp("+") || IsOp("-")) return new UnaryNode(op, ParseUnary());
                FormulaNode left = new UnaryNode(op, ParsePostfix());
                if (IsOp("^"))
                {
                    Advance();
                    return new BinaryNode("^", left, ParseUnary());
                }
                return left;
            }
            return ParsePower();
        }

        private FormulaNode ParsePower()
        {
            FormulaNode left = ParsePostfix();
            if (IsOp("^"))
            {
                Advance();
                FormulaNode right = ParseUnary();
                return new BinaryNode("^", left, right);
            }
            return left;
        }

        private FormulaNode ParsePostfix()
        {
            FormulaNode node = ParsePrimary();
            while (Cur.Kind == TokenKind.Percent)
            {
                Advance();
                node = new PercentNode(node);
            }
            return node;
        }

        private FormulaNode ParsePrimary()
        {
            Token t = Cur;
            if (t.Kind == TokenKind.Number)
            {
                if (Peek(1).Kind == TokenKind.Colon && Peek(2).Kind == TokenKind.Number)
                {
                    int r1 = (int)t.Number;
                    int r2 = (int)Peek(2).Number;
                    if (r1 < 1 || r2 < 1 || r1 > Limits.MaxRows || r2 > Limits.MaxRows ||
                        r1 != t.Number || r2 != Peek(2).Number)
                        throw new FormulaParseException("Fila no válida.", t.Pos);
                    Advance(); Advance(); Advance();
                    return new RangeNode(null, r1 - 1, 0, r2 - 1, Limits.MaxCols - 1, null);
                }
                Advance();
                return new NumberNode(t.Number);
            }
            if (t.Kind == TokenKind.Array)
            {
                Advance();
                return new ArrayNode(t.Text);
            }
            if (t.Kind == TokenKind.Text)
            {
                Advance();
                return new TextNode(t.Text);
            }
            if (t.Kind == TokenKind.ErrorValue)
            {
                Advance();
                return new ErrorNode(t.Text);
            }
            if (t.Kind == TokenKind.LParen)
            {
                Advance();
                FormulaNode n = ParseExpression();
                Expect(TokenKind.RParen, "Falta cerrar un paréntesis.");
                return n;
            }
            if (t.Kind == TokenKind.Ident)
            {
                string upper = t.Text.ToUpperInvariant();
                if ((upper == "TRUE" || upper == "FALSE" || upper == "VERDADERO" || upper == "FALSO") &&
                    Peek(1).Kind != TokenKind.LParen)
                {
                    Advance();
                    return new BoolNode(upper == "TRUE" || upper == "VERDADERO");
                }
                if (Peek(1).Kind == TokenKind.LParen)
                    return ParseFunction();
            }
            if (t.Kind == TokenKind.QuotedSheet)
            {
                RefAtom atom;
                ParseAtom(out atom);
                if (Cur.Kind == TokenKind.Colon)
                {
                    Advance();
                    RefAtom end;
                    ParseAtomWithOptionalSheet(out end);
                    return MakeRange(atom, end);
                }
                if (atom.IsCell)
                    return new RefNode(atom.SheetName, atom.Row, atom.Col, atom.AbsRow, atom.AbsCol);
                throw new FormulaParseException("Referencia no válida.", t.Pos);
            }
            if (t.Kind == TokenKind.Ident)
            {
                RefAtom atom;
                RefAtom endAtom;
                int save = _i;
                try
                {
                    ParseAtom(out atom);
                }
                catch (FormulaParseException)
                {
                    _i = save;
                    Advance();
                    return new NameNode(t.Text);
                }
                if (Cur.Kind == TokenKind.Colon)
                {
                    Advance();
                    ParseAtomWithOptionalSheet(out endAtom);
                    return MakeRange(atom, endAtom);
                }
                if (atom.IsCell)
                    return new RefNode(atom.SheetName, atom.Row, atom.Col, atom.AbsRow, atom.AbsCol);
                _i = save;
                Advance();
                return new NameNode(t.Text);
            }
            if (t.Kind == TokenKind.End)
                throw new FormulaParseException("Fórmula incompleta.", t.Pos);
            throw new FormulaParseException("Se esperaba un valor y se encontró: " + t.Text, t.Pos);
        }

        private FormulaNode MakeRange(RefAtom a, RefAtom b)
        {
            string endSheet = b.SheetName;
            if (a.IsCell && b.IsCell)
                return new RangeNode(a.SheetName, a.Row, a.Col, b.Row, b.Col, endSheet);
            if (a.IsColumn && b.IsColumn)
                return new RangeNode(a.SheetName, 0, a.Col, Limits.MaxRows - 1, b.Col, endSheet);
            if (a.IsRow && b.IsRow)
                return new RangeNode(a.SheetName, a.Row, 0, b.Row, Limits.MaxCols - 1, endSheet);
            throw new FormulaParseException("El rango mezcla tipos de referencia no válidos.", 0);
        }

        private void ParseAtom(out RefAtom atom)
        {
            atom = new RefAtom();
            if (Cur.Kind == TokenKind.QuotedSheet ||
                (Cur.Kind == TokenKind.Ident && Peek(1).Kind == TokenKind.Bang))
            {
                atom.SheetName = Cur.Text;
                Advance();
                Expect(TokenKind.Bang, "Falta el signo ! después del nombre de la hoja.");
            }
            Token t = Cur;
            if (t.Kind == TokenKind.Ident)
            {
                int row, col; bool ar, ac;
                string letters = t.Text;
                if (AddressUtil.TryParseCell(letters, out row, out col, out ar, out ac))
                {
                    atom.IsCell = true; atom.Row = row; atom.Col = col; atom.AbsRow = ar; atom.AbsCol = ac;
                    Advance();
                    return;
                }
                if (AddressUtil.TryColumnIndex(letters, out col))
                {
                    atom.IsColumn = true; atom.Col = col;
                    Advance();
                    return;
                }
                throw new FormulaParseException("Referencia no válida: " + t.Text, t.Pos);
            }
            if (t.Kind == TokenKind.Number)
            {
                int row = (int)t.Number;
                if (row < 1 || row > Limits.MaxRows || row != t.Number)
                    throw new FormulaParseException("Fila no válida.", t.Pos);
                atom.IsRow = true; atom.Row = row - 1;
                Advance();
                return;
            }
            throw new FormulaParseException("Se esperaba una referencia de celda.", t.Pos);
        }

        private void ParseAtomWithOptionalSheet(out RefAtom atom)
        {
            ParseAtom(out atom);
        }

        private FormulaNode ParseFunction()
        {
            Token name = Advance();
            Expect(TokenKind.LParen, "Falta el paréntesis de la función.");
            List<FormulaNode> args = new List<FormulaNode>();
            if (Cur.Kind == TokenKind.RParen)
            {
                Advance();
                return new FunctionNode(name.Text, args);
            }
            while (true)
            {
                if (Cur.Kind == TokenKind.Comma || Cur.Kind == TokenKind.RParen)
                {
                    args.Add(new MissingNode());
                }
                else
                {
                    args.Add(ParseExpression());
                }
                if (Cur.Kind == TokenKind.Comma)
                {
                    Advance();
                    if (Cur.Kind == TokenKind.RParen)
                    {
                        args.Add(new MissingNode());
                        Advance();
                        break;
                    }
                    continue;
                }
                if (Cur.Kind == TokenKind.RParen)
                {
                    Advance();
                    break;
                }
                throw new FormulaParseException("Falta una coma o un paréntesis en la función.", Cur.Pos);
            }
            return new FunctionNode(name.Text, args);
        }
    }

    internal abstract class FormulaNode
    {
        internal abstract CellValue Eval(EvalContext ctx);
    }

    internal sealed class NumberNode : FormulaNode
    {
        private readonly double _v;
        public NumberNode(double v) { _v = v; }
        internal override CellValue Eval(EvalContext ctx) { return CellValue.FromNumber(_v); }
    }

    internal sealed class TextNode : FormulaNode
    {
        private readonly string _v;
        public TextNode(string v) { _v = v; }
        internal override CellValue Eval(EvalContext ctx) { return CellValue.FromText(_v); }
    }

    internal sealed class ArrayNode : FormulaNode
    {
        private readonly string _content;
        public ArrayNode(string content) { _content = content ?? ""; }

        internal override CellValue Eval(EvalContext ctx)
        {
            List<string> rowTexts = SplitArray(_content, ';');
            List<CellValue> values = new List<CellValue>();
            int cols = 1;
            foreach (string rowText in rowTexts)
            {
                List<string> cells = SplitArray(rowText, ',');
                cols = Math.Max(cols, cells.Count);
                foreach (string cellText in cells)
                    values.Add(ParseArrayValue(cellText));
            }
            return CellValue.FromArray(values, rowTexts.Count, cols);
        }

        private static List<string> SplitArray(string text, char separator)
        {
            List<string> parts = new List<string>();
            StringBuilder current = new StringBuilder();
            bool quoted = false;
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (c == '"')
                {
                    current.Append(c);
                    if (quoted && i + 1 < text.Length && text[i + 1] == '"')
                    {
                        current.Append('"'); i++;
                    }
                    else quoted = !quoted;
                }
                else if (c == separator && !quoted)
                {
                    parts.Add(current.ToString()); current.Length = 0;
                }
                else current.Append(c);
            }
            parts.Add(current.ToString());
            return parts;
        }

        private static CellValue ParseArrayValue(string raw)
        {
            string s = (raw ?? "").Trim();
            if (s.Length >= 2 && s[0] == '"' && s[s.Length - 1] == '"')
                return CellValue.FromText(s.Substring(1, s.Length - 2).Replace("\"\"", "\""));
            double d;
            if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out d))
                return CellValue.FromNumber(d);
            if (string.Equals(s, "TRUE", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(s, "VERDADERO", StringComparison.OrdinalIgnoreCase))
                return CellValue.FromBool(true);
            if (string.Equals(s, "FALSE", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(s, "FALSO", StringComparison.OrdinalIgnoreCase))
                return CellValue.FromBool(false);
            if (s.StartsWith("#", StringComparison.Ordinal)) return CellValue.FromError(s);
            return CellValue.Empty;
        }
    }

    internal sealed class BoolNode : FormulaNode
    {
        private readonly bool _v;
        public BoolNode(bool v) { _v = v; }
        internal override CellValue Eval(EvalContext ctx) { return CellValue.FromBool(_v); }
    }

    internal sealed class ErrorNode : FormulaNode
    {
        private readonly string _v;
        public ErrorNode(string v) { _v = v; }
        internal override CellValue Eval(EvalContext ctx) { return CellValue.FromError(_v); }
    }

    internal sealed class MissingNode : FormulaNode
    {
        internal override CellValue Eval(EvalContext ctx) { return CellValue.Empty; }
    }

    internal sealed class NameNode : FormulaNode
    {
        private readonly string _name;
        public string Name { get { return _name; } }
        public NameNode(string name) { _name = name; }
        internal override CellValue Eval(EvalContext ctx)
        {
            CellValue variable;
            if (ctx.TryGetVariable(_name, out variable)) return variable;
            return ctx.EvaluateName(_name);
        }
    }

    internal sealed class RefNode : FormulaNode
    {
        public readonly string SheetName;
        public readonly int Row;
        public readonly int Col;
        public readonly bool AbsRow;
        public readonly bool AbsCol;

        public RefNode(string sheet, int row, int col, bool absRow, bool absCol)
        {
            SheetName = sheet; Row = row; Col = col; AbsRow = absRow; AbsCol = absCol;
        }

        internal override CellValue Eval(EvalContext ctx)
        {
            Sheet sh = ctx.ResolveSheet(SheetName);
            if (sh == null) return CellValue.FromError(Errors.Ref);
            return ctx.GetCell(sh, Row, Col);
        }
    }

    internal sealed class RangeNode : FormulaNode
    {
        public readonly string SheetName;
        public readonly string EndSheetName;
        public readonly int R1;
        public readonly int C1;
        public readonly int R2;
        public readonly int C2;

        public RangeNode(string sheet, int r1, int c1, int r2, int c2, string endSheet)
        {
            SheetName = sheet; EndSheetName = endSheet;
            R1 = Math.Min(r1, r2); C1 = Math.Min(c1, c2);
            R2 = Math.Max(r1, r2); C2 = Math.Max(c1, c2);
        }

        internal override CellValue Eval(EvalContext ctx)
        {
            Sheet sh = ctx.ResolveSheet(SheetName);
            if (sh == null) return CellValue.FromError(Errors.Ref);
            if (!string.IsNullOrEmpty(EndSheetName))
            {
                Sheet end = ctx.ResolveSheet(EndSheetName);
                if (end == null || end != sh) return CellValue.FromError(Errors.Ref);
            }
            return CellValue.FromRange(new RangeRef(sh, R1, C1, R2, C2));
        }
    }

    internal sealed class UnaryNode : FormulaNode
    {
        private readonly string _op;
        private readonly FormulaNode _n;
        public UnaryNode(string op, FormulaNode n) { _op = op; _n = n; }

        internal override CellValue Eval(EvalContext ctx)
        {
            CellValue v = _n.Eval(ctx);
            CellValue err;
            double d;
            if (!EvalUtil.TryNumber(v, ctx, out d, out err)) return err;
            return CellValue.FromNumber(_op == "-" ? -d : d);
        }
    }

    internal sealed class PercentNode : FormulaNode
    {
        private readonly FormulaNode _n;
        public PercentNode(FormulaNode n) { _n = n; }

        internal override CellValue Eval(EvalContext ctx)
        {
            CellValue v = _n.Eval(ctx);
            CellValue err;
            double d;
            if (!EvalUtil.TryNumber(v, ctx, out d, out err)) return err;
            return CellValue.FromNumber(d / 100.0);
        }
    }

    internal sealed class BinaryNode : FormulaNode
    {
        private readonly string _op;
        private readonly FormulaNode _a;
        private readonly FormulaNode _b;

        public BinaryNode(string op, FormulaNode a, FormulaNode b)
        {
            _op = op; _a = a; _b = b;
        }

        internal override CellValue Eval(EvalContext ctx)
        {
            CellValue rawA = _a.Eval(ctx);
            CellValue rawB = _b.Eval(ctx);
            if (IsMatrix(rawA) || IsMatrix(rawB)) return EvalMatrix(_op, rawA, rawB, ctx);
            CellValue av = EvalUtil.Deref(rawA, ctx);
            CellValue bv = EvalUtil.Deref(rawB, ctx);
            return Apply(_op, av, bv, ctx);
        }

        private static bool IsMatrix(CellValue v)
        {
            return v.Kind == ValueKind.Range || v.Kind == ValueKind.Array;
        }

        private sealed class Shape
        {
            public int Rows, Cols;
            public List<CellValue> Cells;
        }

        private static Shape Flatten(CellValue value, EvalContext ctx)
        {
            Shape s = new Shape();
            if (value.Kind == ValueKind.Range)
            {
                RangeRef r = value.Range;
                s.Rows = r.Rows; s.Cols = r.Cols; s.Cells = new List<CellValue>();
                for (int row = r.Row1; row <= r.Row2; row++)
                    for (int col = r.Col1; col <= r.Col2; col++)
                        s.Cells.Add(ctx.GetCell(r.Sheet, row, col));
            }
            else if (value.Kind == ValueKind.Array)
            {
                s.Rows = value.ArrayRows > 0 ? value.ArrayRows : 1;
                s.Cols = value.ArrayCols > 0 ? value.ArrayCols : Math.Max(1, value.Values == null ? 0 : value.Values.Count);
                s.Cells = value.Values == null ? new List<CellValue>() : new List<CellValue>(value.Values);
            }
            else
            {
                s.Rows = 1; s.Cols = 1;
                s.Cells = new List<CellValue>(); s.Cells.Add(value);
            }
            return s;
        }

        private static CellValue EvalMatrix(string op, CellValue a, CellValue b, EvalContext ctx)
        {
            Shape sa = Flatten(a, ctx);
            Shape sb = Flatten(b, ctx);
            int rows = Math.Max(sa.Rows, sb.Rows);
            int cols = Math.Max(sa.Cols, sb.Cols);
            if ((sa.Rows > 1 && sb.Rows > 1 && sa.Rows != sb.Rows) ||
                (sa.Cols > 1 && sb.Cols > 1 && sa.Cols != sb.Cols))
                return CellValue.FromError(Errors.Value);
            if ((long)rows * cols > 100000) return CellValue.FromError(Errors.Value);
            List<CellValue> result = new List<CellValue>();
            for (int r = 0; r < rows; r++)
                for (int c = 0; c < cols; c++)
                {
                    CellValue av = sa.Cells[(sa.Rows == 1 ? 0 : Math.Min(r, sa.Rows - 1)) * sa.Cols +
                        (sa.Cols == 1 ? 0 : Math.Min(c, sa.Cols - 1))];
                    CellValue bv = sb.Cells[(sb.Rows == 1 ? 0 : Math.Min(r, sb.Rows - 1)) * sb.Cols +
                        (sb.Cols == 1 ? 0 : Math.Min(c, sb.Cols - 1))];
                    av = EvalUtil.Deref(av, ctx);
                    bv = EvalUtil.Deref(bv, ctx);
                    result.Add(Apply(op, av, bv, ctx));
                }
            return CellValue.FromArray(result, rows, cols);
        }

        internal static CellValue Apply(string op, CellValue av, CellValue bv, EvalContext ctx)
        {
            if (av.IsError) return av;
            if (bv.IsError) return bv;
            if (op == "&")
            {
                CellValue err;
                string sa = EvalUtil.ToText(av, ctx, out err);
                if (err.IsError) return err;
                string sb = EvalUtil.ToText(bv, ctx, out err);
                if (err.IsError) return err;
                return CellValue.FromText(sa + sb);
            }
            if (op == "=" || op == "<>" || op == "<" || op == ">" || op == "<=" || op == ">=")
            {
                int cmp = EvalUtil.Compare(av, bv);
                bool result;
                switch (op)
                {
                    case "=": result = cmp == 0; break;
                    case "<>": result = cmp != 0; break;
                    case "<": result = cmp < 0; break;
                    case ">": result = cmp > 0; break;
                    case "<=": result = cmp <= 0; break;
                    default: result = cmp >= 0; break;
                }
                return CellValue.FromBool(result);
            }
            CellValue ea, eb;
            double x, y;
            if (!EvalUtil.TryNumber(av, ctx, out x, out ea)) return ea;
            if (!EvalUtil.TryNumber(bv, ctx, out y, out eb)) return eb;
            double r;
            switch (op)
            {
                case "+": r = x + y; break;
                case "-": r = x - y; break;
                case "*": r = x * y; break;
                case "/":
                    if (y == 0) return CellValue.FromError(Errors.Div0);
                    r = x / y;
                    break;
                case "^":
                    r = Math.Pow(x, y);
                    if (double.IsNaN(r) || double.IsInfinity(r)) return CellValue.FromError(Errors.Num);
                    break;
                default: return CellValue.FromError(Errors.Value);
            }
            if (double.IsNaN(r) || double.IsInfinity(r)) return CellValue.FromError(Errors.Num);
            return CellValue.FromNumber(r);
        }
    }

    internal sealed class FunctionNode : FormulaNode
    {
        public readonly string Name;
        public readonly List<FormulaNode> Args;

        public FunctionNode(string name, List<FormulaNode> args)
        {
            Name = name; Args = args;
        }

        internal override CellValue Eval(EvalContext ctx)
        {
            return Builtins.Invoke(this, ctx);
        }
    }

    internal sealed class EvalContext
    {
        public readonly Workbook Workbook;
        public Sheet CurrentSheet;
        public int CurrentRow;
        public int CurrentCol;
        private readonly Dictionary<Cell, CellValue> _cache = new Dictionary<Cell, CellValue>();
        private readonly HashSet<Cell> _visiting = new HashSet<Cell>();
        private readonly Dictionary<string, CellValue> _nameCache = new Dictionary<string, CellValue>(StringComparer.OrdinalIgnoreCase);
        public readonly Dictionary<string, CellValue> Variables =
            new Dictionary<string, CellValue>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _nameVisiting = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private int _depth;
        private const int MaxEvalDepth = 512;

        public EvalContext(Workbook wb)
        {
            Workbook = wb;
        }

        public Sheet ResolveSheet(string name)
        {
            if (string.IsNullOrEmpty(name)) return CurrentSheet;
            return Workbook.FindSheet(name);
        }

        public bool TryGetVariable(string name, out CellValue value)
        {
            return Variables.TryGetValue(name ?? "", out value);
        }

        public CellValue EvaluateName(string name)
        {
            int currentIndex = -1;
            if (CurrentSheet != null) currentIndex = Workbook.Sheets.IndexOf(CurrentSheet);
            DefinedName def = Workbook.FindDefinedName(name, currentIndex);
            if (def == null) return CellValue.FromError(Errors.Name);
            string key = def.LocalSheetId + "|" + def.Name;
            CellValue cached;
            if (_nameCache.TryGetValue(key, out cached)) return cached;
            if (_nameVisiting.Contains(key)) return CellValue.FromError(Errors.Circular);
            if (_depth >= MaxEvalDepth) return CellValue.FromError(Errors.Value);
            _nameVisiting.Add(key);
            _depth++;
            Sheet oldSheet = CurrentSheet;
            int oldRow = CurrentRow;
            int oldCol = CurrentCol;
            CellValue value;
            try
            {
                if (def.LocalSheetId >= 0 && def.LocalSheetId < Workbook.Sheets.Count)
                    CurrentSheet = Workbook.Sheets[def.LocalSheetId];
                FormulaNode ast = FormulaEngine.ParseInternal(def.Formula);
                value = ast.Eval(this);
            }
            catch (FormulaParseException)
            {
                value = CellValue.FromError(Errors.Syntax);
            }
            catch
            {
                value = CellValue.FromError(Errors.Error);
            }
            finally
            {
                CurrentSheet = oldSheet;
                CurrentRow = oldRow;
                CurrentCol = oldCol;
                _nameVisiting.Remove(key);
                _depth--;
            }
            _nameCache[key] = value;
            return value;
        }

        public CellValue GetCell(Sheet sheet, int row, int col)
        {
            Cell cell = sheet.GetCell(row, col);
            if (cell == null) return CellValue.Empty;
            if (!cell.IsFormula) return cell.Value;
            CellValue cached;
            if (_cache.TryGetValue(cell, out cached)) return cached;
            if (_visiting.Contains(cell))
                return CellValue.FromError(Errors.Circular);
            if (_depth >= MaxEvalDepth)
                return CellValue.FromError(Errors.Value);

            _visiting.Add(cell);
            _depth++;
            Sheet oldSheet = CurrentSheet;
            int oldRow = CurrentRow;
            int oldCol = CurrentCol;
            CellValue value;
            try
            {
                CurrentSheet = sheet;
                CurrentRow = row;
                CurrentCol = col;

                if (!cell.AstReady)
                {
                    cell.AstReady = true;
                    cell.Ast = null;
                    try
                    {
                        cell.Ast = FormulaEngine.ParseInternal(cell.Formula);
                    }
                    catch (FormulaParseException)
                    {
                        cell.Ast = null;
                    }
                }

                if (cell.Ast == null)
                    value = CellValue.FromError(Errors.Syntax);
                else
                {
                    try
                    {
                        value = cell.Ast.Eval(this);
                    }
                    catch (Exception)
                    {
                        value = CellValue.FromError(Errors.Value);
                    }
                }
            }
            finally
            {
                CurrentSheet = oldSheet;
                CurrentRow = oldRow;
                CurrentCol = oldCol;
                _visiting.Remove(cell);
                _depth--;
            }
            // Si Excel guardó un resultado válido y el motor no entiende una
            // función o un nombre de este libro, conservamos el valor cacheado
            // en vez de mostrar #NAME?/#VALUE! y perder el dato visible.
            if (value.IsError && cell.HasCachedValue && !cell.CachedValue.IsError)
                value = cell.CachedValue;
            cell.Value = value;
            _cache[cell] = value;
            return value;
        }
    }

    internal static class EvalUtil
    {
        public static CellValue Deref(CellValue v, EvalContext ctx)
        {
            if (v.Kind != ValueKind.Range) return v;
            RangeRef r = v.Range;
            if (r.Rows == 1 && r.Cols == 1) return ctx.GetCell(r.Sheet, r.Row1, r.Col1);
            return CellValue.FromError(Errors.Value);
        }

        public static bool TryNumber(CellValue v, EvalContext ctx, out double d, out CellValue error)
        {
            d = 0;
            error = CellValue.Empty;
            v = Deref(v, ctx);
            if (v.IsError) { error = v; return false; }
            switch (v.Kind)
            {
                case ValueKind.Empty: d = 0; return true;
                case ValueKind.Number:
                    if (double.IsNaN(v.Number) || double.IsInfinity(v.Number))
                    {
                        error = CellValue.FromError(Errors.Num);
                        return false;
                    }
                    d = v.Number; return true;
                case ValueKind.Bool: d = v.Bool ? 1 : 0; return true;
                case ValueKind.Text:
                    if (NumberParser.TryParse(v.Text, out d) &&
                        !double.IsNaN(d) && !double.IsInfinity(d)) return true;
                    error = CellValue.FromError(Errors.Value);
                    return false;
                default:
                    error = CellValue.FromError(Errors.Value);
                    return false;
            }
        }

        public static string ToText(CellValue v, EvalContext ctx, out CellValue error)
        {
            error = CellValue.Empty;
            v = Deref(v, ctx);
            if (v.IsError) { error = v; return ""; }
            switch (v.Kind)
            {
                case ValueKind.Empty: return "";
                case ValueKind.Text: return v.Text ?? "";
                case ValueKind.Number: return NumberFormatter.FormatGeneral(v.Number);
                case ValueKind.Bool: return v.Bool ? "VERDADERO" : "FALSO";
                default: error = CellValue.FromError(Errors.Value); return "";
            }
        }

        public static bool TryBool(CellValue v, EvalContext ctx, out bool b, out CellValue error)
        {
            b = false;
            error = CellValue.Empty;
            v = Deref(v, ctx);
            if (v.IsError) { error = v; return false; }
            if (v.Kind == ValueKind.Bool) { b = v.Bool; return true; }
            if (v.Kind == ValueKind.Number) { b = v.Number != 0; return true; }
            if (v.Kind == ValueKind.Empty) { b = false; return true; }
            if (v.Kind == ValueKind.Text)
            {
                string u = (v.Text ?? "").Trim().ToUpperInvariant();
                if (u == "VERDADERO" || u == "TRUE" || u == "1") { b = true; return true; }
                if (u == "FALSO" || u == "FALSE" || u == "0" || u.Length == 0) { b = false; return true; }
            }
            error = CellValue.FromError(Errors.Value);
            return false;
        }

        public static int Compare(CellValue a, CellValue b)
        {
            if (a.IsError || b.IsError) return 0;
            if (a.Kind == ValueKind.Empty && b.Kind == ValueKind.Empty) return 0;
            if (a.Kind == ValueKind.Empty)
            {
                if (b.Kind == ValueKind.Number) return 0.0.CompareTo(b.Number);
                if (b.Kind == ValueKind.Bool) return 0.0.CompareTo(b.Bool ? 1.0 : 0.0);
                if (b.Kind == ValueKind.Text) return string.Compare("", b.Text ?? "", StringComparison.CurrentCultureIgnoreCase);
            }
            if (b.Kind == ValueKind.Empty)
            {
                if (a.Kind == ValueKind.Number) return a.Number.CompareTo(0.0);
                if (a.Kind == ValueKind.Bool) return (a.Bool ? 1.0 : 0.0).CompareTo(0.0);
                if (a.Kind == ValueKind.Text) return string.Compare(a.Text ?? "", "", StringComparison.CurrentCultureIgnoreCase);
            }
            if (a.Kind == ValueKind.Number && b.Kind == ValueKind.Number) return a.Number.CompareTo(b.Number);
            if (a.Kind == ValueKind.Bool && b.Kind == ValueKind.Bool) return a.Bool.CompareTo(b.Bool);
            if (a.Kind == ValueKind.Text && b.Kind == ValueKind.Text)
                return string.Compare(a.Text ?? "", b.Text ?? "", StringComparison.CurrentCultureIgnoreCase);
            // Excel no convierte texto en número al comparar: el texto siempre
            // es mayor que cualquier número (aunque su contenido sea "10").
            if (a.Kind == ValueKind.Number && b.Kind == ValueKind.Text) return -1;
            if (a.Kind == ValueKind.Text && b.Kind == ValueKind.Number) return 1;
            int rankA = TypeRank(a.Kind);
            int rankB = TypeRank(b.Kind);
            return rankA.CompareTo(rankB);
        }

        private static int TypeRank(ValueKind k)
        {
            switch (k)
            {
                case ValueKind.Number: return 0;
                case ValueKind.Text: return 1;
                case ValueKind.Bool: return 2;
                default: return 3;
            }
        }
    }

    public static class FormulaEngine
    {
        internal static FormulaNode ParseInternal(string formula)
        {
            string f = formula ?? "";
            if (f.Length > 8192)
                throw new FormulaParseException("La fórmula supera el límite de 8192 caracteres.", 0);
            if (f.StartsWith("=", StringComparison.Ordinal)) f = f.Substring(1);
            if (f.Trim().Length == 0) throw new FormulaParseException("La fórmula está vacía.", 0);
            FormulaLexer lex = new FormulaLexer(f);
            FormulaParser parser = new FormulaParser(lex.Tokenize());
            return parser.ParseFormula();
        }

        public static bool Validate(string formula, out string error)
        {
            try
            {
                ParseInternal(formula);
                error = null;
                return true;
            }
            catch (FormulaParseException ex)
            {
                error = ex.Message;
                return false;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        public static void Recalculate(Workbook wb)
        {
            if (wb == null) return;
            foreach (Sheet sheet in wb.Sheets)
            {
                // Las celdas de derrame son valores generados: se retiran antes
                // de volver a evaluar para que no queden restos de una fórmula
                // que ya no devuelve una matriz.
                List<long> spillKeys = new List<long>();
                foreach (KeyValuePair<long, Cell> kv in sheet.Cells)
                    if (kv.Value.SpillOwner != 0) spillKeys.Add(kv.Key);
                foreach (long key in spillKeys) sheet.Cells.Remove(key);

                for (int pass = 0; pass < 2; pass++)
                {
                    EvalContext ctx = new EvalContext(wb);
                    List<long> keys = new List<long>(sheet.Cells.Keys);
                    foreach (long key in keys)
                    {
                        Cell c;
                        if (!sheet.Cells.TryGetValue(key, out c)) continue;
                        if (!c.IsFormula) continue;
                        int row = (int)(key >> 14);
                        int col = (int)(key & 0x3FFF);
                        try
                        {
                            ctx.GetCell(sheet, row, col);
                            CellValue value = c.Value;
                            if (value.Kind == ValueKind.Array)
                                Spill(sheet, row, col, key + 1, value);
                        }
                        catch
                        {
                            c.Value = CellValue.FromError(Errors.Value);
                        }
                    }
                }
                sheet.RecalculateUsedRange();
                sheet.DataVersion++;
            }
        }

        private static void Spill(Sheet sheet, int row, int col, long ownerKey, CellValue value)
        {
            int rows = value.ArrayRows;
            int cols = value.ArrayCols;
            if (rows <= 1 && cols <= 1) return;
            if (rows < 1 || cols < 1 || value.Values == null) return;
            if ((long)rows * cols > 100000) return;
            if (row + rows > Limits.MaxRows || col + cols > Limits.MaxCols) return;
            for (int r = 0; r < rows; r++)
                for (int c = 0; c < cols; c++)
                {
                    if (r == 0 && c == 0) continue;
                    Cell existing = sheet.GetCell(row + r, col + c);
                    if (existing != null && existing.SpillOwner != ownerKey &&
                        (existing.IsFormula || existing.SpillOwner != 0 ||
                         existing.Value.Kind != ValueKind.Empty || !string.IsNullOrEmpty(existing.Comment)))
                    {
                        Cell owner = sheet.GetCell(row, col);
                        if (owner != null) owner.Value = CellValue.FromError(Errors.Spill);
                        return;
                    }
                }
            for (int r = 0; r < rows; r++)
                for (int c = 0; c < cols; c++)
                {
                    if (r == 0 && c == 0) continue;
                    int index = r * cols + c;
                    CellValue cellValue = index < value.Values.Count ? value.Values[index] : CellValue.Empty;
                    Cell target = sheet.GetOrCreate(row + r, col + c);
                    target.Formula = null;
                    target.Raw = null;
                    target.Ast = null;
                    target.AstReady = false;
                    target.Value = cellValue;
                    target.CachedValue = cellValue;
                    target.HasCachedValue = true;
                    target.SpillOwner = ownerKey;
                    target.Dirty = false;
                }
        }

        public static string TranslateFormula(string formula, int dRow, int dCol)
        {
            if (string.IsNullOrEmpty(formula)) return formula;
            return RewriteRefs(formula, delegate(RefToken r)
            {
                int nr = r.Row + (r.AbsRow ? 0 : dRow);
                int nc = r.Col + (r.AbsCol ? 0 : dCol);
                if (nr < 0 || nr >= Limits.MaxRows || nc < 0 || nc >= Limits.MaxCols)
                    return "#REF!";
                r.Row = nr; r.Col = nc;
                return r.Build();
            });
        }

        public static string ShiftFormula(string formula, string formulaSheet, string targetSheet,
                                          bool insert, bool rows, int at, int count)
        {
            if (string.IsNullOrEmpty(formula) || count <= 0) return formula;
            return RewriteRefs(formula, delegate(RefToken r)
            {
                string effective = string.IsNullOrEmpty(r.SheetName) ? formulaSheet : r.SheetName;
                if (!string.Equals(effective, targetSheet, StringComparison.OrdinalIgnoreCase))
                    return r.Original;

                int index = rows ? r.Row : r.Col;
                if (insert)
                {
                    if (index >= at) index += count;
                }
                else
                {
                    if (index >= at && index < at + count) return "#REF!";
                    if (index >= at + count) index -= count;
                }
                if (rows) r.Row = index; else r.Col = index;
                if (r.Row < 0 || r.Row >= Limits.MaxRows || r.Col < 0 || r.Col >= Limits.MaxCols)
                    return "#REF!";
                return r.Build();
            });
        }

        public static string RenameSheetReferences(string formula, string oldName, string newName)
        {
            if (string.IsNullOrEmpty(formula)) return formula;
            StringBuilder sb = new StringBuilder();
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
                    sb.Append(formula.Substring(i, Math.Min(j, formula.Length) - i));
                    i = Math.Min(j, formula.Length);
                    continue;
                }
                if (ch == '\'' || char.IsLetter(ch) || ch == '_')
                {
                    int j = i;
                    string name = null;
                    if (ch == '\'')
                    {
                        j++;
                        StringBuilder q = new StringBuilder();
                        while (j < formula.Length)
                        {
                            if (formula[j] == '\'')
                            {
                                if (j + 1 < formula.Length && formula[j + 1] == '\'') { q.Append('\''); j += 2; continue; }
                                j++;
                                break;
                            }
                            q.Append(formula[j]); j++;
                        }
                        name = q.ToString();
                    }
                    else
                    {
                        while (j < formula.Length && (char.IsLetterOrDigit(formula[j]) || formula[j] == '_' || formula[j] == '.'))
                            j++;
                        name = formula.Substring(i, j - i);
                    }
                    int k = j;
                    while (k < formula.Length && formula[k] == ' ') k++;
                    if (k < formula.Length && formula[k] == '!' && name != null &&
                        string.Equals(name, oldName, StringComparison.OrdinalIgnoreCase))
                    {
                        bool needsQuotes = newName.IndexOf(' ') >= 0 || newName.IndexOf('\'') >= 0;
                        string display = needsQuotes ? "'" + newName.Replace("'", "''") + "'" : newName;
                        sb.Append(display);
                        i = k;
                        continue;
                    }
                    sb.Append(formula.Substring(i, j - i));
                    i = j;
                    continue;
                }
                sb.Append(ch);
                i++;
            }
            return sb.ToString();
        }

        private sealed class RefToken
        {
            public string Original;
            public string SheetName;
            public int Row;
            public int Col;
            public bool AbsRow;
            public bool AbsCol;

            public string Build()
            {
                return (AbsCol ? "$" : "") + AddressUtil.ColumnName(Col) +
                       (AbsRow ? "$" : "") + (Row + 1).ToString(CultureInfo.InvariantCulture);
            }
        }

        private delegate string RefTransform(RefToken token);

        private static string RewriteRefs(string formula, RefTransform transform)
        {
            StringBuilder sb = new StringBuilder();
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

                if (ch == '$' || char.IsLetter(ch))
                {
                    int start = i;
                    int j = i;
                    if (formula[j] == '$') j++;
                    int letterStart = j;
                    while (j < formula.Length && char.IsLetter(formula[j]) && j - letterStart < 3) j++;
                    int letterCount = j - letterStart;
                    if (letterCount > 0)
                    {
                        if (j < formula.Length && formula[j] == '$') j++;
                        int digitStart = j;
                        while (j < formula.Length && char.IsDigit(formula[j]) && j - digitStart < 7) j++;
                        int digitCount = j - digitStart;
                        if (digitCount > 0)
                        {
                            bool beforeOk = start == 0 || !IsIdentChar(formula[start - 1]);
                            bool afterOk = j >= formula.Length || (!IsIdentChar(formula[j]) && formula[j] != '(' && formula[j] != '!');
                            string token = formula.Substring(start, j - start);
                            int row, col; bool ar, ac;
                            if (beforeOk && afterOk && AddressUtil.TryParseCell(token, out row, out col, out ar, out ac))
                            {
                                RefToken rt = new RefToken();
                                rt.Original = token;
                                rt.Row = row; rt.Col = col;
                                rt.AbsRow = ar; rt.AbsCol = ac;
                                rt.SheetName = FindSheetPrefix(formula, start);
                                string replacement = transform(rt);
                                sb.Append(replacement);
                                i = j;
                                continue;
                            }
                        }
                    }
                }
                sb.Append(ch);
                i++;
            }
            return sb.ToString();
        }

        private static bool IsIdentChar(char c)
        {
            return char.IsLetterOrDigit(c) || c == '_' || c == '.' || c == '$';
        }

        private static string FindSheetPrefix(string formula, int pos)
        {
            int k = pos - 1;
            while (k >= 0 && formula[k] == ' ') k--;
            if (k < 0 || formula[k] != '!') return null;
            k--;
            while (k >= 0 && formula[k] == ' ') k--;
            if (k < 0) return null;
            if (formula[k] == '\'')
            {
                int end = k;
                k--;
                StringBuilder sb = new StringBuilder();
                while (k >= 0)
                {
                    if (formula[k] == '\'')
                    {
                        if (k - 1 >= 0 && formula[k - 1] == '\'') { sb.Insert(0, '\''); k -= 2; continue; }
                        break;
                    }
                    sb.Insert(0, formula[k]);
                    k--;
                }
                return sb.ToString();
            }
            int end2 = k;
            while (k >= 0 && (char.IsLetterOrDigit(formula[k]) || formula[k] == '_' || formula[k] == '.'))
                k--;
            if (k == end2) return null;
            return formula.Substring(k + 1, end2 - k);
        }
    }
}
