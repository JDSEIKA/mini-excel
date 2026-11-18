using System;
using CalcRapido;

public static class BasicCompatProbe
{
    private static int _ok;
    private static int _bad;

    private static void Set(Sheet s, string a, string v)
    {
        int r, c; bool ar, ac; AddressUtil.TryParseCell(a, out r, out c, out ar, out ac);
        s.SetInput(r, c, v);
    }

    private static string Get(Sheet s, string a)
    {
        int r, c; bool ar, ac; AddressUtil.TryParseCell(a, out r, out c, out ar, out ac);
        Cell x = s.GetCell(r, c);
        return x == null ? "" : x.Value.ToDisplayString();
    }

    private static void Check(string name, string expected)
    {
        string actual = Get(_s, name);
        if (actual == expected) { _ok++; Console.WriteLine("OK  " + name + " = " + actual); }
        else { _bad++; Console.WriteLine("BAD " + name + " esperado=[" + expected + "] real=[" + actual + "]"); }
    }

    private static Sheet _s;

    public static int Main()
    {
        Workbook wb = new Workbook();
        _s = wb.Sheets[0];
        Set(_s, "A1", "1");
        Set(_s, "A2", "texto");
        Set(_s, "A4", "=1/0");
        Set(_s, "A5", "=\"\"");
        Set(_s, "A6", "2");
        Set(_s, "A7", "3");

        Set(_s, "C1", "=SUM(1,2,3)");
        Set(_s, "C2", "=AVERAGE(1,2,3)");
        Set(_s, "C3", "=MIN(3,1,2)");
        Set(_s, "C4", "=MAX(3,1,2)");
        Set(_s, "C5", "=COUNT(A1:A7)");
        Set(_s, "C6", "=COUNTA(A1:A7)");
        Set(_s, "C7", "=COUNTBLANK(A1:A7)");
        Set(_s, "C8", "=SUM(A1:A7)");

        Set(_s, "D1", "=ROUND(2.5;0)");
        Set(_s, "D2", "=ROUND(-2.5;0)");
        Set(_s, "D3", "=ROUNDUP(2.1;0)");
        Set(_s, "D4", "=ROUNDDOWN(2.9;0)");
        Set(_s, "D5", "=INT(-2.1)");
        Set(_s, "D6", "=TRUNC(-2.9)");
        Set(_s, "D7", "=MOD(-3;2)");
        Set(_s, "D8", "=ABS(-5)");
        Set(_s, "D9", "=SQRT(9)");
        Set(_s, "D10", "=POWER(2;3)");
        Set(_s, "D11", "=FIXED(1234.567;1;VERDADERO)");
        Set(_s, "D12", "=FIXED(1234.567;1;FALSO)");

        Set(_s, "E1", "=IF(1=1;10;20)");
        Set(_s, "E2", "=IFERROR(1/0;99)");
        Set(_s, "E3", "=AND(TRUE;1)");
        Set(_s, "E4", "=OR(FALSE;0)");
        Set(_s, "E5", "=NOT(FALSE)");

        Set(_s, "F1", "=LEN(\"Hola\")");
        Set(_s, "F2", "=LEFT(\"Hola\";2)");
        Set(_s, "F3", "=RIGHT(\"Hola\";2)");
        Set(_s, "F4", "=MID(\"Hola\";2;2)");
        Set(_s, "F5", "=UPPER(\"hola\")");
        Set(_s, "F6", "=TRIM(\" a  b \")");
        Set(_s, "F7", "=PROPER(\"hola mundo\")");
        Set(_s, "F8", "=FIND(\"l\";\"Hola\")");
        Set(_s, "F9", "=SEARCH(\"L\";\"Hola\")");

        Set(_s, "G1", "=DATE(2026;1;2)");
        Set(_s, "G2", "=YEAR(G1)");
        Set(_s, "G3", "=MONTH(G1)");
        Set(_s, "G4", "=DAY(G1)");
        Set(_s, "G5", "=HOUR(0.5)");
        Set(_s, "G6", "=WEEKDAY(G1)");

        Set(_s, "H1", "1");
        Set(_s, "H2", "2");
        Set(_s, "H3", "3");
        Set(_s, "H4", "4");
        Set(_s, "I1", "A");
        Set(_s, "I2", "B");
        Set(_s, "I3", "C");
        Set(_s, "I4", "D");
        Set(_s, "J1", "=VLOOKUP(2;H1:I4;2;FALSE)");
        Set(_s, "J2", "=MATCH(3;H1:H4;0)");
        Set(_s, "J3", "=INDEX(I1:I4;2)");
        Set(_s, "J4", "=COUNTIF(H1:H4;\">2\")");
        Set(_s, "J5", "=SUMIF(H1:H4;\">2\")");
        Set(_s, "J6", "=AVERAGEIF(H1:H4;\">2\")");
        Set(_s, "J7", "=SUMPRODUCT(H1:H4;H1:H4)");

        wb.Recalculate();

        Check("C1", "6");
        Check("C2", "2");
        Check("C3", "1");
        Check("C4", "3");
        Check("C5", "3"); // 1, 2 y 3; ignora texto, vacÃ­o y error
        Check("C6", "6"); // 1, texto, error, "" , 2 y 3
        Check("C7", "2"); // A3 vacÃ­o y A5 fÃ³rmula ""
        Check("C8", "#¡DIV/0!");
        Check("D1", "3");
        Check("D2", "-3");
        Check("D3", "3");
        Check("D4", "2");
        Check("D5", "-3");
        Check("D6", "-2");
        Check("D7", "1");
        Check("D8", "5");
        Check("D9", "3");
        Check("D10", "8");
        Check("D11", "1234,6");
        Check("D12", "1.234,6");
        Check("E1", "10");
        Check("E2", "99");
        Check("E3", "VERDADERO");
        Check("E4", "FALSO");
        Check("E5", "VERDADERO");
        Check("F1", "4");
        Check("F2", "Ho");
        Check("F3", "la");
        Check("F4", "ol");
        Check("F5", "HOLA");
        Check("F6", "a b");
        Check("F7", "Hola Mundo");
        Check("F8", "3");
        Check("F9", "3");
        Check("G1", "46024");
        Check("G2", "2026");
        Check("G3", "1");
        Check("G4", "2");
        Check("G5", "12");
        Check("G6", "6");
        Check("J1", "B");
        Check("J2", "3");
        Check("J3", "B");
        Check("J4", "2");
        Check("J5", "7");
        Check("J6", "3,5");
        Check("J7", "30");

        Console.WriteLine("BasicCompat: " + _ok + " correctas, " + _bad + " fallidas.");
        return _bad == 0 ? 0 : 1;
    }
}

