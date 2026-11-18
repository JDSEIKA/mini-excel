using System;
using CalcRapido;

public static class ModernFunctionsProbe
{
    private static int _ok, _bad;
    private static Sheet _s;

    private static void Set(string a, string v)
    {
        int r, c; bool ar, ac; AddressUtil.TryParseCell(a, out r, out c, out ar, out ac);
        _s.SetInput(r, c, v);
    }

    private static string Get(string a)
    {
        int r, c; bool ar, ac; AddressUtil.TryParseCell(a, out r, out c, out ar, out ac);
        Cell x = _s.GetCell(r, c);
        return x == null ? "" : x.Value.ToDisplayString();
    }

    private static void Check(string a, string expected)
    {
        string actual = Get(a);
        if (actual == expected) { _ok++; Console.WriteLine("OK  " + a + " = " + actual); }
        else { _bad++; Console.WriteLine("BAD " + a + " esperado=[" + expected + "] real=[" + actual + "]"); }
    }

    public static int Main()
    {
        Workbook wb = new Workbook();
        _s = wb.Sheets[0];
        Set("A1", "1"); Set("A2", "2"); Set("A3", "3"); Set("A4", "4");
        Set("B1", "A"); Set("B2", "B"); Set("B3", "A"); Set("B4", "B");
        Set("C1", "10"); Set("C2", "20"); Set("C3", "30"); Set("C4", "40");

        Set("E1", "=SUMIFS(C1:C4;A1:A4;\">1\";B1:B4;\"A\")");
        Set("E2", "=COUNTIFS(A1:A4;\">1\";B1:B4;\"B\")");
        Set("E3", "=AVERAGEIFS(C1:C4;A1:A4;\">1\";B1:B4;\"A\")");
        Set("E4", "=MAXIFS(C1:C4;A1:A4;\">1\")");
        Set("E5", "=MINIFS(C1:C4;A1:A4;\">1\")");
        Set("E6", "=SUBTOTAL(9;A1:A4)");
        Set("E7", "=SUBTOTAL(1;A1:A4)");
        Set("E8", "=SUBTOTAL(3;A1:A4)");
        Set("E9", "=ROUND(TIME(1;2;3)*86400;0)");
        Set("E10", "=TIMEVALUE(\"13:30\")*24");
        Set("E11", "=TEXTJOIN(\"-\";VERDADERO;A1:A4)");
        Set("E12", "=SWITCH(2;1;\"uno\";2;\"dos\";\"otro\")");
        Set("E13", "=XLOOKUP(3;A1:A4;C1:C4)");
        Set("E14", "=XMATCH(3;A1:A4)");
        Set("E15", "=ISNA(NA())");
        Set("E16", "=ISNA(1/0)");
        Set("E17", "=ISERR(1/0)");
        Set("E18", "=ROUND(STDEV.S(A1:A4);4)");
        Set("E19", "=DAYS(DATE(2026;1;10);DATE(2026;1;1))");
        Set("E20", "=DATEDIF(DATE(2020;1;15);DATE(2026;3;20);\"Y\")");
        Set("E21", "=NETWORKDAYS(DATE(2026;1;1);DATE(2026;1;7))");
        Set("E22", "=WORKDAY(DATE(2026;1;2);1)");
        Set("E23", "=ROUND(YEARFRAC(DATE(2026;1;1);DATE(2026;7;1);1);4)");
        Set("E24", "=RANK(3;A1:A4)");
        Set("E25", "=PERCENTILE.INC(A1:A4;0.5)");
        Set("E26", "=QUARTILE.INC(A1:A4;1)");
        Set("E27", "=CEILING.MATH(-4.3;2)");
        Set("E28", "=FLOOR.MATH(-4.3;2)");
        Set("E29", "=AGGREGATE(9;6;A1:A4)");
        Set("E30", "=AGGREGATE(1;6;A1:A4)");
        Set("E31", "=TEXTBEFORE(\"uno-dos-tres\";\"-\";2)");
        Set("E32", "=TEXTAFTER(\"uno-dos-tres\";\"-\";2)");
        Set("E33", "=TEXTBEFORE(\"abc\";\"-\")");

        wb.Recalculate(); // no-op after SetInput, but keeps the intent explicit

        Check("E1", "30");
        Check("E2", "2");
        Check("E3", "30");
        Check("E4", "40");
        Check("E5", "20");
        Check("E6", "10");
        Check("E7", "2,5");
        Check("E8", "4");
        Check("E9", "3723");
        Check("E10", "13,5");
        Check("E11", "1-2-3-4");
        Check("E12", "dos");
        Check("E13", "30");
        Check("E14", "3");
        Check("E15", "VERDADERO");
        Check("E16", "FALSO");
        Check("E17", "VERDADERO");
        Check("E18", "1,291");
        Check("E19", "9");
        Check("E20", "6");
        Check("E21", "5");
        Check("E22", "46027");
        Check("E23", "0,4959");
        Check("E24", "2");
        Check("E25", "2,5");
        Check("E26", "1,75");
        Check("E27", "-4");
        Check("E28", "-6");
        Check("E29", "10");
        Check("E30", "2,5");
        Check("E31", "uno-dos");
        Check("E32", "tres");
        Check("E33", "#N/A");

        Console.WriteLine("ModernFunctions: " + _ok + " correctas, " + _bad + " fallidas.");
        return _bad == 0 ? 0 : 1;
    }
}
