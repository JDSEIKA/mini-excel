using System;
using CalcRapido;

public static class XlsNamesSharedTest
{
    private static int _ok, _bad;

    private static void Check(string what, string actual, string expected)
    {
        if (actual == expected) { _ok++; Console.WriteLine("OK  " + what + " = " + actual); }
        else { _bad++; Console.WriteLine("BAD " + what + " esperado=[" + expected + "] real=[" + actual + "]"); }
    }

    private static string Cell(Sheet s, string address)
    {
        int r, c; bool ar, ac; AddressUtil.TryParseCell(address, out r, out c, out ar, out ac);
        Cell x = s.GetCell(r, c);
        return x == null ? "" : x.Raw;
    }

    public static int Main(string[] args)
    {
        string shared = args.Length > 0 ? args[0] : "tests/corpus/shared_formulas.xls";
        string names = args.Length > 1 ? args[1] : "tests/corpus/StringContinueRecords.xls";
        Workbook wb = XlsIO.Load(shared);
        Sheet s = wb.Sheets[0];
        Check("shared A2", Cell(s, "A2"), "=B2");
        Check("shared A3", Cell(s, "A3"), "=B3");
        wb = XlsIO.Load(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(shared), "SharedFormulaTest.xls"));
        s = wb.Sheets[0];
        Check("shared normalizada B2", Cell(s, "B2"), "=(A$1*2)");

        wb = XlsIO.Load(names);
        Sheet param = wb.FindSheet("Param");
        Check("name formula P3A", Cell(param, "A3"), "=VLOOKUP(MyTestPack;TestPack_Definition;1;0)");
        Check("name MyEnv", wb.FindDefinedName("MyEnv", -1).Formula, "'Param'!$C$3");
        Check("name MyTestPack", wb.FindDefinedName("MyTestPack", -1).Formula, "'file load sequence'!$J$1");
        Cell p5b = param.GetCell(4, 1);
        Check("name value P5B", p5b == null ? "" : p5b.Value.ToDisplayString(),
            "/apps/tibco/acetppf1/exploit/scripts/");

        Console.WriteLine("XlsNamesShared: " + _ok + " correctas, " + _bad + " fallidas.");
        return _bad == 0 ? 0 : 1;
    }
}
