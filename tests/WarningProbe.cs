using System;
using System.IO;
using CalcRapido;

public static class WarningProbe
{
    private static int _ok, _bad;

    private static void Check(string name, bool condition, string detail)
    {
        if (condition) { _ok++; Console.WriteLine("OK  " + name); }
        else { _bad++; Console.WriteLine("BAD " + name + " [" + detail + "]"); }
    }

    public static int Main(string[] args)
    {
        Workbook wb = new Workbook();
        wb.Sheets[0].SetInput(0, 0, "=SUM(1;2)");
        CompatibilityScanner.ScanFormulas(wb);
        Check("funciones soportadas sin aviso", wb.UnsupportedFunctions.Count == 0,
            string.Join(",", wb.UnsupportedFunctions));

        wb = new Workbook();
        wb.Sheets[0].SetInput(0, 0, "=LAMBDA(x;x+1)(2)");
        CompatibilityScanner.ScanFormulas(wb);
        Check("detecta LAMBDA", wb.UnsupportedFunctions.Contains("LAMBDA"),
            string.Join(",", wb.UnsupportedFunctions));

        string advanced = args.Length > 0 ? args[0] : "tests/bin/advanced.xlsx";
        if (File.Exists(advanced))
        {
            Workbook loaded = XlsxIO.Load(advanced);
            bool chartOrImage = false;
            foreach (string warning in loaded.LoadWarnings)
                if (warning.IndexOf("gráfico", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    warning.IndexOf("imagen", StringComparison.OrdinalIgnoreCase) >= 0) chartOrImage = true;
            Check("detecta gráficos/imágenes", chartOrImage, string.Join(" | ", loaded.LoadWarnings));
        }

        string xls = args.Length > 1 ? args[1] : "tests/corpus/StringContinueRecords.xls";
        if (File.Exists(xls))
        {
            Workbook legacy = XlsIO.Load(xls);
            bool xlsWarning = false;
            foreach (string warning in legacy.LoadWarnings)
                if (warning.IndexOf("97-2003", StringComparison.OrdinalIgnoreCase) >= 0) xlsWarning = true;
            Check("aviso de .xls compatible", xlsWarning, string.Join(" | ", legacy.LoadWarnings));
        }

        Console.WriteLine("Warnings: " + _ok + " correctas, " + _bad + " fallidas.");
        return _bad == 0 ? 0 : 1;
    }
}
