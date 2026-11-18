using System;
using System.IO;
using System.IO.Compression;
using System.Windows.Forms;
using CalcRapido;

public static class AuditRegressionTest
{
    private static int _ok;
    private static int _bad;

    private static void Check(string name, bool condition, string detail)
    {
        if (condition) { _ok++; Console.WriteLine("OK  " + name); }
        else { _bad++; Console.WriteLine("BAD " + name + " :: " + detail); }
    }

    private static string P(string name)
    {
        return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", name);
    }

    [STAThread]
    public static int Main()
    {
        TestFormulaLimits();
        TestFormulaSemantics();
        TestTextAndNumbers();
        TestNativeStructure();
        TestCutAcrossSheets();
        TestSaveFallback();
        TestCloneIsolation();
        Console.WriteLine("AuditRegression: " + _ok + " correctas, " + _bad + " fallidas.");
        return _bad == 0 ? 0 : 1;
    }

    private static void TestFormulaLimits()
    {
        Workbook wb = new Workbook();
        Sheet s = wb.Sheets[0];
        for (int i = 0; i < 1200; i++)
            s.SetInput(i, 0, i == 1199 ? "1" : "=A" + (i + 2));
        wb.Recalculate();
        Check("cadena profunda no desborda", s.GetCell(0, 0).Value.IsError,
            s.GetCell(0, 0).Value.ToDisplayString());

        string nested = "=" + new string('(', 2000) + "1" + new string(')', 2000);
        string error;
        Check("parser limita anidamiento", !FormulaEngine.Validate(nested, out error), "validó");
    }

    private static void TestFormulaSemantics()
    {
        Workbook wb = new Workbook();
        Sheet s = wb.Sheets[0];
        s.SetInput(0, 0, "=SUM(1,2)");
        s.SetInput(1, 0, "=SUM(1,5;2)");
        s.SetInput(2, 0, "=DATE(2026,11,9)");
        s.SetInput(3, 0, "=-2^2");
        s.SetInput(4, 0, "=HOUR(0.5)");
        s.SetInput(5, 0, "=AND(Z1)");
        s.SetInput(6, 0, "=REPT(\"\";2000000000)");
        s.SetInput(7, 0, "=RANDBETWEEN(1;2147483647)");
        s.SetInput(1048575, 16383, "1");
        s.SetInput(8, 0, "=SUMPRODUCT(A1:XFD1048576)");
        wb.Recalculate();
        Check("coma como separador", s.GetCell(0, 0).Value.Number == 3, s.GetCell(0, 0).Value.ToDisplayString());
        Check("coma decimal con punto y coma", Math.Abs(s.GetCell(1, 0).Value.Number - 3.5) < 1e-9, s.GetCell(1, 0).Value.ToDisplayString());
        Check("DATE con comas", !s.GetCell(2, 0).Value.IsError, s.GetCell(2, 0).Value.ToDisplayString());
        Check("precedencia unaria", s.GetCell(3, 0).Value.Number == 4, s.GetCell(3, 0).Value.ToDisplayString());
        Check("HOUR de solo hora", s.GetCell(4, 0).Value.Number == 12, s.GetCell(4, 0).Value.ToDisplayString());
        Check("AND vacío", !s.GetCell(5, 0).Value.IsError && !s.GetCell(5, 0).Value.Bool, s.GetCell(5, 0).Value.ToDisplayString());
        Check("REPT acotado", s.GetCell(6, 0).Value.IsError, s.GetCell(6, 0).Value.ToDisplayString());
        Check("RANDBETWEEN sin excepción", !s.GetCell(7, 0).Value.IsError, s.GetCell(7, 0).Value.ToDisplayString());
        Check("SUMPRODUCT acotado", s.GetCell(8, 0).Value.IsError, s.GetCell(8, 0).Value.ToDisplayString());
    }

    private static void TestTextAndNumbers()
    {
        string xlsx = P("tests/bin/text-types.xlsx");
        if (File.Exists(xlsx))
        {
            Workbook wb = XlsxIO.Load(xlsx);
            Cell a1 = wb.Sheets[0].GetCell(0, 0);
            Check("xlsx texto 00123", a1.Value.Kind == ValueKind.Text && a1.Value.Text == "00123", a1.Value.ToDisplayString());
            string output = P("tests/bin/audit-text-out.xlsx");
            XlsxIO.SaveNew(wb, output);
            Workbook back = XlsxIO.Load(output);
            Check("roundtrip texto 00123", back.Sheets[0].GetCell(0, 0).Value.Kind == ValueKind.Text &&
                back.Sheets[0].GetCell(0, 0).Value.Text == "00123", back.Sheets[0].GetCell(0, 0).Value.ToDisplayString());
        }
        else Check("fixture text-types.xlsx", false, "no existe");

        string xls = P("tests/bin/numbers.xls");
        if (File.Exists(xls))
        {
            Workbook wb = XlsIO.Load(xls);
            Cell a1 = wb.Sheets[0].GetCell(0, 0);
            Cell b1 = wb.Sheets[0].GetCell(0, 1);
            Cell c1 = wb.Sheets[0].GetCell(0, 2);
            Check("xls 12.345", Math.Abs(a1.Value.Number - 12.345) < 1e-12, a1.Value.ToDisplayString());
            Check("xls unario negativo", Math.Abs(b1.Value.Number + 12.345) < 1e-12, b1.Value.ToDisplayString());
            Check("xls fila 20000", c1.Value.Number == 7, c1.Value.ToDisplayString());
        }
        else Check("fixture numbers.xls", false, "no existe");
    }

    private static void TestNativeStructure()
    {
        string path = P("tests/bin/audit-native.crb");
        Workbook wb = new Workbook();
        Sheet s = wb.Sheets[0];
        s.SetTextLiteral(0, 0, "=1+1");
        s.Merges.Add(new MergeRange(0, 0, 0, 1));
        s.RebuildMergeLookup();
        s.SetRowHeightPointsImported(0, 40f);
        s.FreezeRows = 1;
        s.GetOrCreate(1, 0).Hyperlink = "https://example.com";
        s.GetOrCreate(1, 0).HyperlinkTooltip = "ejemplo";
        wb.DefinedNames.Add(new DefinedName { Name = "X", Formula = "Hoja1!$A$1" });
        wb.Date1904 = true;
        FileIO.SaveNative(wb, path);
        Workbook back = FileIO.LoadNative(path);
        Sheet b = back.Sheets[0];
        Check("crb combinaciones", b.Merges.Count == 1, b.Merges.Count.ToString());
        Check("crb alto", Math.Abs(b.GetRowHeightPoints(0) - 40f) < 0.01, b.GetRowHeightPoints(0).ToString());
        Check("crb freeze", b.FreezeRows == 1, b.FreezeRows.ToString());
        Check("crb texto =", b.GetCell(0, 0).Value.Kind == ValueKind.Text, b.GetCell(0, 0).Value.ToDisplayString());
        Check("crb hipervínculo", b.GetCell(1, 0).Hyperlink == "https://example.com", b.GetCell(1, 0).Hyperlink ?? "");
        Check("crb date1904", back.Date1904, back.Date1904.ToString());
        Check("crb nombres", back.DefinedNames.Count == 1, back.DefinedNames.Count.ToString());
    }

    private static void TestCutAcrossSheets()
    {
        Workbook wb = new Workbook();
        Sheet one = wb.Sheets[0]; one.Name = "Uno"; one.SetInput(0, 0, "A");
        Sheet two = wb.AddSheet("Dos"); two.SetInput(0, 0, "B");
        using (SheetGrid grid = new SheetGrid())
        {
            grid.Attach(wb, 0);
            grid.SelectCell(0, 0);
            grid.CopySelection(true);
            grid.SetSheet(two);
            grid.SelectCell(0, 1);
            grid.PasteClipboard();
            Check("corte entre hojas", one.GetCell(0, 0) == null && two.GetCell(0, 1) != null &&
                two.GetCell(0, 1).Value.Text == "A" && two.GetCell(0, 0).Value.Text == "B", "estado incorrecto");
        }
    }

    private static void TestCloneIsolation()
    {
        Workbook wb = new Workbook();
        Sheet s = wb.Sheets[0];
        s.PivotRanges.Add(new PivotRange { Row1 = 0, Col1 = 0, Row2 = 2, Col2 = 2,
            Cache = new PivotCache() });
        s.Slicers.Add(new SheetSlicer { Name = "S", FieldName = "Categoría" });
        s.Shapes.Add(new SheetShape { Name = "Shape1" });
        Workbook copy = wb.Clone();
        Check("clone pivote aislado", !object.ReferenceEquals(s.PivotRanges[0], copy.Sheets[0].PivotRanges[0]), "compartido");
        Check("clone segmentador aislado", !object.ReferenceEquals(s.Slicers[0], copy.Sheets[0].Slicers[0]), "compartido");
        Check("clone forma aislada", !object.ReferenceEquals(s.Shapes[0], copy.Sheets[0].Shapes[0]), "compartido");
    }

    private static void TestSaveFallback()
    {
        string sample = P("dist/Ejemplo.xlsx");
        string copy = P("tests/bin/audit-preserve.xlsx");
        if (!File.Exists(sample)) { Check("fixture Ejemplo.xlsx", false, "no existe"); return; }
        File.Copy(sample, copy, true);
        Workbook wb = XlsxIO.Load(copy);
        File.SetLastWriteTimeUtc(copy, DateTime.UtcNow.AddSeconds(5));
        bool threw = false;
        try { XlsxIO.Save(wb, copy); }
        catch (FileFormatException) { threw = true; }
        int entries;
        using (ZipArchive z = ZipFile.OpenRead(copy)) entries = z.Entries.Count;
        Check("guardado no destructivo", threw && entries > 8, "entries=" + entries + " threw=" + threw);
    }
}
