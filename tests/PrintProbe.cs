using System;
using CalcRapido;

public static class PrintProbe
{
    private static int _ok, _bad;

    private static void Check(string name, bool ok, string detail)
    {
        if (ok) { _ok++; Console.WriteLine("OK  " + name); }
        else { _bad++; Console.WriteLine("BAD " + name + " [" + detail + "]"); }
    }

    public static int Main()
    {
        Workbook wb = new Workbook();
        Sheet s = wb.Sheets[0];
        for (int r = 0; r < 40; r++)
            for (int c = 0; c < 8; c++)
                s.SetNumber(r, c, r * 10 + c);
        s.AddMerge(new MergeRange(1, 1, 2, 3));
        s.SetRowHeightPoints(1, 30);
        wb.AddSheet("Otra");
        wb.Sheets[1].SetInput(0, 0, "Segunda hoja");

        PrintOptions selection = new PrintOptions();
        selection.Scope = PrintScope.Selection;
        selection.Fit = PrintFitMode.FitPage;
        selection.PrintGridlines = true;
        using (WorkbookPrintDocument doc = new WorkbookPrintDocument(wb, selection, 0, 2, 1, 6, 4))
        {
            doc.PrepareLayout();
            Check("selección prepara al menos una página", doc.PreparedPageCount >= 1,
                doc.PreparedPageCount.ToString());
            Check("selección ajustada a una página", doc.PreparedPageCount == 1,
                doc.PreparedPageCount.ToString());
        }

        PrintOptions all = new PrintOptions();
        all.Scope = PrintScope.AllSheets;
        all.Fit = PrintFitMode.FitWidth;
        using (WorkbookPrintDocument doc = new WorkbookPrintDocument(wb, all, 0, 0, 0, 0, 0))
        {
            doc.PrepareLayout();
            Check("todas las hojas produce páginas", doc.PreparedPageCount >= 2,
                doc.PreparedPageCount.ToString());
        }

        PrintOptions chosen = new PrintOptions();
        chosen.Scope = PrintScope.SelectedSheets;
        chosen.SheetIndices.Add(0);
        chosen.SheetIndices.Add(1);
        chosen.Fit = PrintFitMode.FitPage;
        using (WorkbookPrintDocument doc = new WorkbookPrintDocument(wb, chosen, 0, 0, 0, 0, 0))
        {
            doc.PrepareLayout();
            Check("hojas seleccionadas produce páginas", doc.PreparedPageCount >= 2,
                doc.PreparedPageCount.ToString());
        }

        using (PrintOptionsDialog dialog = new PrintOptionsDialog(wb, 0))
        {
            dialog.CreateControl();
            Check("diálogo de impresión se construye", dialog.Controls.Count > 0,
                dialog.Controls.Count.ToString());
        }

        Console.WriteLine("Print: " + _ok + " correctas, " + _bad + " fallidas.");
        return _bad == 0 ? 0 : 1;
    }
}
