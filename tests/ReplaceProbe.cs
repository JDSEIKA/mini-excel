using System;
using System.Windows.Forms;
using CalcRapido;

public static class ReplaceProbe
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
        s.SetTextLiteral(0, 0, "Hola");
        s.SetTextLiteral(1, 0, "hola mundo");
        s.SetTextLiteral(2, 0, "otro");
        s.SetNumber(3, 0, 10);
        s.SetInput(4, 0, "=LEN(A1)");
        SheetGrid grid = new SheetGrid();
        grid.SetSheet(s);
        int n = grid.ReplaceAll("hola", "adiós", false, false, true);
        Check("reemplazo sin distinguir mayúsculas", n == 2, "n=" + n);
        Check("texto reemplazado", s.GetCell(0, 0).Value.Text == "adiós",
            s.GetCell(0, 0).Value.Text ?? "");
        Check("segunda coincidencia", s.GetCell(1, 0).Value.Text == "adiós mundo",
            s.GetCell(1, 0).Value.Text ?? "");
        Check("no toca números", s.GetCell(3, 0).Value.Kind == ValueKind.Number, "número");
        Check("no toca fórmula si no coincide", s.GetCell(4, 0).IsFormula, "fórmula");
        grid.Dispose();
        Console.WriteLine("Replace: " + _ok + " correctas, " + _bad + " fallidas.");
        return _bad == 0 ? 0 : 1;
    }
}
