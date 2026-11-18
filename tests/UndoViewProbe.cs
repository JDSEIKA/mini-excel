using System;
using System.Reflection;
using System.Windows.Forms;
using CalcRapido;

public static class UndoViewProbe
{
    [STAThread]
    public static int Main()
    {
        Application.EnableVisualStyles();
        using (MainForm form = new MainForm(null))
        {
            form.Show();
            Application.DoEvents();
            FieldInfo gf = typeof(MainForm).GetField("_grid", BindingFlags.Instance | BindingFlags.NonPublic);
            FieldInfo wf = typeof(MainForm).GetField("_workbook", BindingFlags.Instance | BindingFlags.NonPublic);
            MethodInfo push = typeof(MainForm).GetMethod("PushUndo", BindingFlags.Instance | BindingFlags.NonPublic);
            MethodInfo undo = typeof(MainForm).GetMethod("Undo", BindingFlags.Instance | BindingFlags.NonPublic);
            SheetGrid grid = (SheetGrid)gf.GetValue(form);
            Workbook wb = (Workbook)wf.GetValue(form);
            wb.Sheets[0].SetInput(500, 3, "lejos");
            grid.SelectCell(500, 3);
            Application.DoEvents();
            int rowBefore = grid.ActiveRow, colBefore = grid.ActiveCol;
            int syBefore = grid.ScrollY, sxBefore = grid.ScrollX;
            push.Invoke(form, null);
            grid.Sheet.SetInput(500, 3, "cambio");
            undo.Invoke(form, null);
            Application.DoEvents();
            Console.WriteLine("before=" + rowBefore + "," + colBefore + " scroll=" + sxBefore + "," + syBefore);
            Console.WriteLine("after=" + grid.ActiveRow + "," + grid.ActiveCol + " scroll=" + grid.ScrollX + "," + grid.ScrollY);
            bool ok = grid.ActiveRow == rowBefore && grid.ActiveCol == colBefore &&
                      grid.ScrollX == sxBefore && grid.ScrollY == syBefore;
            FieldInfo closing = typeof(MainForm).GetField("_closing", BindingFlags.Instance | BindingFlags.NonPublic);
            closing.SetValue(form, true);
            form.Close();
            return ok ? 0 : 1;
        }
    }
}
