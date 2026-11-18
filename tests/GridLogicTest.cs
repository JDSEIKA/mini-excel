using System;
using System.Windows.Forms;
using CalcRapido;

public static class GridLogicTest
{
    private static int _ok;
    private static int _bad;

    private static void Check(string name, string expected, string actual)
    {
        if (string.Equals(expected, actual, StringComparison.Ordinal)) { _ok++; Console.WriteLine("OK  " + name); }
        else { _bad++; Console.WriteLine("BAD " + name + " expected [" + expected + "] actual [" + actual + "]"); }
    }

    private static void Set(Sheet s, string addr, string text)
    {
        int r, c; bool ar, ac;
        AddressUtil.TryParseCell(addr, out r, out c, out ar, out ac);
        s.SetInput(r, c, text);
    }

    private static string Raw(Sheet s, string addr)
    {
        int r, c; bool ar, ac;
        AddressUtil.TryParseCell(addr, out r, out c, out ar, out ac);
        Cell cell = s.GetCell(r, c);
        return cell == null ? "" : (cell.Raw ?? "");
    }

    private static string Val(Sheet s, string addr)
    {
        int r, c; bool ar, ac;
        AddressUtil.TryParseCell(addr, out r, out c, out ar, out ac);
        Cell cell = s.GetCell(r, c);
        return cell == null ? "" : cell.Value.ToDisplayString();
    }

    [STAThread]
    public static int Main(string[] args)
    {
        Application.EnableVisualStyles();
        Workbook wb = new Workbook();
        Sheet s = wb.Sheets[0];
        Set(s, "A1", "10");
        Set(s, "A2", "20");
        Set(s, "A3", "30");
        Set(s, "B1", "=A1+A2");
        Set(s, "C1", "=SUMA(A1:A3)");
        wb.Recalculate();
        Check("formula inicial", "30", Val(s, "B1"));

        using (SheetGrid grid = new SheetGrid())
        {
            grid.Attach(wb, 0);
            grid.SelectRange(0, 0, 0, 1);
            grid.InsertRow();
            wb.Recalculate();
            Check("insertar fila: A1 vacío", "", Raw(s, "A1"));
            Check("insertar fila: A2 10", "10", Val(s, "A2"));
            Check("insertar fila: fórmula movida", "=A2+A3", Raw(s, "B2"));
            Check("insertar fila: fórmula recalculada", "30", Val(s, "B2"));
            Check("insertar fila: rango ajustado", "=SUMA(A2:A4)", Raw(s, "C2"));
            Check("insertar fila: rango valor", "60", Val(s, "C2"));

            grid.SelectRange(0, 0, 0, 0);
            grid.DeleteRow();
            wb.Recalculate();
            Check("eliminar fila: A1 10", "10", Val(s, "A1"));
            Check("eliminar fila: fórmula vuelve", "=A1+A2", Raw(s, "B1"));

            grid.SelectCell(0, 0);
            grid.InsertColumn();
            wb.Recalculate();
            Check("insertar columna: B1 10", "10", Val(s, "B1"));
            Check("insertar columna: fórmula C1", "=B1+B2", Raw(s, "C1"));

            grid.SelectCell(0, 0);
            grid.DeleteColumn();
            wb.Recalculate();
            Check("eliminar columna: A1 10", "10", Val(s, "A1"));
            Check("eliminar columna: fórmula B1", "=A1+A2", Raw(s, "B1"));

            Set(s, "A1", "3");
            Set(s, "A2", "1");
            Set(s, "A3", "2");
            grid.SelectRange(0, 0, 2, 0);
            grid.SortSelection(true);
            wb.Recalculate();
            Check("ordenar: A1", "1", Val(s, "A1"));
            Check("ordenar: A2", "2", Val(s, "A2"));
            Check("ordenar: A3", "3", Val(s, "A3"));

            Set(s, "A1", "5");
            Set(s, "A2", "2");
            Set(s, "A3", "3");
            grid.SelectCell(3, 0);
            grid.AutoSum();
            wb.Recalculate();
            Check("autosuma", "=SUMA(A1:A3)", Raw(s, "A4"));
            Check("autosuma valor", "10", Val(s, "A4"));

            Set(s, "A1", "hola");
            grid.SelectCell(0, 0);
            int fr, fc;
            bool found = grid.FindNext("hola", false, true, false, 10, 0, out fr, out fc);
            Check("buscar encontrado", "True", found.ToString());
            Check("buscar fila", "0", fr.ToString());

            Workbook wb2 = new Workbook();
            Sheet s2 = wb2.Sheets[0];
            Set(s2, "A1", "1");
            Set(s2, "A2", "9");
            grid.Attach(wb2, 0);
            grid.SelectRange(0, 0, 0, 0);
            grid.ApplyStyle(delegate(CellStyle st) { st.Bold = true; st.FillColor = System.Drawing.Color.Yellow; });
            Check("formato negrita", "True", s2.GetCell(0, 0).EffectiveStyle.Bold.ToString());
            Check("formato relleno", "-256", s2.GetCell(0, 0).EffectiveStyle.FillColor.ToArgb().ToString());

            Set(s2, "A1", "1");
            Set(s2, "A2", "2");
            Set(s2, "B1", "=A1*2");
            Set(s2, "B2", "=A2*2");
            grid.SelectRange(0, 0, 1, 1);
            grid.FillDown();
            wb2.Recalculate();
            Check("rellenar abajo: A2", "1", Val(s2, "A2"));
            Check("rellenar abajo: B2 fórmula", "=A2*2", Raw(s2, "B2"));
            Check("rellenar abajo: B2 valor", "2", Val(s2, "B2"));

            Set(s2, "A1", "7");
            Set(s2, "B1", "=A1*2");
            grid.SelectRange(0, 0, 0, 1);
            grid.CopySelection(false);
            grid.SelectCell(2, 0);
            grid.PasteClipboard();
            wb2.Recalculate();
            Check("copiar/pegar: A3", "7", Val(s2, "A3"));
            Check("copiar/pegar: B3 fórmula", "=A3*2", Raw(s2, "B3"));
            Check("copiar/pegar: B3 valor", "14", Val(s2, "B3"));

            grid.SelectRange(5, 0, 6, 2);
            grid.MergeSelection();
            Check("combinar: rango", "A6:C7", s2.Merges.Count > 0 ? s2.Merges[0].ToRangeString() : "");
            grid.SelectCell(6, 2);
            Check("combinar: selección expandida", "A6:C7", grid.SelectionName);
            grid.SelectCell(5, 0);
            grid.InsertRow();
            Check("combinar: desplazada", "A8:C9", s2.Merges.Count > 0 ? s2.Merges[0].ToRangeString() : "");
            grid.SelectCell(7, 0);
            grid.UnmergeSelection();
            Check("separar: sin fusiones", "0", s2.Merges.Count.ToString());

            Workbook wb3 = new Workbook();
            Sheet s3 = wb3.Sheets[0];
            Set(s3, "A1", "10");
            Set(s3, "A2", "20");
            Set(s3, "A3", "30");
            Set(s3, "A4", "=SUMA(A1:A3)");
            Set(s3, "A5", "5");
            Set(s3, "A6", "7");
            wb3.Recalculate();
            grid.Attach(wb3, 0);
            grid.SelectCell(6, 0);
            grid.AutoSum();
            wb3.Recalculate();
            Check("autosuma no re-suma subtotales", "=SUMA(A4:A6)", Raw(s3, "A7"));
            Check("autosuma subtotal valor", "72", Val(s3, "A7"));

            Workbook wb4 = new Workbook();
            Sheet s4 = wb4.Sheets[0];
            Set(s4, "A1", "10");
            Set(s4, "A2", "20");
            wb4.Recalculate();
            grid.Attach(wb4, 0);
            grid.SelectRange(0, 0, 1, 0);
            grid.AutoSum();
            wb4.Recalculate();
            Check("autosuma de selección", "=SUMA(A1:A2)", Raw(s4, "A3"));
            Check("autosuma selección valor", "30", Val(s4, "A3"));

            Workbook wb5 = new Workbook();
            Sheet s5 = wb5.Sheets[0];
            Set(s5, "A1", "Nombre"); Set(s5, "A2", "Ana"); Set(s5, "A3", "Beto");
            Set(s5, "A4", "Carla");
            grid.Attach(wb5, 0);
            grid.SelectRange(0, 0, 3, 0);
            grid.ApplyFilter(0, 0, "a");
            Check("autofiltro oculta fila", "True", s5.HiddenRows.Contains(2).ToString());
            grid.ClearFilter();
            Check("autofiltro limpia", "0", s5.HiddenRows.Count.ToString());
            grid.ApplyValidation(1, 0, 3, 0, new string[] { "Ana", "Beto" }, "Lista", true);
            Check("validacion registrada", "2", s5.Validations[0].Items.Count.ToString());
            s5.FreezeRows = 1;
            s5.FreezeCols = 1;
            Check("freeze filas", "1", s5.FreezeRows.ToString());
            Check("freeze columnas", "1", s5.FreezeCols.ToString());
            grid.UnfreezePanes();
            Check("freeze desactivado", "0", s5.FreezeRows.ToString());

            Workbook wb6 = new Workbook();
            Sheet s6 = wb6.Sheets[0];
            Set(s6, "A1", "combinada"); Set(s6, "C1", "destino");
            grid.Attach(wb6, 0);
            grid.SelectRange(0, 0, 0, 1);
            grid.MergeSelection();
            grid.SelectCell(0, 1);
            Check("clic interior combinada", "A1:B1", grid.SelectionName);
            Check("activa combinada", "A1", grid.ActiveCellName);
            grid.MoveActiveCell(0, 1);
            Check("salir de celda combinada", "C1", grid.ActiveCellName);
        }
        Console.WriteLine("GridLogic: " + _ok + " correctas, " + _bad + " fallidas.");
        return _bad == 0 ? 0 : 1;
    }
}
