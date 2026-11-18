using System;
using System.IO;
using CalcRapido;

public static class ChartProbe
{
    private static int _ok, _bad;

    private static void Check(string name, bool ok, string detail)
    {
        if (ok) { _ok++; Console.WriteLine("OK  " + name); }
        else { _bad++; Console.WriteLine("BAD " + name + " [" + detail + "]"); }
    }

    public static int Main()
    {
        string dir = Path.Combine(Directory.GetCurrentDirectory(), "tests", "bin",
            "chartprobe-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        string xlsx = Path.Combine(dir, "chart.xlsx");
        string crb = Path.Combine(dir, "chart.crb");
        try
        {
            Workbook wb = new Workbook();
            Sheet s = wb.Sheets[0];
            s.SetTextLiteral(0, 0, "Producto");
            s.SetTextLiteral(0, 1, "Importe");
            string[] names = { "Alfa", "Beta", "Gamma" };
            double[] values = { 10, 20, 30 };
            for (int i = 0; i < names.Length; i++)
            {
                s.SetTextLiteral(i + 1, 0, names[i]);
                s.SetNumber(i + 1, 1, values[i]);
            }
            SheetChart chart = new SheetChart
            {
                Kind = ChartKind.Column,
                Title = "Ventas",
                ShowLegend = true,
                ShowAxes = true,
                FromRow = 0,
                FromCol = 3,
                WidthPx = 360,
                HeightPx = 220
            };
            chart.Series.Add(new ChartSeries
            {
                NameFormula = "'Hoja1'!$B$1",
                CategoriesFormula = "'Hoja1'!$A$2:$A$4",
                ValuesFormula = "'Hoja1'!$B$2:$B$4"
            });
            s.Charts.Add(chart);
            s.ChartsDirty = true;

            Check("resuelve categorías", ChartData.ReadTexts(wb, s, chart.Series[0].CategoriesFormula).Length == 3,
                "categorías");
            Check("resuelve valores", ChartData.ReadNumbers(wb, s, chart.Series[0].ValuesFormula).Length == 3,
                "valores");

            XlsxIO.SaveNew(wb, xlsx);
            Workbook x = XlsxIO.Load(xlsx);
            Check("xlsx conserva gráfico", x.Sheets[0].Charts.Count == 1 &&
                x.Sheets[0].Charts[0].Series.Count == 1 &&
                x.Sheets[0].Charts[0].Title == "Ventas", "recarga xlsx");

            SheetChart second = new SheetChart
            {
                Kind = ChartKind.Line, Title = "Tendencia", FromRow = 6, FromCol = 3,
                WidthPx = 320, HeightPx = 200
            };
            second.Series.Add(new ChartSeries
            {
                CategoriesFormula = "'Hoja1'!$A$2:$A$4",
                ValuesFormula = "'Hoja1'!$B$2:$B$4"
            });
            x.Sheets[0].Charts.Add(second);
            x.Sheets[0].ChartsDirty = true;
            string xlsx2 = Path.Combine(dir, "chart2.xlsx");
            XlsxIO.Save(x, xlsx2);
            Workbook x2 = XlsxIO.Load(xlsx2);
            Check("agrega gráfico conservando el original", x2.Sheets[0].Charts.Count == 2,
                "gráficos=" + x2.Sheets[0].Charts.Count);

            FileIO.SaveNative(wb, crb);
            Workbook n = FileIO.LoadNative(crb);
            Check("crb conserva gráfico", n.Sheets[0].Charts.Count == 1 &&
                n.Sheets[0].Charts[0].Title == "Ventas", "recarga crb");

            string advanced = Path.Combine(Directory.GetCurrentDirectory(), "tests", "bin", "advanced.xlsx");
            if (File.Exists(advanced))
            {
                Workbook imageBook = XlsxIO.Load(advanced);
                bool hasImage = false;
                foreach (Sheet sh in imageBook.Sheets) if (sh.Images.Count > 0) hasImage = true;
                Check("lee y conserva imágenes del dibujo", hasImage, "imágenes");
            }

            string example = Path.Combine(Directory.GetCurrentDirectory(), "dist", "Ejemplo.xlsx");
            if (File.Exists(example))
            {
                Workbook ex = XlsxIO.Load(example);
                Sheet es = ex.Sheets[0];
                double[] all = ChartData.ReadNumbers(ex, es, es.Charts[0].Series[0].ValuesFormula, false);
                double[] visible = ChartData.ReadNumbers(ex, es, es.Charts[0].Series[0].ValuesFormula, true);
                Check("gráfico excluye filas ocultas", all.Length == 4 && visible.Length == 3,
                    "all=" + all.Length + " visible=" + visible.Length);
            }

            string online = Path.Combine(Directory.GetCurrentDirectory(), "tests", "corpus", "online");
            if (Directory.Exists(online))
            {
                string[] samples = { "column.xlsx", "line.xlsx", "pie.xlsx", "scatter.xlsx",
                                     "area.xlsx", "radar.xlsx", "charts.xlsx" };
                int loaded = 0;
                foreach (string name in samples)
                {
                    string path = Path.Combine(online, name);
                    if (!File.Exists(path)) continue;
                    Workbook sample = XlsxIO.Load(path);
                    if (sample.Sheets.Count > 0 && sample.Sheets[0].Charts.Count == 1) loaded++;
                }
                Check("gráficos de plantillas reales", loaded == samples.Length, "cargados=" + loaded);
            }
        }
        catch (Exception ex)
        {
            _bad++;
            Console.WriteLine("BAD excepción: " + ex.Message);
        }
        finally
        {
            try { Directory.Delete(dir, true); } catch { }
        }
        Console.WriteLine("Charts: " + _ok + " correctas, " + _bad + " fallidas.");
        return _bad == 0 ? 0 : 1;
    }
}
