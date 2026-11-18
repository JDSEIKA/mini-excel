using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace CalcRapido
{
    public static class SelfTest
    {
        private static int _passed;
        private static int _failed;
        private static readonly List<string> _failures = new List<string>();

        public static int Run()
        {
            _passed = 0; _failed = 0; _failures.Clear();
            TestArithmetic();
            TestAggregates();
            TestLogic();
            TestText();
            TestLookup();
            TestDates();
            TestErrorsAndCycles();
            TestRefsAndTranslation();
            TestFormatting();
            TestParsing();
            TestIo();
            Console.WriteLine();
            Console.WriteLine("Pruebas: " + _passed + " correctas, " + _failed + " fallidas.");
            foreach (string f in _failures) Console.WriteLine("  FALLO: " + f);
            return _failed == 0 ? 0 : 1;
        }

        private static Workbook Book()
        {
            Workbook wb = new Workbook();
            wb.Sheets[0].Name = "Hoja1";
            return wb;
        }

        private static void Set(Sheet s, string addr, string text)
        {
            int r, c; bool ar, ac;
            if (!AddressUtil.TryParseCell(addr, out r, out c, out ar, out ac)) throw new Exception("Dirección mala: " + addr);
            s.SetInput(r, c, text);
        }

        private static CellValue Calc(Workbook wb, string formula)
        {
            Cell c = wb.ActiveSheet.GetOrCreate(500, 20);
            c.SetInput(formula);
            wb.Recalculate();
            return c.Value;
        }

        private static void Check(string name, string expected, string actual)
        {
            if (string.Equals(expected, actual, StringComparison.Ordinal))
            {
                _passed++;
            }
            else
            {
                _failed++;
                _failures.Add(name + " -> esperado [" + expected + "] obtenido [" + actual + "]");
            }
        }

        private static void CheckNum(string name, double expected, CellValue v)
        {
            if (v.Kind == ValueKind.Number && Math.Abs(v.Number - expected) < 1e-8)
                _passed++;
            else
            {
                _failed++;
                _failures.Add(name + " -> esperado [" + expected.ToString("G", CultureInfo.InvariantCulture) + "] obtenido [" + v.ToDisplayString() + "]");
            }
        }

        private static void CheckErr(string name, string expected, CellValue v)
        {
            if (v.Kind == ValueKind.Error && v.Error == expected) _passed++;
            else
            {
                _failed++;
                _failures.Add(name + " -> esperado error [" + expected + "] obtenido [" + v.ToDisplayString() + "]");
            }
        }

        private static void TestArithmetic()
        {
            Workbook wb = Book();
            CheckNum("precedencia", 7, Calc(wb, "=1+2*3"));
            CheckNum("paréntesis", 9, Calc(wb, "=(1+2)*3"));
            CheckNum("potencia", 8, Calc(wb, "=2^3"));
            CheckNum("porcentaje", 0.5, Calc(wb, "=50%"));
            CheckNum("negativo", -9, Calc(wb, "=-(4+5)"));
            CheckNum("división", 2.5, Calc(wb, "=5/2"));
            Check("comparación", "VERDADERO", Calc(wb, "=2>1").ToDisplayString());
        }

        private static void TestAggregates()
        {
            Workbook wb = Book();
            Sheet s = wb.Sheets[0];
            Set(s, "A1", "1"); Set(s, "A2", "2"); Set(s, "A3", "3"); Set(s, "A4", "4");
            CheckNum("SUMA", 10, Calc(wb, "=SUMA(A1:A4)"));
            CheckNum("SUM", 10, Calc(wb, "=SUM(A1:A4)"));
            CheckNum("PROMEDIO", 2.5, Calc(wb, "=PROMEDIO(A1:A4)"));
            CheckNum("MIN", 1, Calc(wb, "=MIN(A1:A4)"));
            CheckNum("MAX", 4, Calc(wb, "=MAX(A1:A4)"));
            CheckNum("CONTAR", 4, Calc(wb, "=CONTAR(A1:A4)"));
            CheckNum("SUMAR.SI", 7, Calc(wb, "=SUMAR.SI(A1:A4;\">2\")"));
            CheckNum("CONTAR.SI", 2, Calc(wb, "=CONTAR.SI(A1:A4;\">2\")"));
            CheckNum("PROMEDIO.SI", 3.5, Calc(wb, "=PROMEDIO.SI(A1:A4;\">2\")"));
            CheckNum("REDONDEAR", 3.46, Calc(wb, "=REDONDEAR(3.456;2)"));
            CheckNum("ENTERO", 3, Calc(wb, "=ENTERO(3.9)"));
            CheckNum("RESIDUO", 1, Calc(wb, "=RESIDUO(7;3)"));
            CheckNum("RAIZ", 3, Calc(wb, "=RAIZ(9)"));
            CheckNum("SUMAPRODUCTO", 30, Calc(wb, "=SUMAPRODUCTO(A1:A4;A1:A4)"));
            CheckNum("columna completa", 10, Calc(wb, "=SUMA(A:A)"));
            CheckNum("función con punto y coma", 3, Calc(wb, "=SUMA(1;2)"));
            CheckNum("COUNTBLANK escalar", 1, Calc(wb, "=COUNTBLANK(Z1)"));
            CheckNum("TEXT porcentaje", 0, Calc(wb, "=SI(TEXTO(0,5;\"0%\")=\"50%\";0;1)"));
            CheckNum("AVERAGEIF", 3.5, Calc(wb, "=PROMEDIO.SI(A1:A4;\">2\")"));
        }

        private static void TestLogic()
        {
            Workbook wb = Book();
            Set(wb.Sheets[0], "A1", "10");
            Check("SI verdadero", "sí", Calc(wb, "=SI(A1>5;\"sí\";\"no\")").Text);
            Check("SI falso", "no", Calc(wb, "=SI(A1<5;\"sí\";\"no\")").Text);
            Check("SI.ERROR", "99", Calc(wb, "=SI.ERROR(1/0;99)").ToDisplayString());
            Check("Y", "FALSO", Calc(wb, "=Y(A1>5;A1<20;A1=9)").ToDisplayString());
            Check("O", "VERDADERO", Calc(wb, "=O(A1=10;A1=0)").ToDisplayString());
            Check("NO", "FALSO", Calc(wb, "=NO(VERDADERO)").ToDisplayString());
            Check("ESBLANCO", "VERDADERO", Calc(wb, "=ESBLANCO(Z1)").ToDisplayString());
        }

        private static void TestText()
        {
            Workbook wb = Book();
            Check("MAYUSC", "HOLA", Calc(wb, "=MAYUSC(\"hola\")").Text);
            Check("IZQUIERDA", "ab", Calc(wb, "=IZQUIERDA(\"abcdef\";2)").Text);
            Check("DERECHA", "ef", Calc(wb, "=DERECHA(\"abcdef\";2)").Text);
            Check("EXTRAE", "cde", Calc(wb, "=EXTRAE(\"abcdef\";3;3)").Text);
            Check("CONCATENAR", "A2", Calc(wb, "=CONCATENAR(\"A\";1+1)").Text);
            Check("LARGO", "5", Calc(wb, "=LARGO(\"hoja \")").ToDisplayString());
            Check("TEXTO", "1.234,50", Calc(wb, "=TEXTO(1234.5;\"#,##0.00\")").Text);
            Check("HALLAR", "3", Calc(wb, "=HALLAR(\"c\";\"abc\")").ToDisplayString());
        }

        private static void TestLookup()
        {
            Workbook wb = Book();
            Sheet s = wb.Sheets[0];
            Set(s, "A1", "a"); Set(s, "B1", "10");
            Set(s, "A2", "b"); Set(s, "B2", "20");
            Set(s, "A3", "c"); Set(s, "B3", "30");
            CheckNum("BUSCARV exacto", 20, Calc(wb, "=BUSCARV(\"b\";A1:B3;2;FALSO)"));
            CheckNum("BUSCARV aprox", 30, Calc(wb, "=BUSCARV(\"c\";A1:B3;2)"));
            CheckNum("COINCIDIR", 2, Calc(wb, "=COINCIDIR(\"b\";A1:A3;0)"));
            CheckNum("INDICE", 30, Calc(wb, "=INDICE(B1:B3;3)"));
            CheckNum("ELEGIR", 9, Calc(wb, "=ELEGIR(2;7;9;11)"));
        }

        private static void TestDates()
        {
            Workbook wb = Book();
            CheckNum("AÑO", 2026, Calc(wb, "=AÑO(FECHA(2026;11;9))"));
            CheckNum("MES", 11, Calc(wb, "=MES(FECHA(2026;11;9))"));
            CheckNum("DIA", 9, Calc(wb, "=DIA(FECHA(2026;11;9))"));
            CellValue v = Calc(wb, "=HOY()");
            Check("HOY es fecha", "True", (v.Kind == ValueKind.Number && v.Number > 40000).ToString());
        }

        private static void TestErrorsAndCycles()
        {
            Workbook wb = Book();
            CheckErr("div cero", Errors.Div0, Calc(wb, "=1/0"));
            CheckErr("nombre", Errors.Name, Calc(wb, "=NOEXISTE(1)"));
            string err;
            Check("sintaxis inválida", "False", FormulaEngine.Validate("=SUM(1;2", out err).ToString());
            Check("sintaxis válida", "True", FormulaEngine.Validate("=SUM(1;2)", out err).ToString());

            Sheet s = wb.Sheets[0];
            Set(s, "A1", "=A2");
            Set(s, "A2", "=A1");
            wb.Recalculate();
            CheckErr("circular", Errors.Circular, s.GetCell(0, 0).Value);
        }

        private static void TestRefsAndTranslation()
        {
            Workbook wb = Book();
            Sheet s = wb.Sheets[0];
            Set(s, "B1", "5");
            CheckNum("absoluta", 6, Calc(wb, "=$B$1+1"));
            Sheet s2 = wb.AddSheet("Hoja2");
            Set(s2, "A1", "7");
            CheckNum("otra hoja", 14, Calc(wb, "=Hoja2!A1*2"));

            string t = FormulaEngine.TranslateFormula("=SUM($A1:B2)+C$3", 1, 1);
            Check("traslado", "=SUM($A2:C3)+D$3", t);

            string sh = FormulaEngine.ShiftFormula("=A5+B3", "Hoja1", "Hoja1", true, true, 3, 1);
            Check("insertar fila", "=A6+B3", sh);
            string sh2 = FormulaEngine.ShiftFormula("=A5+B1", "Hoja1", "Hoja1", false, true, 2, 2);
            Check("eliminar filas", "=A3+B1", sh2);
            Check("renombrar hoja", "=HojaNueva!A1+SUMA(HojaNueva!A2:A3)",
                FormulaEngine.RenameSheetReferences("=Hoja2!A1+SUMA(Hoja2!A2:A3)", "Hoja2", "HojaNueva"));
            Check("renombrar hoja con espacios", "='Otra Hoja'!A1",
                FormulaEngine.RenameSheetReferences("='Mi Hoja'!A1", "Mi Hoja", "Otra Hoja"));
        }

        private static void TestParsing()
        {
            double d;
            Check("número 10", "True", NumberParser.TryParse("10", out d).ToString());
            CheckNum("valor 10", 10, CellValue.FromNumber(d));
            Check("número es-CO", "True", NumberParser.TryParse("1.234,56", out d).ToString());
            CheckNum("valor 1.234,56", 1234.56, CellValue.FromNumber(d));
            Check("número en-US", "True", NumberParser.TryParse("1,234.56", out d).ToString());
            CheckNum("valor 1,234.56", 1234.56, CellValue.FromNumber(d));
            Check("porcentaje literal", "True", NumberParser.TryParse("50%", out d).ToString());
            CheckNum("valor 50%", 0.5, CellValue.FromNumber(d));
            Check("negativo paréntesis", "True", NumberParser.TryParse("(123)", out d).ToString());
            CheckNum("valor (123)", -123, CellValue.FromNumber(d));

            Workbook wb = Book();
            Set(wb.Sheets[0], "A1", "12,5%");
            Cell c = wb.Sheets[0].GetCell(0, 0);
            CheckNum("porcentaje automático", 0.125, c.Value);
            Check("formato porcentaje automático", "Percent", c.EffectiveStyle.Format.ToString());
            Check("decimales porcentaje", "1", c.EffectiveStyle.Decimals.ToString());
            Check("texto porcentaje", "12,5%", NumberFormatter.Format(c.Value, c.EffectiveStyle));
        }

        private static void TestIo()
        {
            string dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "CalcRapidoTest_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                Workbook wb = new Workbook();
                Sheet s = wb.Sheets[0];
                s.Name = "Datos";
                Set(s, "A1", "Hola");
                Set(s, "A2", "10");
                Set(s, "B1", "=A2*2");
                CellStyle st = new CellStyle();
                st.Bold = true;
                st.Format = NumberFormatKind.Currency;
                st.FillColor = System.Drawing.Color.FromArgb(255, 240, 200);
                s.SetStyle(1, 1, st);
                s.SetColWidth(0, 160);
                wb.AddSheet("Otra");
                Set(wb.Sheets[1], "A1", "=Datos!A2+5");
                wb.Recalculate();

                string native = Path.Combine(dir, "prueba.crb");
                FileIO.SaveNative(wb, native);
                Workbook back = FileIO.LoadNative(native);
                Check("nativo: hojas", "2", back.Sheets.Count.ToString());
                Check("nativo: nombre", "Datos", back.Sheets[0].Name);
                Check("nativo: texto", "Hola", back.Sheets[0].GetCell(0, 0).Value.Text);
                CheckNum("nativo: fórmula", 20, back.Sheets[0].GetCell(0, 1).Value);
                Check("nativo: negrita", "True", back.Sheets[0].GetCell(1, 1).EffectiveStyle.Bold.ToString());
                CheckNum("nativo: otra hoja", 15, back.Sheets[1].GetCell(0, 0).Value);
                CheckNum("nativo: ancho", 160, CellValue.FromNumber(back.Sheets[0].GetColWidth(0)));

                string csv = Path.Combine(dir, "prueba.csv");
                FileIO.SaveCsv(wb, csv);
                Workbook fromCsv = FileIO.LoadCsv(csv);
                Check("csv: texto", "Hola", fromCsv.Sheets[0].GetCell(0, 0).Value.Text);
                CheckNum("csv: número", 10, fromCsv.Sheets[0].GetCell(1, 0).Value);
                CheckNum("csv: fórmula exportada", 20, fromCsv.Sheets[0].GetCell(0, 1).Value);

                string xlsx = Path.Combine(dir, "prueba.xlsx");
                FileIO.Save(wb, xlsx);
                Workbook fromXlsx = FileIO.Open(xlsx);
                Check("xlsx: texto", "Hola", fromXlsx.Sheets[0].GetCell(0, 0).Value.Text);
                CheckNum("xlsx: fórmula", 20, fromXlsx.Sheets[0].GetCell(0, 1).Value);
            }
            catch (Exception ex)
            {
                _failed++;
                _failures.Add("IO: " + ex.Message);
            }
            finally
            {
                try { Directory.Delete(dir, true); } catch { }
            }
        }

        private static void TestFormatting()
        {
            CellValue v = CellValue.FromNumber(1234.5);
            CellStyle c = new CellStyle(); c.Format = NumberFormatKind.Currency; c.Decimals = 2;
            Check("moneda", "$ 1.234,50", NumberFormatter.Format(v, c));
            c = new CellStyle(); c.Format = NumberFormatKind.Percent; c.Decimals = 1;
            Check("porcentaje", "50,0%", NumberFormatter.Format(CellValue.FromNumber(0.5), c));
            c = new CellStyle(); c.Format = NumberFormatKind.Number; c.Decimals = 0;
            Check("miles", "12.346", NumberFormatter.Format(CellValue.FromNumber(12345.6), c));
        }
    }
}
