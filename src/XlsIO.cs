using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace CalcRapido
{
    // Lector de libros Excel 97-2003 (.xls, BIFF8). No ejecuta macros ni
    // interpreta el contenedor como un ZIP: extrae el flujo Workbook del
    // archivo OLE2 y reconstruye celdas, estilos, fórmulas y estructura.
    internal static class XlsIO
    {
        private const int MaxCells = 2000000;
        private const int MaxCellsTotal = 4000000;
        private const int MaxSheets = 512;
        private const int MaxSstStrings = 1000000;
        private const int MaxMergeRows = 10000;
        private const int MaxMergeCols = 1000;

        public static Workbook Load(string path)
        {
            byte[] stream = Ole2.ReadWorkbookStream(path);
            XlsContext ctx = ReadGlobals(stream);
            if (ctx.Encrypted) throw new FileFormatException("El archivo .xls está protegido con contraseña.");
            Workbook wb = new Workbook();
            wb.Sheets.Clear();
            wb.Date1904 = ctx.Date1904;
            int index = 0;
            long totalCells = 0;
            foreach (XlsSheetInfo si in ctx.Sheets)
            {
                Sheet sheet = new Sheet(wb, MakeUniqueName(wb, si.Name));
                ParseSheet(stream, si.Offset, sheet, ctx);
                totalCells += sheet.Cells.Count;
                if (totalCells > MaxCellsTotal)
                    throw new FileFormatException("El archivo .xls contiene demasiadas celdas en total.");
                sheet.RecalculateUsedRange();
                wb.Sheets.Add(sheet);
                index++;
            }
            if (wb.Sheets.Count == 0) wb.AddSheet("Hoja1");
            foreach (XlsDefinedName xn in ctx.DefinedNames)
            {
                int local = xn.LocalSheetId;
                if (local >= wb.Sheets.Count) local = -1;
                wb.DefinedNames.Add(new DefinedName
                {
                    Name = xn.Name,
                    Formula = xn.Formula,
                    LocalSheetId = local
                });
            }
            wb.ActiveSheetIndex = Math.Max(0, Math.Min(wb.Sheets.Count - 1, ctx.ActiveSheet));
            wb.LegacyXlsPath = path;
            wb.AddLoadWarning("Libro Excel 97-2003 (.xls): se abre en modo compatible de solo lectura. " +
                "Macros, imágenes, gráficos, comentarios y formatos condicionales no se conservan al convertir.");
            wb.Recalculate();
            return wb;
        }

        private sealed class XlsContext
        {
            public List<XlsSheetInfo> Sheets = new List<XlsSheetInfo>();
            public List<byte[]> Fonts = new List<byte[]>();
            public List<byte[]> Xfs = new List<byte[]>();
            public Dictionary<int, string> Formats = new Dictionary<int, string>();
            public List<string> Strings = new List<string>();
            public int[] Palette = DefaultPalette();
            public CellStyle[] Styles;
            public bool Date1904;
            public bool Encrypted;
            public int ActiveSheet;
            public List<XlsNameRecord> NameRecords = new List<XlsNameRecord>();
            public List<string> DefinedNameNames = new List<string>();
            public List<XlsDefinedName> DefinedNames = new List<XlsDefinedName>();
            public List<int[]> ExternalSheets = new List<int[]>();
            public List<string> FormulaSheetNames = new List<string>();
        }

        private sealed class XlsNameRecord
        {
            public string Name;
            public int SheetIndex;
            public bool Builtin;
            public byte[] Formula;
        }

        private sealed class XlsDefinedName
        {
            public string Name;
            public string Formula;
            public int LocalSheetId;
        }

        private sealed class XlsSharedFormula
        {
            public int Row1, Col1, Row2, Col2;
            public int MasterRow, MasterCol;
            public string Formula;
        }

        private sealed class XlsSheetInfo
        {
            public int Offset;
            public string Name;
            public bool Hidden;
        }

        private static XlsContext ReadGlobals(byte[] data)
        {
            XlsContext ctx = new XlsContext();
            int p = 0;
            bool first = true;
            while (p + 4 <= data.Length)
            {
                int id = U16(data, p);
                int len = U16(data, p + 2);
                if (p + 4 + len > data.Length) break;
                int d = p + 4;
                if (first)
                {
                    if (id != 0x0809) throw new FileFormatException("El archivo no es un libro BIFF8 de Excel.");
                    int vers = U16(data, d);
                    if (vers != 0x0600) throw new FileFormatException("Solo se admiten libros de Excel 97-2003 (BIFF8).");
                    first = false;
                }
                switch (id)
                {
                    case 0x0085: // BOUNDSHEET
                        {
                            if (ctx.Sheets.Count >= MaxSheets)
                                throw new FileFormatException("El archivo .xls contiene demasiadas hojas.");
                            long offLong = U32(data, d);
                            if (offLong >= data.Length)
                                throw new FileFormatException("El archivo .xls contiene una hoja fuera de rango.");
                            int off = (int)offLong;
                            int vis = d + 4 < data.Length ? data[d + 4] : 0;
                            string name = ReadShortString(data, d + 6);
                            ctx.Sheets.Add(new XlsSheetInfo { Offset = off, Name = name, Hidden = vis != 0 });
                            break;
                        }
                    case 0x002F: // FILEPASS
                        ctx.Encrypted = true;
                        break;
                    case 0x00FC: // SST + CONTINUE
                        {
                            List<byte[]> blocks = new List<byte[]>();
                            byte[] firstBlock = new byte[len];
                            Buffer.BlockCopy(data, d, firstBlock, 0, len);
                            blocks.Add(firstBlock);
                            int q = p + 4 + len;
                            while (q + 4 <= data.Length && U16(data, q) == 0x003C)
                            {
                                int cl = U16(data, q + 2);
                                byte[] block = new byte[cl];
                                Buffer.BlockCopy(data, q + 4, block, 0, Math.Min(cl, data.Length - (q + 4)));
                                blocks.Add(block);
                                q += 4 + cl;
                            }
                            p = q - 4 - len;
                            ctx.Strings = ParseSst(blocks);
                            break;
                        }
                    case 0x0031: // FONT
                        {
                            byte[] b = new byte[len];
                            Buffer.BlockCopy(data, d, b, 0, len);
                            ctx.Fonts.Add(b);
                            break;
                        }
                    case 0x0017: // EXTERNSHEET (mapa ixti -> hoja)
                        {
                            if (len < 2) break;
                            int count = U16(data, d);
                            for (int i = 0; i < count && d + 2 + i * 6 + 6 <= d + len; i++)
                                ctx.ExternalSheets.Add(new int[]
                                {
                                    U16(data, d + 2 + i * 6),
                                    U16(data, d + 4 + i * 6),
                                    U16(data, d + 6 + i * 6)
                                });
                            break;
                        }
                    case 0x0018: // NAME
                        {
                            if (len < 14) break;
                            int flags = U16(data, d);
                            int nameLen = data[d + 3];
                            int formulaLen = U16(data, d + 4);
                            int sheetIndex = U16(data, d + 8);
                            if (nameLen > 255 || formulaLen > 8192) break;
                            int q = d + 14;
                            if (q >= d + len) break;
                            int enc = data[q++];
                            bool high = (enc & 1) != 0;
                            StringBuilder name = new StringBuilder(nameLen);
                            for (int i = 0; i < nameLen && q < d + len; i++)
                            {
                                if (high) { name.Append((char)U16(data, q)); q += 2; }
                                else name.Append((char)data[q++]);
                            }
                            if (q + formulaLen > d + len) break;
                            byte[] formula = new byte[formulaLen];
                            Buffer.BlockCopy(data, q, formula, 0, formulaLen);
                            bool builtin = (flags & 0x20) != 0;
                            ctx.NameRecords.Add(new XlsNameRecord
                            {
                                Name = name.ToString(),
                                SheetIndex = sheetIndex,
                                Builtin = builtin,
                                Formula = formula
                            });
                            break;
                        }
                    case 0x041E: // FORMAT
                        {
                            int idFmt = U16(data, d);
                            string code = ReadUnicodeString(data, d + 2);
                            ctx.Formats[idFmt] = code;
                            break;
                        }
                    case 0x00E0: // XF
                        {
                            byte[] b = new byte[len];
                            Buffer.BlockCopy(data, d, b, 0, len);
                            ctx.Xfs.Add(b);
                            break;
                        }
                    case 0x0092: // PALETTE
                        {
                            int count = U16(data, d);
                            int[] pal = DefaultPalette();
                            for (int i = 0; i < count && d + 2 + i * 4 + 4 <= data.Length; i++)
                            {
                                if (i + 8 < pal.Length)
                                    pal[i + 8] = (data[d + 2 + i * 4] << 16) |
                                                 (data[d + 3 + i * 4] << 8) |
                                                  data[d + 4 + i * 4];
                            }
                            ctx.Palette = pal;
                            break;
                        }
                    case 0x0022: // DATEMODE
                        ctx.Date1904 = U16(data, d) != 0;
                        break;
                    case 0x003D: // WINDOW1
                        if (len >= 12) ctx.ActiveSheet = Math.Max(0, U16(data, d + 10));
                        break;
                }
                p += 4 + len;
                if (id == 0x000A) break;
            }
            bool hasDummy = false;
            foreach (byte[] f in ctx.Fonts)
            {
                if (f.Length < 17) continue;
                int fp = 16;
                string fn = ReadChars(f, ref fp, f[14], (f[15] & 1) != 0);
                if (string.Equals(fn, "Dummy Font", StringComparison.OrdinalIgnoreCase)) { hasDummy = true; break; }
            }
            int maxFontIndex = 0;
            foreach (byte[] xf in ctx.Xfs)
                if (xf.Length >= 2) maxFontIndex = Math.Max(maxFontIndex, U16(xf, 0));
            if (!hasDummy && maxFontIndex >= 4 && ctx.Fonts.Count >= 4)
                ctx.Fonts.Insert(4, DummyFont());
            BuildStyles(ctx);
            foreach (XlsNameRecord nr in ctx.NameRecords)
                ctx.DefinedNameNames.Add(nr.Name ?? "");
            foreach (int[] xti in ctx.ExternalSheets)
            {
                int sup = xti[0], firstTab = xti[1];
                if (sup == 0 && firstTab >= 0 && firstTab < ctx.Sheets.Count && firstTab != 0xFFFF)
                    ctx.FormulaSheetNames.Add(ctx.Sheets[firstTab].Name);
                else
                    ctx.FormulaSheetNames.Add(null);
            }
            if (ctx.FormulaSheetNames.Count == 0)
                foreach (XlsSheetInfo si in ctx.Sheets) ctx.FormulaSheetNames.Add(si.Name);

            foreach (XlsNameRecord nr in ctx.NameRecords)
            {
                if (nr.Builtin || string.IsNullOrEmpty(nr.Name)) continue;
                string formula = XlsFormula.Read(nr.Formula, 0, nr.Formula.Length,
                    ctx.FormulaSheetNames, ctx.DefinedNameNames);
                if (string.IsNullOrEmpty(formula)) continue;
                int local = nr.SheetIndex > 0 ? nr.SheetIndex - 1 : -1;
                ctx.DefinedNames.Add(new XlsDefinedName
                {
                    Name = nr.Name,
                    Formula = formula,
                    LocalSheetId = local
                });
            }
            return ctx;
        }

        private static void ParseSheet(byte[] data, int offset, Sheet sheet, XlsContext ctx)
        {
            int p = offset;
            int guard = 0;
            Cell pendingFormula = null;
            Cell lastFormulaCell = null;
            int lastFormulaRow = -1, lastFormulaCol = -1;
            List<XlsSharedFormula> sharedFormulas = new List<XlsSharedFormula>();
            while (p + 4 <= data.Length && guard++ < 1000000)
            {
                int id = U16(data, p);
                int len = U16(data, p + 2);
                if (p + 4 + len > data.Length) break;
                int d = p + 4;
                switch (id)
                {
                    case 0x00FD: // LABELSST
                        {
                            if (len < 10) break;
                            int row = U16(data, d), col = U16(data, d + 2), xf = U16(data, d + 4);
                            int idx = (int)U32(data, d + 6);
                            if (idx >= 0 && idx < ctx.Strings.Count)
                                PutText(sheet, row, col, xf, ctx.Strings[idx], ctx);
                            break;
                        }
                    case 0x0204: // LABEL (puede continuar en registros CONTINUE)
                        {
                            if (len < 9) break;
                            int row = U16(data, d), col = U16(data, d + 2), xf = U16(data, d + 4);
                            int next;
                            string text = ReadLabelString(data, d, len, out next);
                            PutText(sheet, row, col, xf, text, ctx);
                            p = next;
                            continue;
                        }
                    case 0x0203: // NUMBER
                        {
                            if (len < 14) break;
                            int row = U16(data, d), col = U16(data, d + 2), xf = U16(data, d + 4);
                            double value = BitConverter.ToDouble(data, d + 6);
                            PutNumber(sheet, row, col, xf, value, ctx);
                            break;
                        }
                    case 0x027E: // RK
                        {
                            if (len < 10) break;
                            int row = U16(data, d), col = U16(data, d + 2), xf = U16(data, d + 4);
                            PutNumber(sheet, row, col, xf, DecodeRk(U32(data, d + 6)), ctx);
                            break;
                        }
                    case 0x00BD: // MULRK
                        {
                            if (len < 6) break;
                            int row = U16(data, d);
                            int first = U16(data, d + 2);
                            int count = (len - 6) / 6;
                            for (int i = 0; i < count; i++)
                            {
                                int xf = U16(data, d + 4 + i * 6);
                                double value = DecodeRk(U32(data, d + 6 + i * 6));
                                PutNumber(sheet, row, first + i, xf, value, ctx);
                            }
                            break;
                        }
                    case 0x0201: // BLANK
                        {
                            int row = U16(data, d), col = U16(data, d + 2), xf = U16(data, d + 4);
                            if (ValidCell(row, col) && xf > 0 && xf < ctx.Styles.Length && ctx.Styles[xf] != null)
                                sheet.SetStyle(row, col, ctx.Styles[xf].Clone());
                            break;
                        }
                    case 0x00BE: // MULBLANK
                        {
                            if (len < 6) break;
                            int row = U16(data, d);
                            int first = U16(data, d + 2);
                            int count = (len - 6) / 2;
                            for (int i = 0; i < count; i++)
                            {
                                int xf = U16(data, d + 4 + i * 2);
                                if (ValidCell(row, first + i) && xf > 0 && xf < ctx.Styles.Length && ctx.Styles[xf] != null)
                                    sheet.SetStyle(row, first + i, ctx.Styles[xf].Clone());
                            }
                            break;
                        }
                    case 0x04BC: // SHRFMLA
                        {
                            if (len < 10 || lastFormulaCell == null) break;
                            int r1 = U16(data, d), r2 = U16(data, d + 2);
                            int c1 = data[d + 4], c2 = data[d + 5];
                            int cce = U16(data, d + 8);
                            if (cce < 0 || cce > len - 10) break;
                            string sharedFormula = XlsFormula.Read(data, d + 10, cce,
                                ctx.FormulaSheetNames, ctx.DefinedNameNames, r1, c1);
                            if (string.IsNullOrEmpty(sharedFormula)) break;
                            if (!PlausibleSharedFormula(sharedFormula, sheet)) break;
                            CellValue cached = lastFormulaCell.HasCachedValue
                                ? lastFormulaCell.CachedValue : lastFormulaCell.Value;
                            lastFormulaCell.SetInput("=" + sharedFormula);
                            if (cached.Kind != ValueKind.Empty)
                            {
                                lastFormulaCell.CachedValue = cached;
                                lastFormulaCell.HasCachedValue = true;
                            }
                            sharedFormulas.Add(new XlsSharedFormula
                            {
                                Row1 = r1, Row2 = r2, Col1 = c1, Col2 = c2,
                                MasterRow = lastFormulaRow, MasterCol = lastFormulaCol,
                                Formula = sharedFormula
                            });
                            lastFormulaCell = null;
                            break;
                        }
                    case 0x0006: // FORMULA
                        {
                            int row = U16(data, d), col = U16(data, d + 2), xf = U16(data, d + 4);
                            if (!ValidCell(row, col) || len < 22) break;
                            int cce = U16(data, d + 20);
                            if (cce < 0 || cce > len - 22) break;
                            string formula = null;
                            try
                            {
                                if (cce == 5 && data[d + 22] == 0x01)
                                {
                                    foreach (XlsSharedFormula sf in sharedFormulas)
                                    {
                                        if (row < sf.Row1 || row > sf.Row2 ||
                                            col < sf.Col1 || col > sf.Col2) continue;
                                        formula = FormulaEngine.TranslateFormula(
                                            "=" + sf.Formula,
                                            row - sf.MasterRow, col - sf.MasterCol).TrimStart('=');
                                        if (!PlausibleSharedFormula(formula, sheet)) formula = null;
                                        break;
                                    }
                                }
                                if (formula == null)
                                    formula = XlsFormula.Read(data, d + 22, cce,
                                        ctx.FormulaSheetNames, ctx.DefinedNameNames, row, col);
                            }
                            catch { formula = null; }
                            Cell cell;
                            CellValue cached = FormulaResult(data, d + 6);
                            if (!string.IsNullOrEmpty(formula))
                            {
                                sheet.SetInput(row, col, "=" + formula);
                                cell = sheet.GetCell(row, col);
                                if (cached.Kind != ValueKind.Empty)
                                {
                                    cell.CachedValue = cached;
                                    cell.HasCachedValue = true;
                                }
                            }
                            else
                            {
                                cell = sheet.GetOrCreate(row, col);
                                if (cached.Kind == ValueKind.Number) cell.SetInput(cached.Number.ToString(CultureInfo.InvariantCulture));
                                else if (cached.Kind == ValueKind.Text) cell.SetInput(cached.Text);
                                else if (cached.Kind == ValueKind.Bool) cell.SetInput(cached.Bool ? "VERDADERO" : "FALSO");
                            }
                            ApplyStyle(sheet, row, col, xf, ctx);
                            if (cell != null) cell.Dirty = false;
                            lastFormulaCell = cell;
                            lastFormulaRow = row;
                            lastFormulaCol = col;
                            pendingFormula = cell;
                            break;
                        }
                    case 0x0207: // STRING (resultado de fórmula de texto)
                        {
                            int next;
                            string text = ReadLabelString(data, d, len, out next);
                            if (pendingFormula != null)
                            {
                                if (pendingFormula.IsFormula)
                                {
                                    // STRING es el valor cacheado de una fórmula
                                    // de texto; no debe borrar la fórmula.
                                    pendingFormula.Value = CellValue.FromText(text);
                                    pendingFormula.CachedValue = CellValue.FromText(text);
                                    pendingFormula.HasCachedValue = true;
                                    pendingFormula.Dirty = false;
                                }
                                else pendingFormula.SetInput("'" + text);
                                pendingFormula = null;
                            }
                            p = next;
                            continue;
                        }
                    case 0x000A:
                        return;
                    case 0x00E5: // MERGEDCELLS
                        {
                            if (len < 2) break;
                            int count = U16(data, d);
                            if (count > 10000) count = 10000;
                            int recordEnd = d + len;
                            for (int i = 0; i < count; i++)
                            {
                                int q = d + 2 + i * 8;
                                if (q + 8 > recordEnd) break;
                                int r1 = U16(data, q), r2 = U16(data, q + 2);
                                int c1 = U16(data, q + 4), c2 = U16(data, q + 6);
                                int rr1 = Math.Min(r1, r2), rr2 = Math.Max(r1, r2);
                                int cc1 = Math.Min(c1, c2), cc2 = Math.Max(c1, c2);
                                if (cc2 > 255 || rr2 >= Limits.MaxRows) continue;
                                if ((long)(rr2 - rr1 + 1) > MaxMergeRows || (long)(cc2 - cc1 + 1) > MaxMergeCols)
                                    continue;
                                sheet.Merges.Add(new MergeRange(rr1, cc1, rr2, cc2));
                            }
                            sheet.RebuildMergeLookup();
                            sheet.MergesDirty = false;
                            break;
                        }
                    case 0x007D: // COLINFO
                        {
                            if (len < 10) break;
                            int c1 = U16(data, d), c2 = U16(data, d + 2);
                            int width = U16(data, d + 4);
                            int flags = U16(data, d + 8);
                            c1 = Math.Max(0, Math.Min(255, c1));
                            c2 = Math.Max(c1, Math.Min(255, c2));
                            int px = Math.Max(24, (int)Math.Round(width / 256.0 * 7.5));
                            for (int c = c1; c <= c2; c++)
                            {
                                sheet.SetColWidthImported(c, px);
                                sheet.SetColumnHiddenImported(c, (flags & 0x01) != 0);
                            }
                            break;
                        }
                    case 0x0208: // ROW
                        {
                            if (len < 16) break;
                            int row = U16(data, d);
                            int height = U16(data, d + 6);
                            int flags = U16(data, d + 12);
                            if (row < Limits.MaxRows)
                            {
                                if (height > 0) sheet.SetRowHeightPointsImported(row, height / 20f);
                                if ((flags & 0x20) != 0) sheet.SetRowHiddenImported(row, true);
                            }
                            break;
                        }
                    case 0x0041: // PANE
                        {
                            if (len < 8) break;
                            int top = U16(data, d + 4), left = U16(data, d + 6);
                            sheet.FreezeRows = Math.Max(0, Math.Min(Limits.MaxRows - 1, top));
                            sheet.FreezeCols = Math.Max(0, Math.Min(Limits.MaxCols - 1, left));
                            break;
                        }
                }
                if (sheet.Cells.Count > MaxCells)
                    throw new FileFormatException("La hoja .xls contiene demasiadas celdas.");
                p += 4 + len;
            }
        }

        private static bool ValidCell(int row, int col)
        {
            return row >= 0 && row < Limits.MaxRows && col >= 0 && col < Limits.MaxCols;
        }

        private static void PutText(Sheet sheet, int row, int col, int xf, string text, XlsContext ctx)
        {
            if (text == null || !ValidCell(row, col)) return;
            sheet.SetTextLiteral(row, col, text);
            ApplyStyle(sheet, row, col, xf, ctx);
        }

        private static void PutNumber(Sheet sheet, int row, int col, int xf, double value, XlsContext ctx)
        {
            if (!ValidCell(row, col)) return;
            if (ctx.Date1904 && xf >= 0 && xf < ctx.Styles.Length && ctx.Styles[xf] != null &&
                XlsFormula.IsDateStyle(ctx.Styles[xf]))
                value += 1462;
            sheet.SetNumber(row, col, value);
            ApplyStyle(sheet, row, col, xf, ctx);
        }

        private static bool PlausibleSharedFormula(string formula, Sheet sheet)
        {
            int maxCol = Math.Max(64, sheet.UsedMaxCol + 64);
            int maxRow = Math.Max(1024, sheet.UsedMaxRow + 1024);
            try
            {
                foreach (Match m in Regex.Matches(formula, @"(?<![A-Za-z0-9_])([A-Z]{1,3})\$?(\d+)", RegexOptions.None, TimeSpan.FromMilliseconds(250)))
                {
                    int col, row;
                    if (!AddressUtil.TryColumnIndex(m.Groups[1].Value, out col)) return false;
                    if (!int.TryParse(m.Groups[2].Value, out row)) return false;
                    if (col > maxCol || row > maxRow) return false;
                }
            }
            catch (RegexMatchTimeoutException)
            {
                return false;
            }
            return true;
        }

        private static void ApplyStyle(Sheet sheet, int row, int col, int xf, XlsContext ctx)
        {
            if (!ValidCell(row, col)) return;
            if (xf > 0 && xf < ctx.Styles.Length && ctx.Styles[xf] != null)
            {
                Cell cell = sheet.GetOrCreate(row, col);
                cell.Style = ctx.Styles[xf].Clone();
                cell.StyleDirty = false;
                cell.Dirty = false;
            }
        }

        private static CellValue FormulaResult(byte[] data, int p)
        {
            if (data[p + 6] == 0xFF && data[p + 7] == 0xFF)
            {
                int type = data[p];
                if (type == 1) return CellValue.FromBool(data[p + 2] != 0);
                if (type == 2) return CellValue.FromError(ErrorText(data[p + 2]));
                if (type == 0) return CellValue.Empty;
                return CellValue.Empty;
            }
            return CellValue.FromNumber(BitConverter.ToDouble(data, p));
        }

        private static string ErrorText(int code)
        {
            switch (code)
            {
                case 0x00: return Errors.Div0;
                case 0x07: return Errors.Div0;
                case 0x0F: return Errors.Value;
                case 0x17: return Errors.Ref;
                case 0x1D: return Errors.Name;
                case 0x24: return Errors.Num;
                case 0x2A: return Errors.NA;
                default: return Errors.Value;
            }
        }

        private static double DecodeRk(uint rk)
        {
            bool div100 = (rk & 1) != 0;
            bool isInt = (rk & 2) != 0;
            double v;
            if (isInt) v = (int)rk >> 2;
            else
            {
                long bits = ((long)(rk & 0xFFFFFFFC)) << 32;
                v = BitConverter.Int64BitsToDouble(bits);
            }
            return div100 ? v / 100.0 : v;
        }

        private static List<string> ParseSst(List<byte[]> blocks)
        {
            List<string> list = new List<string>();
            if (blocks.Count == 0 || blocks[0].Length < 8) return list;
            int unique = (int)U32(blocks[0], 4);
            if (unique < 0 || unique > MaxSstStrings) throw new FileFormatException("El archivo .xls contiene demasiadas cadenas.");
            SstCursor c = new SstCursor(blocks, 8);
            for (int i = 0; i < unique; i++)
            {
                int cch;
                int flags;
                if (!c.TryReadU16(out cch) || !c.TryReadByte(out flags)) break;
                int rich = 0, ext = 0;
                if ((flags & 0x08) != 0 && !c.TryReadU16(out rich)) break;
                if ((flags & 0x04) != 0)
                {
                    uint e;
                    if (!c.TryReadU32(out e)) break;
                    ext = (int)e;
                    if (ext < 0 || ext > 1048576) throw new FileFormatException("Cadena .xls no válida.");
                }
                if (cch < 0 || cch > 32767) break;
                bool high = (flags & 1) != 0;
                c.BeginString();
                StringBuilder sb = new StringBuilder(cch);
                bool ok = true;
                for (int j = 0; j < cch; j++)
                {
                    int ch;
                    if (!c.TryReadChar(ref high, out ch)) { ok = false; break; }
                    sb.Append((char)ch);
                }
                if (!ok) break;
                if (rich > 0) c.Skip(rich * 4);
                if (ext > 0) c.Skip(ext);
                list.Add(sb.ToString());
            }
            return list;
        }

        private sealed class SstCursor
        {
            private readonly List<byte[]> _blocks;
            private int _block;
            private int _pos;
            private int _stringStartBlock;
            public SstCursor(List<byte[]> blocks, int pos) { _blocks = blocks; _block = 0; _pos = pos; }
            public void BeginString() { _stringStartBlock = _block; }
            private bool Ensure()
            {
                while (_block < _blocks.Count && _pos >= _blocks[_block].Length)
                {
                    _block++;
                    _pos = 0;
                }
                return _block < _blocks.Count && _pos < _blocks[_block].Length;
            }
            public bool TryReadByte(out int value)
            {
                value = 0;
                if (!Ensure()) return false;
                value = _blocks[_block][_pos++];
                return true;
            }
            public bool TryReadU16(out int value)
            {
                value = 0;
                int a, b;
                if (!TryReadByte(out a) || !TryReadByte(out b)) return false;
                value = a | (b << 8);
                return true;
            }
            public bool TryReadU32(out uint value)
            {
                value = 0;
                int a, b, c, d;
                if (!TryReadByte(out a) || !TryReadByte(out b) || !TryReadByte(out c) || !TryReadByte(out d)) return false;
                value = (uint)(a | (b << 8) | (c << 16) | (d << 24));
                return true;
            }
            public bool TryReadChar(ref bool high, out int value)
            {
                value = 0;
                if (!Ensure())
                {
                    return false;
                }
                if (_pos == 0 && _block > _stringStartBlock)
                {
                    // Si el texto continúa en un bloque nuevo, el primer byte
                    // indica la compresión del resto de la cadena, incluso si
                    // la continuación empieza en el primer carácter.
                    int opt = _blocks[_block][_pos++];
                    high = (opt & 1) != 0;
                    if (!Ensure()) return false;
                }
                if (high)
                {
                    int a, b;
                    if (!TryReadByte(out a) || !TryReadByte(out b)) return false;
                    value = a | (b << 8);
                }
                else value = _blocks[_block][_pos++];
                return true;
            }
            public void Skip(int count)
            {
                for (int i = 0; i < count; i++) { int v; if (!TryReadByte(out v)) return; }
            }
        }

        private static string ReadShortString(byte[] data, int p)
        {
            if (p + 2 > data.Length) return "";
            int cch = data[p++];
            int flags = data[p++];
            return ReadChars(data, ref p, cch, (flags & 1) != 0);
        }

        private static string ReadLabelString(byte[] data, int d, int len, out int next)
        {
            List<byte[]> blocks = new List<byte[]>();
            byte[] first = new byte[len];
            Buffer.BlockCopy(data, d, first, 0, len);
            blocks.Add(first);
            int q = d + len;
            while (q + 4 <= data.Length && U16(data, q) == 0x003C)
            {
                int cl = U16(data, q + 2);
                if (q + 4 + cl > data.Length) break;
                byte[] block = new byte[cl];
                Buffer.BlockCopy(data, q + 4, block, 0, cl);
                blocks.Add(block);
                q += 4 + cl;
            }
            next = q;
            SstCursor c = new SstCursor(blocks, 0);
            int cch, flags, rich = 0;
            uint ext = 0;
            if (!c.TryReadU16(out cch) || !c.TryReadByte(out flags)) return "";
            if ((flags & 0x08) != 0) { int r; if (!c.TryReadU16(out r)) return ""; rich = r; }
            if ((flags & 0x04) != 0) { if (!c.TryReadU32(out ext)) return ""; }
            if (cch < 0 || cch > 32767) return "";
            bool high = (flags & 1) != 0;
            c.BeginString();
            StringBuilder sb = new StringBuilder(cch);
            for (int i = 0; i < cch; i++)
            {
                int ch;
                if (!c.TryReadChar(ref high, out ch)) break;
                sb.Append((char)ch);
            }
            if (rich > 0) c.Skip(rich * 4);
            if (ext > 0 && ext <= 1048576) c.Skip((int)ext);
            return sb.ToString();
        }

        private static string ReadUnicodeString(byte[] data, int p)
        {
            if (p + 3 > data.Length) return "";
            int cch = U16(data, p);
            int flags = data[p + 2];
            p += 3;
            if ((flags & 0x08) != 0) p += 2;
            if ((flags & 0x04) != 0) p += 4;
            return ReadChars(data, ref p, cch, (flags & 1) != 0);
        }

        private static string ReadChars(byte[] data, ref int p, int cch, bool high)
        {
            StringBuilder sb = new StringBuilder(cch);
            for (int i = 0; i < cch && p < data.Length; i++)
            {
                if (high)
                {
                    if (p + 2 > data.Length) break;
                    sb.Append((char)U16(data, p)); p += 2;
                }
                else sb.Append((char)data[p++]);
            }
            return sb.ToString();
        }

        private static void BuildStyles(XlsContext ctx)
        {
            ctx.Styles = new CellStyle[ctx.Xfs.Count];
            for (int i = 0; i < ctx.Xfs.Count; i++)
            {
                byte[] xf = ctx.Xfs[i];
                if (xf.Length < 20) continue;
                CellStyle st = new CellStyle();
                int fontIndex = U16(xf, 0);
                int fmtIndex = U16(xf, 2);
                string code;
                if (!ctx.Formats.TryGetValue(fmtIndex, out code)) code = XlsxIO.BuiltInFormat(fmtIndex);
                XlsxIO.ApplyNumberFormat(st, code);
                if (fontIndex >= 0 && fontIndex < ctx.Fonts.Count) ApplyFont(st, ctx.Fonts[fontIndex], ctx);
                int aln = xf[6];
                int horz = aln & 0x07;
                if (horz == 1) st.Align = 0;
                else if (horz == 2) st.Align = 1;
                else if (horz == 3) st.Align = 2;
                uint brd1 = U32(xf, 10);
                uint brd2 = U32(xf, 14);
                int pattern = (int)((brd2 >> 26) & 0x3F);
                int fore = U16(xf, 18) & 0x7F;
                if (pattern != 0 && fore >= 0 && fore < ctx.Palette.Length)
                    st.FillColor = Color.FromArgb(unchecked((int)0xFF000000) | (ctx.Palette[fore] & 0xFFFFFF));
                if ((brd1 & 0x0F) != 0) st.Borders |= BorderFlags.Left;
                if ((brd1 & 0xF0) != 0) st.Borders |= BorderFlags.Right;
                if ((brd1 & 0x0F00) != 0) st.Borders |= BorderFlags.Top;
                if ((brd1 & 0xF000) != 0) st.Borders |= BorderFlags.Bottom;
                ctx.Styles[i] = st;
            }
        }

        private static byte[] DummyFont()
        {
            byte[] b = new byte[21];
            b[0] = 200; // 10 pt
            b[4] = 0xFF; b[5] = 0x7F; // color automático
            b[14] = 5; b[15] = 0;
            b[16] = (byte)'A'; b[17] = (byte)'r'; b[18] = (byte)'i';
            b[19] = (byte)'a'; b[20] = (byte)'l';
            return b;
        }

        private static void ApplyFont(CellStyle st, byte[] font, XlsContext ctx)
        {
            if (font.Length < 16) return;
            int height = U16(font, 0);
            if (height > 0) st.FontSize = Math.Max(6, Math.Min(72, height / 20f));
            int grbit = U16(font, 2);
            st.Bold = (grbit & 1) != 0;
            st.Italic = (grbit & 2) != 0;
            st.Underline = font[10] != 0;
            int color = U16(font, 4);
            if (color >= 0 && color < ctx.Palette.Length)
                st.FontColor = Color.FromArgb(unchecked((int)0xFF000000) | (ctx.Palette[color] & 0xFFFFFF));
            int cch = font[14];
            int flags = font[15];
            int p = 16;
            string name = ReadChars(font, ref p, cch, (flags & 1) != 0);
            if (!string.IsNullOrEmpty(name)) st.FontName = name;
        }

        private static int[] DefaultPalette()
        {
            return new int[]
            {
                0x000000,0xFFFFFF,0xFF0000,0x00FF00,0x0000FF,0xFFFF00,0xFF00FF,0x00FFFF,
                0x000000,0xFFFFFF,0xFF0000,0x00FF00,0x0000FF,0xFFFF00,0xFF00FF,0x00FFFF,
                0x800000,0x008000,0x000080,0x808000,0x800080,0x008080,0xC0C0C0,0x808080,
                0x9999FF,0x993366,0xFFFFCC,0xCCFFFF,0x660066,0xFF8080,0x0066CC,0xCCCCFF,
                0x000080,0xFF00FF,0xFFFF00,0x00FFFF,0x800080,0x800000,0x008080,0x0000FF,
                0x00CCFF,0xCCFFFF,0xCCFFCC,0xFFFF99,0x99CCFF,0xFF99CC,0xCC99FF,0xFFCC99,
                0x3366FF,0x33CCCC,0x99CC00,0xFFCC00,0xFF9900,0xFF6600,0x666699,0x969696,
                0x003366,0x339966,0x003300,0x333300,0x993300,0x993366,0x333399,0x333333
            };
        }

        private static string MakeUniqueName(Workbook wb, string name)
        {
            if (string.IsNullOrEmpty(name)) name = "Hoja";
            string baseName = name;
            int n = 2;
            while (wb.FindSheet(name) != null) name = baseName + n++;
            return name;
        }

        private static int U16(byte[] b, int p)
        {
            if (p < 0 || p + 2 > b.Length) return 0;
            return b[p] | (b[p + 1] << 8);
        }

        private static uint U32(byte[] b, int p)
        {
            if (p < 0 || p + 4 > b.Length) return 0;
            return (uint)(b[p] | (b[p + 1] << 8) | (b[p + 2] << 16) | (b[p + 3] << 24));
        }
    }
}
