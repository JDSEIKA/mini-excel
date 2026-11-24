using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;

namespace CalcRapido
{
    public sealed class FileFormatException : Exception
    {
        public FileFormatException(string message) : base(message) { }
    }

    public static class FileIO
    {
        private const string Magic = "CRB1";
        private const int Version = 7;
        private const int MaxCellsPerSheet = 5000000;
        private const int MaxCellsTotal = 4000000;

        public static Workbook Open(string path)
        {
            string ext = Path.GetExtension(path).ToLowerInvariant();
            Workbook wb;
            if (ext == ".csv" || ext == ".txt" || ext == ".tsv") wb = LoadCsv(path);
            else if (ext == ".xlsx") wb = XlsxIO.Load(path);
            else if (ext == ".xls") wb = XlsIO.Load(path);
            else if (ext == ".crb") wb = LoadNative(path);
            else throw new FileFormatException("Extensión no soportada. Use .xlsx, .xls, .csv, .tsv, .txt o .crb.");
            CompatibilityScanner.ScanFormulas(wb);
            return wb;
        }

        public static void Save(Workbook wb, string path)
        {
            string ext = Path.GetExtension(path).ToLowerInvariant();
            if (ext == ".csv") { SaveCsv(wb, path); return; }
            if (ext == ".xlsx") { XlsxIO.Save(wb, path); return; }
            if (ext == ".xls")
                throw new FileFormatException("El formato .xls es de solo lectura en esta versión. Use Guardar como .xlsx.");
            if (ext == ".crb") { SaveNative(wb, path); return; }
            throw new FileFormatException("Extensión de guardado no soportada. Use .xlsx, .csv o .crb.");
        }

        public static void SaveNative(Workbook wb, string path)
        {
            string dir = Path.GetDirectoryName(Path.GetFullPath(path));
            if (!Directory.Exists(dir)) throw new DirectoryNotFoundException("No existe la carpeta de destino.");
            string temp = Path.Combine(dir, "~calcrapido_" + Guid.NewGuid().ToString("N") + ".tmp");
            try
            {
                using (FileStream fs = new FileStream(temp, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                using (BinaryWriter w = new BinaryWriter(fs, new UTF8Encoding(false)))
                {
                    w.Write(Encoding.ASCII.GetBytes(Magic));
                    w.Write(Version);
                    w.Write(wb.Sheets.Count);
                    w.Write(wb.ActiveSheetIndex);
                    w.Write(wb.Date1904);
                    w.Write(wb.DefinedNames.Count);
                    foreach (DefinedName dn in wb.DefinedNames)
                    {
                        w.Write(dn.Name ?? "");
                        w.Write(dn.Formula ?? "");
                        w.Write(dn.LocalSheetId);
                        w.Write(dn.Hidden);
                        w.Write(dn.Comment ?? "");
                    }
                    foreach (Sheet sheet in wb.Sheets)
                    {
                        w.Write(sheet.Name ?? "Hoja");
                        w.Write(sheet.ColWidths.Count);
                        foreach (KeyValuePair<int, float> kv in sheet.ColWidths)
                        {
                            w.Write(kv.Key);
                            w.Write(kv.Value);
                        }
                        w.Write(sheet.RowHeights.Count);
                        foreach (KeyValuePair<int, float> kv in sheet.RowHeights)
                        {
                            w.Write(kv.Key);
                            w.Write(kv.Value);
                        }
                        w.Write(sheet.HiddenRows.Count);
                        foreach (int row in sheet.HiddenRows) w.Write(row);
                        w.Write(sheet.HiddenCols.Count);
                        foreach (int col in sheet.HiddenCols) w.Write(col);
                        w.Write(sheet.Merges.Count);
                        foreach (MergeRange m in sheet.Merges)
                        {
                            w.Write(m.Row1); w.Write(m.Col1); w.Write(m.Row2); w.Write(m.Col2);
                        }
                        w.Write(sheet.FreezeRows);
                        w.Write(sheet.FreezeCols);
                        List<long> keys = new List<long>(sheet.Cells.Keys);
                        keys.Sort();
                        w.Write(keys.Count);
                        foreach (long key in keys)
                        {
                            Cell cell = sheet.Cells[key];
                            w.Write((int)(key >> 14));
                            w.Write((int)(key & 0x3FFF));
                            byte kind = cell.IsFormula ? (byte)255 : (byte)cell.Value.Kind;
                            w.Write(kind);
                            w.Write(cell.Raw != null);
                            if (cell.Raw != null) w.Write(cell.Raw);
                            if (kind == (byte)ValueKind.Number) w.Write(cell.Value.Number);
                            else if (kind == (byte)ValueKind.Bool) w.Write(cell.Value.Bool);
                            else if (kind == (byte)ValueKind.Error) w.Write(cell.Value.Error ?? "");
                            bool hasStyle = cell.Style != null && !cell.Style.IsDefault();
                            w.Write(hasStyle);
                            if (hasStyle) WriteStyle(w, cell.Style);
                            bool hasLink = !string.IsNullOrEmpty(cell.Hyperlink);
                            w.Write(hasLink);
                            if (hasLink)
                            {
                                w.Write(cell.Hyperlink);
                                w.Write(cell.HyperlinkInternal);
                                w.Write(cell.HyperlinkTooltip ?? "");
                            }
                        }
                        w.Write(sheet.Charts.Count);
                        foreach (SheetChart chart in sheet.Charts)
                        {
                            w.Write((byte)chart.Kind);
                            w.Write(chart.Title ?? "");
                            w.Write(chart.ShowLegend);
                            w.Write(chart.ShowAxes);
                            w.Write(chart.Stacked);
                            w.Write(chart.FromRow);
                            w.Write(chart.FromCol);
                            w.Write(chart.ToRow);
                            w.Write(chart.ToCol);
                            w.Write(chart.FromColOff);
                            w.Write(chart.FromRowOff);
                            w.Write(chart.ToColOff);
                            w.Write(chart.ToRowOff);
                            w.Write(chart.WidthPx);
                            w.Write(chart.HeightPx);
                            w.Write(chart.Series.Count);
                            foreach (ChartSeries ser in chart.Series)
                            {
                                w.Write(ser.Name ?? "");
                                w.Write(ser.NameFormula ?? "");
                                w.Write(ser.CategoriesFormula ?? "");
                                w.Write(ser.ValuesFormula ?? "");
                                w.Write(ser.XValuesFormula ?? "");
                                w.Write(ser.Color.IsEmpty ? 0 : ser.Color.ToArgb());
                            }
                        }
                        w.Write(sheet.Validations.Count);
                        foreach (DataValidationRule rule in sheet.Validations)
                        {
                            w.Write(rule.Row1); w.Write(rule.Col1);
                            w.Write(rule.Row2); w.Write(rule.Col2);
                            w.Write(rule.SourceRange ?? "");
                            w.Write(rule.Message ?? "");
                            w.Write(rule.AllowBlank);
                            w.Write(rule.Items.Count);
                            foreach (string item in rule.Items) w.Write(item ?? "");
                        }
                        w.Write(sheet.Images.Count);
                        foreach (SheetImage img in sheet.Images)
                        {
                            w.Write(img.FromRow); w.Write(img.FromCol);
                            w.Write(img.ToRow); w.Write(img.ToCol);
                            w.Write(img.FromColOff); w.Write(img.FromRowOff);
                            w.Write(img.ToColOff); w.Write(img.ToRowOff);
                            w.Write(img.WidthPx); w.Write(img.HeightPx);
                            byte[] bytes = new byte[0];
                            if (img.Image != null)
                            {
                                using (MemoryStream ms = new MemoryStream())
                                {
                                    img.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                                    bytes = ms.ToArray();
                                }
                            }
                            w.Write(bytes.Length);
                            w.Write(bytes);
                        }
                        w.Write(sheet.AutoFilterRef ?? "");
                    }
                }
                if (File.Exists(path)) File.Replace(temp, path, null);
                else File.Move(temp, path);
            }
            finally
            {
                if (File.Exists(temp))
                {
                    try { File.Delete(temp); } catch { }
                }
            }
            wb.XlsxSource = null;
        }

        private static void WriteStyle(BinaryWriter w, CellStyle s)
        {
            int flags = 0;
            if (s.Bold) flags |= 1;
            if (s.Italic) flags |= 2;
            if (s.Underline) flags |= 4;
            if (s.WrapText) flags |= 8;
            w.Write((byte)flags);
            w.Write((sbyte)s.Align);
            w.Write((byte)s.Format);
            w.Write((byte)Math.Max(0, Math.Min(10, s.Decimals)));
            w.Write(s.FontSize);
            w.Write(s.FontColor.ToArgb());
            w.Write(s.FillColor.IsEmpty ? 0 : s.FillColor.ToArgb());
            w.Write((byte)s.Borders);
            w.Write(s.FontName ?? "Segoe UI");
            w.Write(s.CustomFormat ?? "");
        }

        private static string ReadBoundedString(BinaryReader r, int maxBytes)
        {
            int length = 0;
            int shift = 0;
            while (true)
            {
                if (shift > 28) throw new FileFormatException("Cadena no válida en el archivo .crb.");
                byte b = r.ReadByte();
                length |= (b & 0x7F) << shift;
                if ((b & 0x80) == 0) break;
                shift += 7;
            }
            if (length < 0 || length > maxBytes)
                throw new FileFormatException("Cadena demasiado grande en el archivo .crb.");
            byte[] bytes = r.ReadBytes(length);
            if (bytes.Length != length) throw new EndOfStreamException();
            return Encoding.UTF8.GetString(bytes);
        }

        private static CellStyle ReadStyle(BinaryReader r, int version)
        {
            CellStyle s = new CellStyle();
            int flags = r.ReadByte();
            s.Bold = (flags & 1) != 0;
            s.Italic = (flags & 2) != 0;
            s.Underline = (flags & 4) != 0;
            s.WrapText = (flags & 8) != 0;
            s.Align = r.ReadSByte();
            if (s.Align < -1 || s.Align > 2) s.Align = -1;
            s.Format = (NumberFormatKind)r.ReadByte();
            if ((int)s.Format < 0 || (int)s.Format > 7) s.Format = NumberFormatKind.General;
            s.Decimals = r.ReadByte();
            s.FontSize = r.ReadSingle();
            if (s.FontSize < 6 || s.FontSize > 72) s.FontSize = 10;
            s.FontColor = Color.FromArgb(r.ReadInt32());
            int fill = r.ReadInt32();
            s.FillColor = fill == 0 ? Color.Empty : Color.FromArgb(fill);
            s.Borders = (BorderFlags)r.ReadByte();
            s.FontName = ReadBoundedString(r, 1024);
            if (string.IsNullOrEmpty(s.FontName)) s.FontName = "Segoe UI";
            if (version >= 2) s.CustomFormat = ReadBoundedString(r, 8192);
            return s;
        }

        public static Workbook LoadNative(string path)
        {
            FileInfo fi = new FileInfo(path);
            if (fi.Length > 120L * 1024 * 1024) throw new FileFormatException("El archivo .crb es demasiado grande.");
            try
            {
                using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (BinaryReader r = new BinaryReader(fs, new UTF8Encoding(false)))
                {
                    byte[] magic = r.ReadBytes(4);
                    if (magic.Length != 4 || Encoding.ASCII.GetString(magic) != Magic)
                        throw new FileFormatException("El archivo no es un libro de Mini Excel.");
                    int version = r.ReadInt32();
                    if (version < 1 || version > Version)
                        throw new FileFormatException("Versión de archivo no compatible.");
                    int sheetCount = r.ReadInt32();
                    if (sheetCount < 1 || sheetCount > 1000)
                        throw new FileFormatException("Número de hojas no válido.");
                    int active = r.ReadInt32();
                    Workbook wb = new Workbook();
                    wb.Sheets.Clear();
                    long totalCells = 0;
                    if (version >= 3)
                    {
                        wb.Date1904 = r.ReadBoolean();
                        int names = r.ReadInt32();
                        if (names < 0 || names > 100000) throw new FileFormatException("Nombres definidos no válidos.");
                        for (int i = 0; i < names; i++)
                        {
                            DefinedName d = new DefinedName();
                            d.Name = ReadBoundedString(r, 255);
                            d.Formula = ReadBoundedString(r, 8192);
                            d.LocalSheetId = r.ReadInt32();
                            d.Hidden = r.ReadBoolean();
                            d.Comment = ReadBoundedString(r, 4096);
                            wb.DefinedNames.Add(d);
                        }
                    }
                    for (int si = 0; si < sheetCount; si++)
                    {
                        string name = ReadBoundedString(r, 255);
                        if (string.IsNullOrEmpty(name)) name = "Hoja" + (si + 1);
                        Sheet sheet = new Sheet(wb, name);
                        int widthCount = r.ReadInt32();
                        if (widthCount < 0 || widthCount > Limits.MaxCols)
                            throw new FileFormatException("Ancho de columnas no válido.");
                        for (int i = 0; i < widthCount; i++)
                        {
                            int col = r.ReadInt32();
                            float width = r.ReadSingle();
                            if (col >= 0 && col < Limits.MaxCols && width > 0)
                                sheet.ColWidths[col] = Math.Max(28f, Math.Min(800f, width));
                        }
                        if (version >= 3)
                        {
                            int rowHeights = r.ReadInt32();
                            if (rowHeights < 0 || rowHeights > Limits.MaxRows) throw new FileFormatException("Altos de fila no válidos.");
                            for (int i = 0; i < rowHeights; i++)
                            {
                                int row = r.ReadInt32();
                                float height = r.ReadSingle();
                                if (row >= 0 && row < Limits.MaxRows && height >= 0)
                                    sheet.SetRowHeightPointsImported(row, height);
                            }
                            int hiddenRows = r.ReadInt32();
                            if (hiddenRows < 0 || hiddenRows > Limits.MaxRows) throw new FileFormatException("Filas ocultas no válidas.");
                            for (int i = 0; i < hiddenRows; i++)
                            {
                                int row = r.ReadInt32();
                                if (row >= 0 && row < Limits.MaxRows) sheet.SetRowHiddenImported(row, true);
                            }
                            int hiddenCols = r.ReadInt32();
                            if (hiddenCols < 0 || hiddenCols > Limits.MaxCols) throw new FileFormatException("Columnas ocultas no válidas.");
                            for (int i = 0; i < hiddenCols; i++)
                            {
                                int col = r.ReadInt32();
                                if (col >= 0 && col < Limits.MaxCols) sheet.SetColumnHiddenImported(col, true);
                            }
                            int merges = r.ReadInt32();
                            if (merges < 0 || merges > 100000) throw new FileFormatException("Combinaciones no válidas.");
                            for (int i = 0; i < merges; i++)
                            {
                                int r1 = r.ReadInt32(), c1 = r.ReadInt32(), r2 = r.ReadInt32(), c2 = r.ReadInt32();
                                if (r1 >= 0 && c1 >= 0 && r2 < Limits.MaxRows && c2 < Limits.MaxCols)
                                    sheet.Merges.Add(new MergeRange(r1, c1, r2, c2));
                            }
                            sheet.RebuildMergeLookup();
                            sheet.FreezeRows = Math.Max(0, Math.Min(Limits.MaxRows - 1, r.ReadInt32()));
                            sheet.FreezeCols = Math.Max(0, Math.Min(Limits.MaxCols - 1, r.ReadInt32()));
                        }
                        int cellCount = r.ReadInt32();
                        if (cellCount < 0 || cellCount > MaxCellsPerSheet)
                            throw new FileFormatException("El archivo contiene demasiadas celdas.");
                        for (int i = 0; i < cellCount; i++)
                        {
                            int row = r.ReadInt32();
                            int col = r.ReadInt32();
                            if (row < 0 || row >= Limits.MaxRows || col < 0 || col >= Limits.MaxCols)
                                throw new FileFormatException("Referencia de celda fuera de rango.");
                            byte kind = version >= 3 ? r.ReadByte() : (byte)255;
                            string raw = r.ReadBoolean() ? ReadBoundedString(r, 1048576) : null;
                            double number = 0;
                            bool boolValue = false;
                            string errorValue = null;
                            if (version >= 3)
                            {
                                if (kind == (byte)ValueKind.Number) number = r.ReadDouble();
                                else if (kind == (byte)ValueKind.Bool) boolValue = r.ReadBoolean();
                                else if (kind == (byte)ValueKind.Error) errorValue = ReadBoundedString(r, 256);
                            }
                            CellStyle style = r.ReadBoolean() ? ReadStyle(r, version) : null;
                            bool hasLink = version >= 3 && r.ReadBoolean();
                            string link = hasLink ? ReadBoundedString(r, 4096) : null;
                            bool linkInternal = hasLink && r.ReadBoolean();
                            string linkTooltip = hasLink ? ReadBoundedString(r, 4096) : null;
                            Cell cell = new Cell();
                            if (version >= 3)
                            {
                                if (kind == 255) cell.SetInput(raw);
                                else if (kind == (byte)ValueKind.Text) cell.SetTextLiteral(raw);
                                else if (kind == (byte)ValueKind.Number) cell.SetNumber(number);
                                else if (kind == (byte)ValueKind.Bool) cell.SetInput(boolValue ? "VERDADERO" : "FALSO");
                                else if (kind == (byte)ValueKind.Error) cell.SetError(errorValue);
                                else { cell.Raw = raw; cell.Value = CellValue.Empty; }
                            }
                            else if (raw != null) cell.SetInput(raw);
                            cell.Style = style;
                            cell.Hyperlink = link;
                            cell.HyperlinkInternal = linkInternal;
                            cell.HyperlinkTooltip = linkTooltip;
                            sheet.Cells[Sheet.Key(row, col)] = cell;
                        }
                        if (version >= 4)
                        {
                            int chartCount = r.ReadInt32();
                            if (chartCount < 0 || chartCount > 10000)
                                throw new FileFormatException("Gráficos no válidos en el archivo .crb.");
                            for (int ci = 0; ci < chartCount; ci++)
                            {
                                SheetChart chart = new SheetChart();
                                chart.Kind = (ChartKind)r.ReadByte();
                                if ((int)chart.Kind < 0 || (int)chart.Kind > (int)ChartKind.Unknown)
                                    chart.Kind = ChartKind.Column;
                                chart.Title = ReadBoundedString(r, 4096);
                                chart.ShowLegend = r.ReadBoolean();
                                chart.ShowAxes = r.ReadBoolean();
                                chart.Stacked = r.ReadBoolean();
                                chart.FromRow = r.ReadInt32();
                                chart.FromCol = r.ReadInt32();
                                chart.ToRow = r.ReadInt32();
                                chart.ToCol = r.ReadInt32();
                                chart.FromColOff = r.ReadInt64();
                                chart.FromRowOff = r.ReadInt64();
                                chart.ToColOff = r.ReadInt64();
                                chart.ToRowOff = r.ReadInt64();
                                chart.WidthPx = Math.Max(40, Math.Min(4000, r.ReadInt32()));
                                chart.HeightPx = Math.Max(30, Math.Min(4000, r.ReadInt32()));
                                int seriesCount = r.ReadInt32();
                                if (seriesCount < 0 || seriesCount > 1000)
                                    throw new FileFormatException("Series de gráfico no válidas.");
                                for (int sidx = 0; sidx < seriesCount; sidx++)
                                {
                                    ChartSeries ser = new ChartSeries();
                                    ser.Name = ReadBoundedString(r, 4096);
                                    ser.NameFormula = ReadBoundedString(r, 8192);
                                    ser.CategoriesFormula = ReadBoundedString(r, 8192);
                                    ser.ValuesFormula = ReadBoundedString(r, 8192);
                                    ser.XValuesFormula = ReadBoundedString(r, 8192);
                                    int color = r.ReadInt32();
                                    ser.Color = color == 0 ? Color.Empty : Color.FromArgb(color);
                                    chart.Series.Add(ser);
                                }
                                sheet.Charts.Add(chart);
                            }
                        }
                        if (version >= 5)
                        {
                            int validationCount = r.ReadInt32();
                            if (validationCount < 0 || validationCount > 10000)
                                throw new FileFormatException("Validaciones no válidas en el archivo .crb.");
                            for (int vi = 0; vi < validationCount; vi++)
                            {
                                DataValidationRule rule = new DataValidationRule();
                                rule.Row1 = r.ReadInt32(); rule.Col1 = r.ReadInt32();
                                rule.Row2 = r.ReadInt32(); rule.Col2 = r.ReadInt32();
                                rule.SourceRange = ReadBoundedString(r, 4096);
                                rule.Message = ReadBoundedString(r, 4096);
                                rule.AllowBlank = r.ReadBoolean();
                                int itemCount = r.ReadInt32();
                                if (itemCount < 0 || itemCount > 10000)
                                    throw new FileFormatException("Lista de validación no válida.");
                                for (int ii = 0; ii < itemCount; ii++)
                                    rule.Items.Add(ReadBoundedString(r, 4096));
                                sheet.Validations.Add(rule);
                            }
                        }
                        if (version >= 6)
                        {
                            int imageCount = r.ReadInt32();
                            if (imageCount < 0 || imageCount > 1000)
                                throw new FileFormatException("Imágenes no válidas en el archivo .crb.");
                            for (int im = 0; im < imageCount; im++)
                            {
                                SheetImage image = new SheetImage();
                                image.FromRow = r.ReadInt32(); image.FromCol = r.ReadInt32();
                                image.ToRow = r.ReadInt32(); image.ToCol = r.ReadInt32();
                                image.FromColOff = r.ReadInt64(); image.FromRowOff = r.ReadInt64();
                                image.ToColOff = r.ReadInt64(); image.ToRowOff = r.ReadInt64();
                                image.WidthPx = Math.Max(1, Math.Min(10000, r.ReadInt32()));
                                image.HeightPx = Math.Max(1, Math.Min(10000, r.ReadInt32()));
                                int len = r.ReadInt32();
                                if (len < 0 || len > 20 * 1024 * 1024) throw new FileFormatException("Imagen demasiado grande.");
                                byte[] bytes = r.ReadBytes(len);
                                if (bytes.Length != len) throw new EndOfStreamException();
                                if (len > 0)
                                {
                                    if (!XlsxIO.IsSafeRasterImage(bytes)) continue;
                                    using (MemoryStream ms = new MemoryStream(bytes))
                                    using (Image src = Image.FromStream(ms))
                                    {
                                        if (src.Width <= 0 || src.Height <= 0 ||
                                            src.Width > 12000 || src.Height > 12000 ||
                                            (long)src.Width * src.Height > 25000000) continue;
                                        image.Image = new Bitmap(src);
                                    }
                                }
                                sheet.Images.Add(image);
                            }
                        }
                        if (version >= 7) sheet.AutoFilterRef = ReadBoundedString(r, 4096);
                        totalCells += sheet.Cells.Count;
                        if (totalCells > MaxCellsTotal)
                            throw new FileFormatException("El archivo .crb contiene demasiadas celdas en total.");
                        sheet.RecalculateUsedRange();
                        wb.Sheets.Add(sheet);
                    }
                    wb.ActiveSheetIndex = Math.Max(0, Math.Min(wb.Sheets.Count - 1, active));
                    wb.Recalculate();
                    return wb;
                }
            }
            catch (FileFormatException) { throw; }
            catch (EndOfStreamException) { throw new FileFormatException("El archivo está incompleto o dañado."); }
            catch (IOException ex) { throw new IOException("No se pudo leer el archivo: " + ex.Message, ex); }
            catch (Exception ex) { throw new FileFormatException("No se pudo abrir el libro: " + ex.Message); }
        }

        public static void SaveCsv(Workbook wb, string path)
        {
            Sheet s = wb.ActiveSheet;
            int maxRow = s.UsedMaxRow;
            int maxCol = s.UsedMaxCol;
            if (maxRow < 0) maxRow = 0;
            if (maxCol < 0) maxCol = 0;
            maxRow = Math.Min(maxRow, 200000);
            maxCol = Math.Min(maxCol, 1000);
            if ((long)(maxRow + 1) * (maxCol + 1) > 2000000)
                throw new FileFormatException("La hoja es demasiado grande para exportarla a CSV.");
            char separator = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator == ","
                ? ';' : ',';
            StringBuilder sb = new StringBuilder();
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    if (col > 0) sb.Append(separator);
                    Cell c = s.GetCell(row, col);
                    if (c == null) continue;
                    string text = CsvText(c, separator);
                    if (text.IndexOfAny(new char[] { separator, '"', '\r', '\n' }) >= 0)
                    {
                        sb.Append('"');
                        sb.Append(text.Replace("\"", "\"\""));
                        sb.Append('"');
                    }
                    else sb.Append(text);
                }
                sb.Append("\r\n");
            }
            AtomicIO.Save(path, delegate(string temp)
            {
                File.WriteAllText(temp, sb.ToString(), new UTF8Encoding(true));
            });
        }

        private static string CsvText(Cell cell, char separator)
        {
            if (cell == null) return "";
            CellValue value = cell.Value;
            string text;
            if (value.Kind == ValueKind.Number)
            {
                CellStyle style = cell.EffectiveStyle;
                if (style != null && (style.Format == NumberFormatKind.Date ||
                    style.Format == NumberFormatKind.Time || style.Format == NumberFormatKind.Percent ||
                    style.Format == NumberFormatKind.Currency || style.Format == NumberFormatKind.Custom))
                    text = NumberFormatter.Format(value, style);
                else
                    text = value.Number.ToString("R",
                        separator == ';' ? CultureInfo.CurrentCulture : CultureInfo.InvariantCulture);
            }
            else if (value.Kind == ValueKind.Empty) text = "";
            else text = value.ToDisplayString();

            // Protección contra inyección CSV también para resultados de formulas.
            if (text.Length > 0 && "=+-@\t\r".IndexOf(text[0]) >= 0)
                text = "'" + text;
            return text;
        }

        public static Workbook LoadCsv(string path)
        {
            FileInfo fi = new FileInfo(path);
            if (fi.Length > 50L * 1024 * 1024) throw new FileFormatException("El archivo CSV es demasiado grande.");
            byte[] bytes = File.ReadAllBytes(path);
            bool usedFallback;
            string text = DecodeCsv(bytes, out usedFallback);
            char delimiter = DetectDelimiter(text);
            List<List<string>> rows = ParseDelimited(text, delimiter);
            Workbook wb = new Workbook();
            if (usedFallback)
                wb.AddLoadWarning("CSV sin UTF-8 válido: se interpretó como Windows-1252.");
            Sheet sheet = wb.Sheets[0];
            int maxRows = Math.Min(rows.Count, Limits.MaxRows);
            int cells = 0;
            for (int r = 0; r < maxRows; r++)
            {
                List<string> fields = rows[r];
                int maxCols = Math.Min(fields.Count, 4000);
                for (int c = 0; c < maxCols; c++)
                {
                    string value = fields[c];
                    if (value == null || value.Length == 0) continue;
                    if (value.Length > 1 && value[0] == '\'' && "=+-@\t\r".IndexOf(value[1]) >= 0)
                        value = value.Substring(1);
                    if (++cells > 2000000)
                        throw new FileFormatException("El CSV contiene demasiadas celdas.");
                    // Los teléfonos con + y espacios deben conservarse como texto,
                    // no convertirse en un número que pierde el prefijo.
                    if (value.Length > 1 && value[0] == '+')
                        sheet.SetTextLiteral(r, c, value);
                    else
                        sheet.SetInput(r, c, value);
                }
            }
            sheet.RecalculateUsedRange();
            wb.Recalculate();
            return wb;
        }

        private static string DecodeCsv(byte[] bytes, out bool usedFallback)
        {
            usedFallback = false;
            if (bytes == null || bytes.Length == 0) return "";
            if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
                return Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3);
            if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
                return Encoding.Unicode.GetString(bytes, 2, bytes.Length - 2);
            if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
                return Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2);
            try
            {
                return new UTF8Encoding(false, true).GetString(bytes);
            }
            catch (DecoderFallbackException)
            {
                usedFallback = true;
                try { return Encoding.GetEncoding(1252).GetString(bytes); }
                catch { return Encoding.Default.GetString(bytes); }
            }
        }

        private static char DetectDelimiter(string text)
        {
            int lineEnd = text.IndexOf('\n');
            if (lineEnd < 0) lineEnd = text.Length;
            string first = text.Substring(0, lineEnd);
            int commas = 0, semis = 0, tabs = 0;
            bool quotes = false;
            for (int i = 0; i < first.Length; i++)
            {
                char ch = first[i];
                if (ch == '"') quotes = !quotes;
                else if (!quotes)
                {
                    if (ch == ',') commas++;
                    else if (ch == ';') semis++;
                    else if (ch == '\t') tabs++;
                }
            }
            if (tabs > commas && tabs > semis) return '\t';
            if (semis > commas) return ';';
            return ',';
        }

        private static List<List<string>> ParseDelimited(string text, char delimiter)
        {
            List<List<string>> rows = new List<List<string>>();
            List<string> row = new List<string>();
            StringBuilder field = new StringBuilder();
            bool quoted = false;
            for (int i = 0; i < text.Length; i++)
            {
                char ch = text[i];
                if (quoted)
                {
                    if (ch == '"')
                    {
                        if (i + 1 < text.Length && text[i + 1] == '"') { field.Append('"'); i++; }
                        else quoted = false;
                    }
                    else field.Append(ch);
                }
                else
                {
                    if (ch == '"') quoted = true;
                    else if (ch == delimiter)
                    {
                        row.Add(field.ToString());
                        field.Length = 0;
                    }
                    else if (ch == '\r')
                    {
                        if (i + 1 < text.Length && text[i + 1] == '\n') i++;
                        row.Add(field.ToString());
                        field.Length = 0;
                        rows.Add(row);
                        row = new List<string>();
                    }
                    else if (ch == '\n')
                    {
                        row.Add(field.ToString());
                        field.Length = 0;
                        rows.Add(row);
                        row = new List<string>();
                    }
                    else field.Append(ch);
                }
            }
            if (field.Length > 0 || row.Count > 0)
            {
                row.Add(field.ToString());
                rows.Add(row);
            }
            return rows;
        }
    }
}
