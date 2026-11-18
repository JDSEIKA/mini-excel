using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.Globalization;
using System.Text;
using System.Windows.Forms;

namespace CalcRapido
{
    public enum PrintScope
    {
        Selection,
        ActiveSheet,
        SelectedSheets,
        AllSheets
    }

    public enum PrintFitMode
    {
        Actual,
        FitWidth,
        FitPage
    }

    public sealed class PrintOptions
    {
        public PrintScope Scope = PrintScope.ActiveSheet;
        public List<int> SheetIndices = new List<int>();
        public bool Landscape;
        public PaperKind Paper = PaperKind.A4;
        public float MarginMm = 12f;
        public PrintFitMode Fit = PrintFitMode.Actual;
        public bool PrintGridlines;
        public bool PrintHeadings;
        public bool CenterHorizontally;
        public bool PreviewRequested;
    }

    internal sealed class PrintPageInfo
    {
        public Sheet Sheet;
        public int Row1;
        public int Row2;
        public int Col1;
        public int Col2;
        public float Scale;
        public int PageNumber;
        public int PageCount;
        public int TitleRow1 = -1;
        public int TitleRow2 = -1;
        public int TitleCol1 = -1;
        public int TitleCol2 = -1;
    }

    public sealed class WorkbookPrintDocument : PrintDocument
    {
        private readonly Workbook _workbook;
        private readonly PrintOptions _options;
        private readonly int _activeSheet;
        private readonly int _selTop, _selLeft, _selBottom, _selRight;
        private readonly List<PrintPageInfo> _pages = new List<PrintPageInfo>();
        private int _nextPage;

        public int PreparedPageCount { get { return _pages.Count; } }

        public void PrepareLayout()
        {
            OnBeginPrint(new PrintEventArgs());
        }

        public WorkbookPrintDocument(Workbook workbook, PrintOptions options, int activeSheet,
                                     int selTop, int selLeft, int selBottom, int selRight)
        {
            _workbook = workbook;
            _options = options;
            _activeSheet = activeSheet;
            _selTop = selTop; _selLeft = selLeft; _selBottom = selBottom; _selRight = selRight;
            DocumentName = "Mini Excel";
            DefaultPageSettings.Landscape = options.Landscape;
            DefaultPageSettings.Margins = new Margins(
                MmToHundredths(options.MarginMm), MmToHundredths(options.MarginMm),
                MmToHundredths(options.MarginMm), MmToHundredths(options.MarginMm));
            DefaultPageSettings.PaperSize = PaperSizeFor(options.Paper);
        }

        private static int MmToHundredths(float mm)
        {
            return Math.Max(0, (int)Math.Round(mm / 25.4 * 100.0));
        }

        private static PaperSize PaperSizeFor(PaperKind kind)
        {
            if (kind == PaperKind.Letter) return new PaperSize("Carta", 850, 1100);
            if (kind == PaperKind.Legal) return new PaperSize("Legal", 850, 1400);
            return new PaperSize("A4", 827, 1169);
        }

        protected override void OnBeginPrint(PrintEventArgs e)
        {
            base.OnBeginPrint(e);
            _pages.Clear();
            _nextPage = 0;
            List<int> indices = new List<int>();
            if (_options.Scope == PrintScope.Selection || _options.Scope == PrintScope.ActiveSheet)
                indices.Add(_activeSheet);
            else if (_options.Scope == PrintScope.SelectedSheets)
                indices.AddRange(_options.SheetIndices);
            else
                for (int i = 0; i < _workbook.Sheets.Count; i++) indices.Add(i);

            foreach (int index in indices)
            {
                if (index < 0 || index >= _workbook.Sheets.Count) continue;
                Sheet sheet = _workbook.Sheets[index];
                int r1 = 0, c1 = 0, r2 = Math.Max(0, sheet.UsedMaxRow), c2 = Math.Max(0, sheet.UsedMaxCol);
                if (_options.Scope == PrintScope.Selection && index == _activeSheet)
                {
                    r1 = _selTop; c1 = _selLeft; r2 = _selBottom; c2 = _selRight;
                }
                else
                {
                    int ar1, ac1, ar2, ac2;
                    if (TryParsePrintArea(sheet, sheet.PrintArea, out ar1, out ac1, out ar2, out ac2))
                    {
                        r1 = ar1; c1 = ac1; r2 = ar2; c2 = ac2;
                    }
                }
                AddSheetPages(sheet, r1, c1, r2, c2);
            }
            if (_pages.Count == 0)
                _pages.Add(new PrintPageInfo
                {
                    Sheet = _workbook.ActiveSheet, Row1 = 0, Row2 = 0, Col1 = 0, Col2 = 0,
                    Scale = 1f, PageNumber = 1, PageCount = 1
                });
            for (int i = 0; i < _pages.Count; i++)
            {
                _pages[i].PageNumber = i + 1;
                _pages[i].PageCount = _pages.Count;
            }
        }

        private static bool TryParsePrintArea(Sheet sheet, string area,
                                              out int r1, out int c1, out int r2, out int c2)
        {
            r1 = c1 = r2 = c2 = 0;
            if (string.IsNullOrEmpty(area)) return false;
            string text = area;
            int comma = text.IndexOf(',');
            if (comma >= 0) text = text.Substring(0, comma);
            int bang = text.LastIndexOf('!');
            if (bang >= 0)
            {
                string sheetName = text.Substring(0, bang).Trim();
                if (sheetName.Length >= 2 && sheetName[0] == '\'' && sheetName[sheetName.Length - 1] == '\'')
                    sheetName = sheetName.Substring(1, sheetName.Length - 2).Replace("''", "'");
                if (!string.Equals(sheetName, sheet.Name, StringComparison.OrdinalIgnoreCase)) return false;
                text = text.Substring(bang + 1);
            }
            return AddressUtil.TryParseRange(text.Trim().Replace("$", ""), out r1, out c1, out r2, out c2);
        }

        private static bool TryParsePrintTitles(string text, out int row1, out int row2)
        {
            row1 = row2 = -1;
            if (string.IsNullOrEmpty(text)) return false;
            foreach (string raw in text.Split(','))
            {
                string value = raw;
                int bang = value.LastIndexOf('!');
                if (bang >= 0) value = value.Substring(bang + 1);
                string[] parts = value.Trim().Split(':');
                if (parts.Length != 2) continue;
                string a = parts[0].Trim().Replace("$", "");
                string b = parts[1].Trim().Replace("$", "");
                int n1, n2;
                if (!int.TryParse(a, out n1) || !int.TryParse(b, out n2)) continue;
                row1 = Math.Min(n1, n2) - 1;
                row2 = Math.Max(n1, n2) - 1;
                return row1 >= 0;
            }
            return false;
        }

        private static bool TryParsePrintTitleCols(string text, out int col1, out int col2)
        {
            col1 = col2 = -1;
            if (string.IsNullOrEmpty(text)) return false;
            foreach (string raw in text.Split(','))
            {
                string value = raw;
                int bang = value.LastIndexOf('!');
                if (bang >= 0) value = value.Substring(bang + 1);
                string[] parts = value.Trim().Split(':');
                if (parts.Length != 2) continue;
                string a = parts[0].Trim().Replace("$", "");
                string b = parts[1].Trim().Replace("$", "");
                int n1, n2;
                if (!TryColumnIndex(a, out n1) || !TryColumnIndex(b, out n2)) continue;
                col1 = Math.Min(n1, n2);
                col2 = Math.Max(n1, n2);
                return col1 >= 0;
            }
            return false;
        }

        private static bool TryColumnIndex(string value, out int index)
        {
            index = -1;
            if (string.IsNullOrEmpty(value)) return false;
            int n = 0;
            foreach (char c in value.ToUpperInvariant())
            {
                if (c < 'A' || c > 'Z') return false;
                n = n * 26 + (c - 'A' + 1);
            }
            index = n - 1;
            return index >= 0;
        }

        private void AddSheetPages(Sheet sheet, int r1, int c1, int r2, int c2)
        {
            r1 = Math.Max(0, Math.Min(Limits.MaxRows - 1, r1));
            r2 = Math.Max(r1, Math.Min(Limits.MaxRows - 1, r2));
            c1 = Math.Max(0, Math.Min(Limits.MaxCols - 1, c1));
            c2 = Math.Max(c1, Math.Min(Limits.MaxCols - 1, c2));

            int titleRow1, titleRow2;
            bool hasTitles = TryParsePrintTitles(sheet.PrintTitleRows, out titleRow1, out titleRow2);
            if (hasTitles)
            {
                titleRow1 = Math.Max(0, Math.Min(Limits.MaxRows - 1, titleRow1));
                titleRow2 = Math.Max(titleRow1, Math.Min(Limits.MaxRows - 1, titleRow2));
                if (titleRow2 >= r1 && titleRow1 <= r2)
                    r1 = Math.Max(r1, titleRow2 + 1);
            }
            int titleCol1, titleCol2;
            bool hasTitleCols = TryParsePrintTitleCols(sheet.PrintTitleCols, out titleCol1, out titleCol2);
            if (hasTitleCols)
            {
                titleCol1 = Math.Max(0, Math.Min(Limits.MaxCols - 1, titleCol1));
                titleCol2 = Math.Max(titleCol1, Math.Min(Limits.MaxCols - 1, titleCol2));
                if (titleCol2 >= c1 && titleCol1 <= c2)
                    c1 = Math.Max(c1, titleCol2 + 1);
            }

            float paperW = DefaultPageSettings.Landscape
                ? DefaultPageSettings.PaperSize.Height : DefaultPageSettings.PaperSize.Width;
            float paperH = DefaultPageSettings.Landscape
                ? DefaultPageSettings.PaperSize.Width : DefaultPageSettings.PaperSize.Height;
            // PaperSize y Margins están en centésimas de pulgada; la tabla se
            // mide en puntos. Convertir antes de comparar es imprescindible:
            // sin esto una selección grande se sale de la página.
            float usableW = (paperW -
                DefaultPageSettings.Margins.Left - DefaultPageSettings.Margins.Right) * 72f / 100f;
            float usableH = (paperH -
                DefaultPageSettings.Margins.Top - DefaultPageSettings.Margins.Bottom) * 72f / 100f;
            if (usableW < 100) usableW = 100;
            if (usableH < 100) usableH = 100;
            // Reserva para encabezado, pie y nombre de la hoja.
            usableH = Math.Max(80, usableH - 24f);

            float totalW = 0, totalH = 0;
            for (int c = c1; c <= c2; c++)
                if (!sheet.HiddenCols.Contains(c)) totalW += ColPoints(sheet, c);
            for (int r = r1; r <= r2; r++)
                if (!sheet.HiddenRows.Contains(r)) totalH += RowPoints(sheet, r);

            float reserveW = _options.PrintHeadings ? 28f : 0f;
            float reserveH = _options.PrintHeadings ? 20f : 0f;
            float scale = sheet.PrintScalePercent > 0 ? sheet.PrintScalePercent / 100f : 1f;
            if (_options.Fit == PrintFitMode.FitPage)
                scale = Math.Min(1f, Math.Min(
                    Math.Max(1f, usableW - reserveW) / Math.Max(1f, totalW),
                    Math.Max(1f, usableH - reserveH) / Math.Max(1f, totalH)));
            else if (_options.Fit == PrintFitMode.FitWidth)
                scale = Math.Min(1f, Math.Max(1f, usableW - reserveW) / Math.Max(1f, totalW));
            scale = Math.Max(0.08f, scale);

            float headerW = _options.PrintHeadings ? 24f * scale : 0f;
            float headerH = _options.PrintHeadings ? 16f * scale : 0f;
            float titleW = 0;
            if (hasTitleCols)
                for (int tc = titleCol1; tc <= titleCol2 && tc <= c2; tc++)
                    if (!sheet.HiddenCols.Contains(tc)) titleW += ColPoints(sheet, tc) * scale;
            float pageW = Math.Max(60f, usableW - headerW - titleW);
            float titleH = 0;
            if (hasTitles)
                for (int tr = titleRow1; tr <= titleRow2 && tr <= r2; tr++)
                    if (!sheet.HiddenRows.Contains(tr)) titleH += RowPoints(sheet, tr) * scale;
            float pageH = Math.Max(60f, usableH - headerH - titleH);

            // Si se ajusta a una página, no se parte la tabla.
            if (_options.Fit == PrintFitMode.FitPage)
            {
                _pages.Add(new PrintPageInfo
                {
                    Sheet = sheet, Row1 = r1, Row2 = r2, Col1 = c1, Col2 = c2, Scale = scale,
                    TitleRow1 = hasTitles ? titleRow1 : -1,
                    TitleRow2 = hasTitles ? titleRow2 : -1,
                    TitleCol1 = hasTitleCols ? titleCol1 : -1,
                    TitleCol2 = hasTitleCols ? titleCol2 : -1
                });
                return;
            }

            int col = c1;
            while (col <= c2)
            {
                int startCol = col;
                float width = 0;
                while (col <= c2)
                {
                    if (sheet.HiddenCols.Contains(col)) { col++; continue; }
                    float w = ColPoints(sheet, col) * scale;
                    if (col > startCol && width + w > pageW) break;
                    width += w;
                    col++;
                }
                if (col == startCol) col++;
                int endCol = Math.Max(startCol, col - 1);

                int row = r1;
                while (row <= r2)
                {
                    int startRow = row;
                    float height = 0;
                    while (row <= r2)
                    {
                        if (sheet.HiddenRows.Contains(row)) { row++; continue; }
                        float h = RowPoints(sheet, row) * scale;
                        if (row > startRow && height + h > pageH) break;
                        height += h;
                        row++;
                        if (sheet.PrintRowBreaks.Contains(row - 1)) break;
                    }
                    if (row == startRow) row++;
                    int endRow = Math.Max(startRow, row - 1);
                    _pages.Add(new PrintPageInfo
                    {
                        Sheet = sheet, Row1 = startRow, Row2 = endRow,
                        Col1 = startCol, Col2 = endCol, Scale = scale,
                        TitleRow1 = hasTitles ? titleRow1 : -1,
                        TitleRow2 = hasTitles ? titleRow2 : -1,
                        TitleCol1 = hasTitleCols ? titleCol1 : -1,
                        TitleCol2 = hasTitleCols ? titleCol2 : -1
                    });
                }
            }
        }

        private static float ColPoints(Sheet sheet, int col)
        {
            return Math.Max(8f, sheet.GetColWidth(col) * 72f / 96f);
        }

        private static float RowPoints(Sheet sheet, int row)
        {
            return Math.Max(6f, sheet.GetRowHeightPoints(row));
        }

        protected override void OnPrintPage(PrintPageEventArgs e)
        {
            base.OnPrintPage(e);
            if (_nextPage < 0 || _nextPage >= _pages.Count)
            {
                e.HasMorePages = false;
                return;
            }
            PrintPageInfo page = _pages[_nextPage++];
            DrawPage(e.Graphics, e, page);
            e.HasMorePages = _nextPage < _pages.Count;
        }

        private void DrawPage(Graphics g, PrintPageEventArgs e, PrintPageInfo page)
        {
            GraphicsUnit oldUnit = g.PageUnit;
            g.PageUnit = GraphicsUnit.Point;
            try
            {
                RectangleF area = new RectangleF(
                    e.MarginBounds.Left * 72f / 100f,
                    e.MarginBounds.Top * 72f / 100f,
                    e.MarginBounds.Width * 72f / 100f,
                    e.MarginBounds.Height * 72f / 100f);
                float headerW = _options.PrintHeadings ? 24f * page.Scale : 0f;
                float headerH = _options.PrintHeadings ? 16f * page.Scale : 0f;
                float x0 = area.Left + headerW;
                float titleBand = 16f;
                float y0 = area.Top + titleBand + headerH;
                float scale = page.Scale;

                using (Font titleFont = new Font("Segoe UI", 8f, FontStyle.Bold))
                using (Font normalFont = new Font("Segoe UI", 7f))
                using (SolidBrush textBrush = new SolidBrush(Color.Black))
                using (SolidBrush grayBrush = new SolidBrush(Color.Gray))
                {
                    string header = page.Sheet.PrintHeader;
                    if (string.IsNullOrEmpty(header))
                        g.DrawString(page.Sheet.Name, titleFont, textBrush, new PointF(area.Left, area.Top));
                    else
                        DrawHeaderFooter(g, header, page, titleFont, textBrush,
                            area.Left, area.Right, area.Top);
                    if (string.IsNullOrEmpty(page.Sheet.PrintFooter))
                    {
                        string pageText = "Página " + page.PageNumber + " de " + page.PageCount;
                        SizeF ps = g.MeasureString(pageText, normalFont);
                        g.DrawString(pageText, normalFont, grayBrush,
                            new PointF(area.Right - ps.Width, area.Bottom - ps.Height));
                    }
                    else
                    {
                        DrawHeaderFooter(g, page.Sheet.PrintFooter, page, normalFont, textBrush,
                            area.Left, area.Right, area.Bottom - 10f);
                    }
                }

                // Columnas visibles y sus posiciones.
                List<int> cols = new List<int>();
                List<float> colW = new List<float>();
                float x = x0;
                if (page.TitleCol1 >= 0)
                {
                    for (int c = page.TitleCol1; c <= page.TitleCol2; c++)
                    {
                        if (page.Sheet.HiddenCols.Contains(c)) continue;
                        cols.Add(c);
                        float w = ColPoints(page.Sheet, c) * scale;
                        colW.Add(w);
                        x += w;
                    }
                }
                for (int c = page.Col1; c <= page.Col2; c++)
                {
                    if (page.Sheet.HiddenCols.Contains(c)) continue;
                    cols.Add(c);
                    float w = ColPoints(page.Sheet, c) * scale;
                    colW.Add(w);
                    x += w;
                }
                float tableW = x - x0;

                List<int> rows = new List<int>();
                List<float> rowH = new List<float>();
                if (page.TitleRow1 >= 0)
                {
                    for (int r = page.TitleRow1; r <= page.TitleRow2; r++)
                    {
                        if (page.Sheet.HiddenRows.Contains(r)) continue;
                        rows.Add(r);
                        rowH.Add(RowPoints(page.Sheet, r) * scale);
                    }
                }
                for (int r = page.Row1; r <= page.Row2; r++)
                {
                    if (page.Sheet.HiddenRows.Contains(r)) continue;
                    rows.Add(r);
                    rowH.Add(RowPoints(page.Sheet, r) * scale);
                }

                float offsetX = 0;
                if (_options.CenterHorizontally && tableW < area.Width)
                    offsetX = (area.Width - tableW) / 2f;
                x0 += offsetX;

                // La tabla empieza arriba, como en Excel; centrarla verticalmente
                // separaba los encabezados de las celdas.
                if (_options.PrintHeadings)
                    DrawHeaders(g, page, area, x0, y0, scale);
                DrawCells(g, page, cols, colW, rows, rowH, x0, y0, scale);
                DrawImages(g, page, x0, y0, scale, area);
                DrawCharts(g, page, cols, colW, rows, rowH, x0, y0, scale, area);
            }
            finally
            {
                g.PageUnit = oldUnit;
            }
        }

        private void DrawHeaders(Graphics g, PrintPageInfo page, RectangleF area,
                                 float x0, float y0, float scale)
        {
            using (Font f = new Font("Segoe UI", Math.Max(3.5f, 6.5f * scale)))
            using (SolidBrush b = new SolidBrush(Color.DimGray))
            using (Pen p = new Pen(Color.Gray, Math.Max(0.25f, 0.5f * scale)))
            {
                float headerH = 16f * scale;
                float headerW = 24f * scale;
                float x = x0;
                for (int c = page.Col1; c <= page.Col2; c++)
                {
                    if (page.Sheet.HiddenCols.Contains(c)) continue;
                    float w = ColPoints(page.Sheet, c) * scale;
                    RectangleF r = new RectangleF(x, y0 - headerH + 2, w, Math.Max(8f, headerH - 2));
                    string name = AddressUtil.ColumnName(c);
                    SizeF sz = g.MeasureString(name, f);
                    g.DrawString(name, f, b, new PointF(r.Left + (r.Width - sz.Width) / 2f, r.Top));
                    g.DrawLine(p, x, y0 - 1, x + w, y0 - 1);
                    x += w;
                }
                float rowHeaderX = Math.Max(area.Left, x0 - headerW);
                float y = y0;
                for (int r = page.Row1; r <= page.Row2; r++)
                {
                    if (page.Sheet.HiddenRows.Contains(r)) continue;
                    float h = RowPoints(page.Sheet, r) * scale;
                    g.DrawString((r + 1).ToString(CultureInfo.InvariantCulture), f, b,
                        new PointF(rowHeaderX + 1, y + 1));
                    g.DrawLine(p, rowHeaderX, y, x0, y);
                    y += h;
                }
            }
        }

        private void DrawImages(Graphics g, PrintPageInfo page, float x0, float y0, float scale,
                                RectangleF area)
        {
            if (page.Sheet.Images.Count == 0) return;
            foreach (SheetImage image in page.Sheet.Images)
            {
                if (image.Image == null) continue;
                float x = x0, y = y0;
                if (image.FromCol < page.Col1)
                {
                    for (int c = image.FromCol; c < page.Col1; c++)
                        if (!page.Sheet.HiddenCols.Contains(c)) x -= ColPoints(page.Sheet, c) * scale;
                }
                else
                {
                    for (int c = page.Col1; c < image.FromCol; c++)
                        if (!page.Sheet.HiddenCols.Contains(c)) x += ColPoints(page.Sheet, c) * scale;
                }
                if (image.FromRow < page.Row1)
                {
                    for (int r = image.FromRow; r < page.Row1; r++)
                        if (!page.Sheet.HiddenRows.Contains(r)) y -= RowPoints(page.Sheet, r) * scale;
                }
                else
                {
                    for (int r = page.Row1; r < image.FromRow; r++)
                        if (!page.Sheet.HiddenRows.Contains(r)) y += RowPoints(page.Sheet, r) * scale;
                }
                x += image.FromColOff / 12700f * scale;
                y += image.FromRowOff / 12700f * scale;
                float w, h;
                if (image.ToCol >= image.FromCol && image.ToRow >= image.FromRow)
                {
                    w = image.ToColOff / 12700f * scale - image.FromColOff / 12700f * scale;
                    h = image.ToRowOff / 12700f * scale - image.FromRowOff / 12700f * scale;
                    for (int c = image.FromCol; c < image.ToCol; c++)
                        if (!page.Sheet.HiddenCols.Contains(c)) w += ColPoints(page.Sheet, c) * scale;
                    for (int r = image.FromRow; r < image.ToRow; r++)
                        if (!page.Sheet.HiddenRows.Contains(r)) h += RowPoints(page.Sheet, r) * scale;
                }
                else
                {
                    w = image.WidthPx * 72f / 96f * scale;
                    h = image.HeightPx * 72f / 96f * scale;
                }
                if (w < 4 || h < 4) continue;
                RectangleF rect = new RectangleF(x, y, w, h);
                RectangleF clip = RectangleF.Intersect(rect, area);
                if (clip.IsEmpty) continue;
                GraphicsState state = g.Save();
                try
                {
                    g.SetClip(clip);
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    g.DrawImage(image.Image, rect);
                }
                finally { g.Restore(state); }
            }
        }

        private void DrawCharts(Graphics g, PrintPageInfo page, List<int> cols, List<float> colW,
                                List<int> rows, List<float> rowH, float x0, float y0, float scale,
                                RectangleF area)
        {
            if (page.Sheet.Charts.Count == 0) return;
            foreach (SheetChart chart in page.Sheet.Charts)
            {
                float x = x0, y = y0;
                if (chart.FromCol < page.Col1)
                {
                    for (int c = chart.FromCol; c < page.Col1; c++)
                        if (!page.Sheet.HiddenCols.Contains(c)) x -= ColPoints(page.Sheet, c) * scale;
                }
                else
                {
                    for (int c = page.Col1; c < chart.FromCol; c++)
                        if (!page.Sheet.HiddenCols.Contains(c)) x += ColPoints(page.Sheet, c) * scale;
                }
                if (chart.FromRow < page.Row1)
                {
                    for (int r = chart.FromRow; r < page.Row1; r++)
                        if (!page.Sheet.HiddenRows.Contains(r)) y -= RowPoints(page.Sheet, r) * scale;
                }
                else
                {
                    for (int r = page.Row1; r < chart.FromRow; r++)
                        if (!page.Sheet.HiddenRows.Contains(r)) y += RowPoints(page.Sheet, r) * scale;
                }
                x += chart.FromColOff / 12700f * scale;
                y += chart.FromRowOff / 12700f * scale;
                float w, h;
                if (chart.ToCol >= chart.FromCol && chart.ToRow >= chart.FromRow)
                {
                    w = chart.ToColOff / 12700f * scale - chart.FromColOff / 12700f * scale;
                    h = chart.ToRowOff / 12700f * scale - chart.FromRowOff / 12700f * scale;
                    for (int c = chart.FromCol; c < chart.ToCol; c++)
                        if (!page.Sheet.HiddenCols.Contains(c)) w += ColPoints(page.Sheet, c) * scale;
                    for (int r = chart.FromRow; r < chart.ToRow; r++)
                        if (!page.Sheet.HiddenRows.Contains(r)) h += RowPoints(page.Sheet, r) * scale;
                }
                else
                {
                    w = chart.WidthPx * 72f / 96f * scale;
                    h = chart.HeightPx * 72f / 96f * scale;
                }
                if (w < 12 || h < 12) continue;
                RectangleF rect = new RectangleF(x, y, w, h);
                RectangleF clip = RectangleF.Intersect(rect, area);
                if (clip.IsEmpty) continue;
                GraphicsState state = g.Save();
                try
                {
                    g.SetClip(clip);
                    ChartRenderer.Draw(g, rect, chart, page.Sheet.Workbook, page.Sheet);
                }
                finally { g.Restore(state); }
            }
        }

        private void DrawCells(Graphics g, PrintPageInfo page, List<int> cols, List<float> colW,
                               List<int> rows, List<float> rowH, float x0, float y0, float scale)
        {
            FontCache fonts = new FontCache();
            try
            {
                float y = y0;
                for (int ri = 0; ri < rows.Count; ri++)
                {
                    int row = rows[ri];
                    float h = rowH[ri];
                    float x = x0;
                    for (int ci = 0; ci < cols.Count; ci++)
                    {
                        int col = cols[ci];
                        float w = colW[ci];
                        MergeRange merge = page.Sheet.GetMergeAt(row, col);
                        if (merge != null && (merge.Row1 != row || merge.Col1 != col))
                        {
                            x += w;
                            continue;
                        }
                        int r2 = row, c2 = col;
                        if (merge != null)
                        {
                            r2 = Math.Min(merge.Row2, page.Row2);
                            c2 = Math.Min(merge.Col2, page.Col2);
                        }
                        float cellW = w, cellH = h;
                        if (merge != null)
                        {
                            for (int k = ci + 1; k < cols.Count && cols[k] <= c2; k++) cellW += colW[k];
                            for (int k = ri + 1; k < rows.Count && rows[k] <= r2; k++) cellH += rowH[k];
                        }
                        Cell cell = page.Sheet.GetCell(row, col);
                        CellStyle st = cell == null ? CellStyle.Default : cell.EffectiveStyle;
                        RectangleF rect = new RectangleF(x, y, cellW, cellH);
                        Color fillColor = st.FillColor;
                        TableFormat table = TableFormatting.Find(page.Sheet, row, col);
                        if (table != null)
                        {
                            Color tableFill = TableFormatting.Fill(table, row, col);
                            if (!tableFill.IsEmpty) fillColor = tableFill;
                        }
                        PivotRange pivot = PivotFormatting.Find(page.Sheet, row, col);
                        if (pivot != null)
                        {
                            Color pivotFill = PivotFormatting.Fill(pivot, row, col);
                            if (!pivotFill.IsEmpty) fillColor = pivotFill;
                        }
                        ConditionalRule cf = ConditionalEvaluator.Find(page.Sheet, row, col);
                        if (cf != null)
                        {
                            if (!cf.ScaleFill.IsEmpty) fillColor = cf.ScaleFill;
                            else if (cf.Style != null && !cf.Style.FillColor.IsEmpty)
                                fillColor = cf.Style.FillColor;
                        }
                        if (!fillColor.IsEmpty)
                            using (SolidBrush fill = new SolidBrush(fillColor)) g.FillRectangle(fill, rect);

                        DrawBorders(g, rect, st.Borders, scale);

                        if (cell != null && cell.Value.Kind != ValueKind.Empty)
                        {
                            string text = NumberFormatter.Format(cell.Value, st);
                            if (!string.IsNullOrEmpty(text))
                            {
                                Color color = cell.Value.IsError ? Color.Firebrick : st.FontColor;
                                float pad = Math.Max(0.15f, 1.2f * scale);
                                float textW = Math.Max(0.5f, rect.Width - pad * 2);
                                int textAlign = NumberFormatter.EffectiveAlign(cell.Value, st);
                                if (!st.WrapText && textAlign == 0 && merge == null)
                                {
                                    for (int k = ci + 1; k < cols.Count; k++)
                                    {
                                        Cell next = page.Sheet.GetCell(row, cols[k]);
                                        if (next != null && next.Value.Kind != ValueKind.Empty) break;
                                        textW += colW[k];
                                    }
                                }
                                Font font = fonts.Get(st, scale);
                                Font fitted = null;
                                if (!st.WrapText)
                                {
                                    SizeF natural = g.MeasureString(text, font);
                                    if (natural.Width > textW && natural.Width > 0)
                                    {
                                        float factor = Math.Max(0.55f, textW / natural.Width);
                                        if (factor < 0.98f)
                                            fitted = new Font(font.FontFamily,
                                                Math.Max(4f, font.SizeInPoints * factor),
                                                font.Style, GraphicsUnit.Point);
                                    }
                                }
                                using (SolidBrush brush = new SolidBrush(color))
                                using (StringFormat sf = new StringFormat())
                                {
                                    int align = textAlign;
                                    sf.Alignment = align == 0 ? StringAlignment.Near :
                                                   align == 1 ? StringAlignment.Center : StringAlignment.Far;
                                    sf.LineAlignment = StringAlignment.Center;
                                    if (st.WrapText)
                                    {
                                        sf.Trimming = StringTrimming.None;
                                    }
                                    else
                                    {
                                        sf.Trimming = StringTrimming.EllipsisCharacter;
                                        sf.FormatFlags |= StringFormatFlags.NoWrap;
                                    }
                                    g.DrawString(text, fitted ?? font, brush,
                                        new RectangleF(rect.X + pad, rect.Y, textW, rect.Height), sf);
                                }
                                if (fitted != null) fitted.Dispose();
                            }
                        }
                        x += w;
                    }
                    y += h;
                }

                if (_options.PrintGridlines)
                {
                    using (Pen p = new Pen(Color.FromArgb(205, 205, 205), 0.5f))
                    {
                        float yy = y0;
                        for (int ri = 0; ri < rows.Count; ri++)
                        {
                            g.DrawLine(p, x0, yy, x0 + Sum(colW), yy);
                            yy += rowH[ri];
                        }
                        float xx = x0;
                        for (int ci = 0; ci < cols.Count; ci++)
                        {
                            g.DrawLine(p, xx, y0, xx, y0 + Sum(rowH));
                            xx += colW[ci];
                        }
                    }
                }
                DrawTableGrid(g, page, cols, colW, rows, rowH, x0, y0, scale);
            }
            finally
            {
                fonts.Dispose();
            }
        }

        private void DrawTableGrid(Graphics g, PrintPageInfo page, List<int> cols, List<float> colW,
                                   List<int> rows, List<float> rowH, float x0, float y0, float scale)
        {
            if (page.Sheet == null || page.Sheet.TableFormats.Count == 0) return;
            using (Pen pen = new Pen(Color.FromArgb(105, 105, 105), Math.Max(0.4f, 0.55f * scale)))
            {
                foreach (TableFormat table in page.Sheet.TableFormats)
                {
                    int c1 = Math.Max(table.Col1, page.Col1);
                    int c2 = Math.Min(table.Col2, page.Col2);
                    int r1 = Math.Max(table.Row1, page.Row1);
                    int r2 = Math.Min(table.Row2, page.Row2);
                    if (c1 > c2 || r1 > r2) continue;
                    float left = TableColX(cols, colW, x0, c1);
                    float right = TableColX(cols, colW, x0, c2 + 1);
                    float top = TableRowY(rows, rowH, y0, r1);
                    float bottom = TableRowY(rows, rowH, y0, r2 + 1);
                    g.DrawRectangle(pen, left, top, Math.Max(1, right - left), Math.Max(1, bottom - top));
                    for (int r = r1; r < r2; r++)
                    {
                        float yy = TableRowY(rows, rowH, y0, r + 1);
                        g.DrawLine(pen, left, yy, right, yy);
                    }
                    for (int c = c1; c < c2; c++)
                    {
                        float xx = TableColX(cols, colW, x0, c + 1);
                        g.DrawLine(pen, xx, top, xx, bottom);
                    }
                }
            }
        }

        private static float TableColX(List<int> cols, List<float> widths, float x0, int col)
        {
            float x = x0;
            for (int i = 0; i < cols.Count; i++)
            {
                if (cols[i] >= col) break;
                x += widths[i];
            }
            return x;
        }

        private static float TableRowY(List<int> rows, List<float> heights, float y0, int row)
        {
            float y = y0;
            for (int i = 0; i < rows.Count; i++)
            {
                if (rows[i] >= row) break;
                y += heights[i];
            }
            return y;
        }

        private static void DrawHeaderFooter(Graphics g, string raw, PrintPageInfo page,
                                             Font font, Brush brush, float left, float right, float y)
        {
            if (string.IsNullOrEmpty(raw)) return;
            string leftText = "", centerText = "", rightText = "";
            int section = 0;
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < raw.Length; i++)
            {
                if (raw[i] != '&' || i + 1 >= raw.Length)
                {
                    sb.Append(raw[i]);
                    continue;
                }
                char code = raw[i + 1];
                if (code == 'L' || code == 'C' || code == 'R')
                {
                    string text = CleanHeaderText(sb.ToString(), page);
                    if (section == 0) leftText += text;
                    else if (section == 1) centerText += text;
                    else rightText += text;
                    sb.Length = 0;
                    section = code == 'L' ? 0 : code == 'C' ? 1 : 2;
                    i++;
                }
                else if (code == '"')
                {
                    int close = raw.IndexOf('"', i + 2);
                    i = close < 0 ? raw.Length - 1 : close;
                }
                else if (char.IsDigit(code))
                {
                    while (i + 1 < raw.Length && char.IsDigit(raw[i + 1])) i++;
                }
                else if (code == 'B' || code == 'I' || code == 'U' || code == 'S' ||
                         code == 'K' || code == 'E' || code == 'X' || code == 'Y')
                {
                    i++;
                }
                else if (code == '&')
                {
                    sb.Append('&'); i++;
                }
                else sb.Append(raw[i]);
            }
            string tail = CleanHeaderText(sb.ToString(), page);
            if (section == 0) leftText += tail;
            else if (section == 1) centerText += tail;
            else rightText += tail;
            if (leftText.Length > 0)
                g.DrawString(leftText, font, brush, new PointF(left, y));
            if (centerText.Length > 0)
            {
                SizeF s = g.MeasureString(centerText, font);
                g.DrawString(centerText, font, brush, new PointF((left + right - s.Width) / 2, y));
            }
            if (rightText.Length > 0)
            {
                SizeF s = g.MeasureString(rightText, font);
                g.DrawString(rightText, font, brush, new PointF(right - s.Width, y));
            }
        }

        private static string CleanHeaderText(string text, PrintPageInfo page)
        {
            if (string.IsNullOrEmpty(text)) return "";
            return text.Replace("&P", page.PageNumber.ToString(CultureInfo.InvariantCulture))
                .Replace("&N", page.PageCount.ToString(CultureInfo.InvariantCulture))
                .Replace("&D", DateTime.Now.ToString("d", CultureInfo.CurrentCulture))
                .Replace("&T", DateTime.Now.ToString("t", CultureInfo.CurrentCulture))
                .Replace("&A", page.Sheet.Name)
                .Replace("&F", page.Sheet.Name);
        }

        private static float Sum(List<float> values)
        {
            float total = 0;
            foreach (float f in values) total += f;
            return total;
        }

        private static void DrawBorders(Graphics g, RectangleF rect, BorderFlags borders, float scale)
        {
            if (borders == BorderFlags.None) return;
            using (Pen p = new Pen(Color.Black, Math.Max(0.25f, 0.7f * scale)))
            {
                if ((borders & BorderFlags.Left) != 0) g.DrawLine(p, rect.Left, rect.Top, rect.Left, rect.Bottom);
                if ((borders & BorderFlags.Right) != 0) g.DrawLine(p, rect.Right, rect.Top, rect.Right, rect.Bottom);
                if ((borders & BorderFlags.Top) != 0) g.DrawLine(p, rect.Left, rect.Top, rect.Right, rect.Top);
                if ((borders & BorderFlags.Bottom) != 0) g.DrawLine(p, rect.Left, rect.Bottom, rect.Right, rect.Bottom);
            }
        }

        private sealed class FontCache : IDisposable
        {
            private readonly Dictionary<string, Font> _fonts = new Dictionary<string, Font>();
            public Font Get(CellStyle st, float scale)
            {
                float size = Math.Max(3.5f, Math.Min(72f, st.FontSize * scale));
                string key = (st.FontName ?? "Segoe UI") + "|" + size.ToString("0.##") +
                    "|" + st.Bold + "|" + st.Italic + "|" + st.Underline;
                Font f;
                if (_fonts.TryGetValue(key, out f)) return f;
                FontStyle style = FontStyle.Regular;
                if (st.Bold) style |= FontStyle.Bold;
                if (st.Italic) style |= FontStyle.Italic;
                if (st.Underline) style |= FontStyle.Underline;
                f = new Font(string.IsNullOrEmpty(st.FontName) ? "Segoe UI" : st.FontName,
                    size, style, GraphicsUnit.Point);
                _fonts[key] = f;
                return f;
            }
            public void Dispose()
            {
                foreach (Font f in _fonts.Values) f.Dispose();
                _fonts.Clear();
            }
        }
    }

    // Renderiza las páginas a mapas de bits para poder mostrarlas en un panel
    // desplazable. PrintPreviewControl no permite desplazamiento con la rueda.
    public sealed class BitmapPrintController : PrintController
    {
        public readonly List<Bitmap> Pages = new List<Bitmap>();
        private Bitmap _bitmap;
        private Graphics _graphics;

        public override bool IsPreview { get { return true; } }

        public override Graphics OnStartPage(PrintDocument document, PrintPageEventArgs e)
        {
            if (_graphics != null) _graphics.Dispose();
            if (_bitmap != null) _bitmap.Dispose();
            Rectangle bounds = document.DefaultPageSettings.Bounds;
            int width = Math.Max(1, bounds.Width * 96 / 100);
            int height = Math.Max(1, bounds.Height * 96 / 100);
            _bitmap = new Bitmap(width, height);
            _graphics = Graphics.FromImage(_bitmap);
            _graphics.Clear(Color.White);
            return _graphics;
        }

        public override void OnEndPage(PrintDocument document, PrintPageEventArgs e)
        {
            if (_bitmap != null)
            {
                Pages.Add(_bitmap);
                _bitmap = null;
            }
            if (_graphics != null)
            {
                _graphics.Dispose();
                _graphics = null;
            }
        }

        public override void OnEndPrint(PrintDocument document, PrintEventArgs e)
        {
            if (_graphics != null) { _graphics.Dispose(); _graphics = null; }
            if (_bitmap != null) { _bitmap.Dispose(); _bitmap = null; }
        }

        public void DisposePages()
        {
            foreach (Bitmap page in Pages) page.Dispose();
            Pages.Clear();
        }
    }

    public sealed class PrintPreviewForm : Form
    {
        private readonly WorkbookPrintDocument _document;
        private readonly Panel _scroll;
        private readonly PreviewCanvas _canvas;
        private readonly Label _pageLabel;
        private readonly ComboBox _zoom;
        private readonly List<Bitmap> _pages = new List<Bitmap>();
        private float _zoomFactor = 1f;

        public PrintPreviewForm(WorkbookPrintDocument document)
        {
            _document = document;
            Text = "Vista previa de impresión - Mini Excel";
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(1000, 760);
            MinimumSize = new Size(720, 520);
            Font = Theme.Font(9f);
            BackColor = Color.FromArgb(100, 100, 100);

            Panel tools = new Panel();
            tools.Dock = DockStyle.Top;
            tools.Height = 42;
            tools.BackColor = Color.FromArgb(245, 245, 245);
            Controls.Add(tools);

            Button print = new Button();
            print.Text = "Imprimir...";
            print.SetBounds(8, 7, 90, 28);
            print.Click += delegate { PrintNow(); };
            tools.Controls.Add(print);

            Button close = new Button();
            close.Text = "Cerrar";
            close.SetBounds(906, 7, 86, 28);
            close.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            close.DialogResult = DialogResult.Cancel;
            tools.Controls.Add(close);
            CancelButton = close;

            Label zoomLabel = new Label();
            zoomLabel.Text = "Zoom:";
            zoomLabel.SetBounds(118, 12, 42, 20);
            tools.Controls.Add(zoomLabel);

            _zoom = new ComboBox();
            _zoom.DropDownStyle = ComboBoxStyle.DropDownList;
            _zoom.Items.Add("50%"); _zoom.Items.Add("75%"); _zoom.Items.Add("100%");
            _zoom.Items.Add("125%"); _zoom.Items.Add("150%");
            _zoom.SelectedIndex = 2;
            _zoom.SetBounds(160, 8, 70, 24);
            _zoom.SelectedIndexChanged += delegate
            {
                _zoomFactor = _zoom.SelectedIndex == 0 ? 0.5f :
                    _zoom.SelectedIndex == 1 ? 0.75f :
                    _zoom.SelectedIndex == 2 ? 1f :
                    _zoom.SelectedIndex == 3 ? 1.25f : 1.5f;
                LayoutPages();
            };
            tools.Controls.Add(_zoom);

            _pageLabel = new Label();
            _pageLabel.SetBounds(250, 12, 260, 20);
            _pageLabel.Text = "Preparando vista previa...";
            tools.Controls.Add(_pageLabel);

            _scroll = new Panel();
            _scroll.Dock = DockStyle.Fill;
            _scroll.AutoScroll = true;
            _scroll.BackColor = Color.FromArgb(100, 100, 100);
            Controls.Add(_scroll);
            _scroll.BringToFront();

            _canvas = new PreviewCanvas();
            _canvas.Owner = this;
            _canvas.BackColor = Color.FromArgb(100, 100, 100);
            _canvas.Location = new Point(0, 0);
            _canvas.MouseWheel += CanvasWheel;
            _scroll.Controls.Add(_canvas);
            _scroll.MouseWheel += CanvasWheel;
            _scroll.Scroll += delegate { _canvas.Invalidate(); };

            Shown += delegate { RenderPages(); };
        }

        private void RenderPages()
        {
            Cursor = Cursors.WaitCursor;
            try
            {
                BitmapPrintController controller = new BitmapPrintController();
                PrintController previous = _document.PrintController;
                _document.PrintController = controller;
                try { _document.Print(); }
                finally { _document.PrintController = previous; }
                _pages.AddRange(controller.Pages);
                controller.Pages.Clear();
                _pageLabel.Text = _pages.Count + (_pages.Count == 1 ? " página" : " páginas");
                LayoutPages();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "No se pudo preparar la vista previa:\n\n" + ex.Message,
                    "Mini Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
            }
            finally { Cursor = Cursors.Default; }
        }

        private void LayoutPages()
        {
            if (_pages.Count == 0) { _canvas.Size = new Size(1, 1); return; }
            int maxW = 0, totalH = 24;
            foreach (Bitmap b in _pages)
                maxW = Math.Max(maxW, (int)(b.Width * _zoomFactor));
            foreach (Bitmap b in _pages)
                totalH += (int)(b.Height * _zoomFactor) + 24;
            _canvas.Size = new Size(maxW + 48, totalH + 8);
            _scroll.AutoScrollMinSize = _canvas.Size;
            _canvas.Invalidate();
        }

        private void CanvasWheel(object sender, MouseEventArgs e)
        {
            int amount = Math.Max(24, SystemInformation.MouseWheelScrollLines * 24);
            int target = -_scroll.AutoScrollPosition.Y - Math.Sign(e.Delta) * amount;
            _scroll.AutoScrollPosition = new Point(0, Math.Max(0, target));
            if (e is HandledMouseEventArgs) ((HandledMouseEventArgs)e).Handled = true;
        }

        private void PrintNow()
        {
            PrintController previous = _document.PrintController;
            _document.PrintController = new PrintControllerWithStatusDialog(new StandardPrintController());
            try
            {
                using (PrintDialog pd = new PrintDialog())
                {
                    pd.Document = _document;
                    pd.UseEXDialog = true;
                    if (pd.ShowDialog(this) == DialogResult.OK) _document.Print();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "No se pudo imprimir:\n\n" + ex.Message, "Mini Excel",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally { _document.PrintController = previous; }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                foreach (Bitmap b in _pages) b.Dispose();
                _pages.Clear();
            }
            base.Dispose(disposing);
        }

        private sealed class PreviewCanvas : Control
        {
            public PrintPreviewForm Owner;
            public PreviewCanvas()
            {
                SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                         ControlStyles.OptimizedDoubleBuffer, true);
                TabStop = true;
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                if (Owner == null) return;
                e.Graphics.Clear(BackColor);
                int y = 12;
                for (int i = 0; i < Owner._pages.Count; i++)
                {
                    Bitmap b = Owner._pages[i];
                    int w = (int)(b.Width * Owner._zoomFactor);
                    int h = (int)(b.Height * Owner._zoomFactor);
                    int x = Math.Max(12, (Width - w) / 2);
                    e.Graphics.DrawImage(b, new Rectangle(x, y, w, h));
                    using (Font f = Theme.Font(8f))
                    using (SolidBrush brush = new SolidBrush(Color.White))
                        e.Graphics.DrawString("Página " + (i + 1), f, brush, x + 2, y + 2);
                    y += h + 24;
                }
            }
        }
    }

    public sealed class PrintOptionsDialog : Form
    {
        private readonly Workbook _workbook;
        private readonly int _activeIndex;
        private readonly CheckedListBox _sheets;
        private readonly RadioButton _selection;
        private readonly RadioButton _active;
        private readonly RadioButton _selected;
        private readonly RadioButton _all;
        private readonly ComboBox _orientation;
        private readonly ComboBox _paper;
        private readonly NumericUpDown _margin;
        private readonly ComboBox _fit;
        private readonly CheckBox _gridlines;
        private readonly CheckBox _headings;
        public PrintOptions Options;

        public PrintOptionsDialog(Workbook workbook, int activeIndex)
        {
            _workbook = workbook;
            _activeIndex = activeIndex;
            Text = "Imprimir";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(470, 500);
            Font = Theme.Font(9f);

            GroupBox what = new GroupBox();
            what.Text = "Qué imprimir";
            what.SetBounds(12, 10, 446, 155);
            Controls.Add(what);

            _selection = Radio("Selección actual", 12, 22, what);
            _active = Radio("Hoja activa", 12, 46, what);
            _selected = Radio("Hojas seleccionadas", 12, 70, what);
            _all = Radio("Todas las hojas", 12, 94, what);
            _active.Checked = true;
            _selection.CheckedChanged += delegate { UpdateSheetList(); };
            _active.CheckedChanged += delegate { UpdateSheetList(); };
            _selected.CheckedChanged += delegate { UpdateSheetList(); };
            _all.CheckedChanged += delegate { UpdateSheetList(); };

            _sheets = new CheckedListBox();
            _sheets.SetBounds(190, 20, 240, 112);
            _sheets.CheckOnClick = true;
            for (int i = 0; i < workbook.Sheets.Count; i++)
                _sheets.Items.Add(workbook.Sheets[i].Name, i == activeIndex);
            what.Controls.Add(_sheets);

            GroupBox page = new GroupBox();
            page.Text = "Página";
            page.SetBounds(12, 175, 446, 150);
            Controls.Add(page);

            Label lo = new Label(); lo.Text = "Orientación:"; lo.SetBounds(12, 24, 80, 20); page.Controls.Add(lo);
            _orientation = new ComboBox(); _orientation.DropDownStyle = ComboBoxStyle.DropDownList;
            _orientation.Items.Add("Vertical"); _orientation.Items.Add("Horizontal");
            _orientation.SelectedIndex = 0; _orientation.SetBounds(100, 21, 120, 24); page.Controls.Add(_orientation);

            Label lp = new Label(); lp.Text = "Papel:"; lp.SetBounds(240, 24, 50, 20); page.Controls.Add(lp);
            _paper = new ComboBox(); _paper.DropDownStyle = ComboBoxStyle.DropDownList;
            _paper.Items.Add("A4"); _paper.Items.Add("Carta"); _paper.Items.Add("Legal");
            _paper.SelectedIndex = 0; _paper.SetBounds(292, 21, 130, 24); page.Controls.Add(_paper);

            Label lm = new Label(); lm.Text = "Margen (mm):"; lm.SetBounds(12, 58, 90, 20); page.Controls.Add(lm);
            _margin = new NumericUpDown(); _margin.Minimum = 5; _margin.Maximum = 40; _margin.Value = 12;
            _margin.SetBounds(100, 56, 70, 24); page.Controls.Add(_margin);

            Label lf = new Label(); lf.Text = "Ajuste:"; lf.SetBounds(12, 92, 80, 20); page.Controls.Add(lf);
            _fit = new ComboBox(); _fit.DropDownStyle = ComboBoxStyle.DropDownList;
            _fit.Items.Add("Tamaño real"); _fit.Items.Add("Ajustar a 1 página de ancho");
            _fit.Items.Add("Ajustar a 1 página"); _fit.SelectedIndex = 0;
            _fit.SetBounds(100, 90, 322, 24); page.Controls.Add(_fit);

            _gridlines = new CheckBox(); _gridlines.Text = "Imprimir líneas de cuadrícula";
            _gridlines.Checked = true; _gridlines.SetBounds(12, 335, 250, 22); Controls.Add(_gridlines);
            _headings = new CheckBox(); _headings.Text = "Imprimir encabezados de fila y columna";
            _headings.Checked = true; _headings.SetBounds(12, 360, 300, 22); Controls.Add(_headings);

            Label note = new Label();
            note.Text = "Se imprimen valores, formato, tablas, gráficos e imágenes compatibles.";
            note.ForeColor = Color.DimGray;
            note.SetBounds(12, 390, 446, 35);
            Controls.Add(note);

            Button preview = new Button(); preview.Text = "Vista previa";
            preview.SetBounds(170, 445, 100, 30); Controls.Add(preview);
            preview.Click += delegate { if (BuildOptions(true)) { DialogResult = DialogResult.OK; Close(); } };

            Button print = new Button(); print.Text = "Imprimir";
            print.SetBounds(278, 445, 90, 30); Controls.Add(print);
            print.Click += delegate { if (BuildOptions(false)) { DialogResult = DialogResult.OK; Close(); } };

            Button cancel = new Button(); cancel.Text = "Cancelar";
            cancel.SetBounds(376, 445, 82, 30); cancel.DialogResult = DialogResult.Cancel;
            Controls.Add(cancel);
            CancelButton = cancel;
            if (workbook != null && activeIndex >= 0 && activeIndex < workbook.Sheets.Count)
            {
                Sheet ps = workbook.Sheets[activeIndex];
                _orientation.SelectedIndex = ps.PrintLandscape ? 1 : 0;
                if (ps.PrintPaperSize == 1) _paper.SelectedIndex = 1;
                else if (ps.PrintPaperSize == 5) _paper.SelectedIndex = 2;
                if (ps.PrintMarginMm >= 5 && ps.PrintMarginMm <= 40)
                    _margin.Value = Math.Max(_margin.Minimum, Math.Min(_margin.Maximum,
                        (decimal)Math.Round(ps.PrintMarginMm)));
                _gridlines.Checked = ps.PrintGridLines;
                _headings.Checked = ps.PrintHeadings;
                if (ps.PrintFitToWidth) _fit.SelectedIndex = 1;
                else if (ps.PrintFitToHeight) _fit.SelectedIndex = 2;
            }
            UpdateSheetList();
        }

        private RadioButton Radio(string text, int x, int y, Control parent)
        {
            RadioButton rb = new RadioButton();
            rb.Text = text; rb.SetBounds(x, y, 170, 22);
            parent.Controls.Add(rb);
            return rb;
        }

        private void UpdateSheetList()
        {
            _sheets.Enabled = _selected.Checked;
        }

        private bool BuildOptions(bool preview)
        {
            PrintOptions o = new PrintOptions();
            o.PreviewRequested = preview;
            if (_selection.Checked) o.Scope = PrintScope.Selection;
            else if (_active.Checked) o.Scope = PrintScope.ActiveSheet;
            else if (_selected.Checked) o.Scope = PrintScope.SelectedSheets;
            else o.Scope = PrintScope.AllSheets;

            if (o.Scope == PrintScope.SelectedSheets)
            {
                for (int i = 0; i < _sheets.Items.Count; i++)
                    if (_sheets.GetItemChecked(i)) o.SheetIndices.Add(i);
                if (o.SheetIndices.Count == 0)
                {
                    MessageBox.Show(this, "Selecciona al menos una hoja.", "Imprimir",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return false;
                }
            }

            o.Landscape = _orientation.SelectedIndex == 1;
            o.Paper = _paper.SelectedIndex == 1 ? PaperKind.Letter :
                      _paper.SelectedIndex == 2 ? PaperKind.Legal : PaperKind.A4;
            o.MarginMm = (float)_margin.Value;
            o.Fit = _fit.SelectedIndex == 1 ? PrintFitMode.FitWidth :
                    _fit.SelectedIndex == 2 ? PrintFitMode.FitPage : PrintFitMode.Actual;
            o.PrintGridlines = _gridlines.Checked;
            o.PrintHeadings = _headings.Checked;
            Options = o;
            return true;
        }
    }
}
