using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Globalization;
using System.Windows.Forms;

namespace CalcRapido
{
    public sealed class SheetGrid : Control
    {
        public Workbook Workbook;
        public Sheet Sheet;
        private int FreezeRows { get { return Sheet == null ? 0 : Math.Max(0, Math.Min(Sheet.FreezeRows, _rowCount)); } }
        private int FreezeCols { get { return Sheet == null ? 0 : Math.Max(0, Math.Min(Sheet.FreezeCols, _colCount)); } }
        public bool ShowFormulas;
        public bool ShowGridLines = true;
        public bool ReadOnly { get { return Sheet != null && Sheet.IsProtected; } }
        public Func<string, string> FormulaValidator;

        public event EventHandler SelectionChanged;
        public event EventHandler Changed;
        public event EventHandler BeforeChange;
        public event EventHandler EditStateChanged;
        public event Action<string> EditTextChanged;
        public event EventHandler ZoomChanged;
        public event EventHandler ContextMenuRequested;

        private const int BaseRowHeight = 21;
        private const int BaseHeaderHeight = 23;
        private const int BaseHeaderWidth = 46;
        private const float MinZoom = 0.6f;
        private const float MaxZoom = 2.0f;

        private int _rowCount = 1000;
        private int _colCount = 256;
        private int _rowHeight = BaseRowHeight;
        private int _headerHeight = BaseHeaderHeight;
        private int _headerWidth = BaseHeaderWidth;
        private int[] _colLeft = new int[257];
        private int[] _rowTop = new int[1001];
        private float _zoom = 1f;
        private float _uiScale = 1f;
        private int _scrollX;
        private int _scrollY;
        private int _viewWidth;
        private int _viewHeight;
        private bool _syncingScroll;

        private int _activeRow;
        private int _activeCol;
        private int _selTop;
        private int _selLeft;
        private int _selBottom;
        private int _selRight;

        private HScrollBar _h;
        private VScrollBar _v;
        private CellEditor _editor;
        private bool _editing;
        private bool _closingEditor;
        private int _editRow;
        private int _editCol;
        private string _editOriginal;

        private enum DragMode { None, Cells, Columns, Rows, ResizeColumn, ResizeRow, Fill }
        private enum ChartDragMode { None, Move, Resize }
        private DragMode _drag;
        private ChartDragMode _chartDrag;
        private SheetChart _selectedChart;
        private Point _chartDragStart;
        private int _chartStartWidth;
        private int _chartStartHeight;
        private Rectangle _chartStartRect;
        private bool _chartPushed;
        private SheetImage _selectedImage;
        private bool _imageDrag;
        private bool _imageResize;
        private Point _imageDragStart;
        private Rectangle _imageStartRect;
        private long _imageStartFromColOff;
        private long _imageStartFromRowOff;
        private long _imageStartToColOff;
        private long _imageStartToRowOff;
        private int _imageStartWidthPx;
        private int _imageStartHeightPx;
        private bool _imagePushed;
        private int _anchorRow;
        private int _anchorCol;
        private int _resizeCol;
        private int _resizeStartX;
        private float _resizeStartWidth;
        private int _resizeRow;
        private int _resizeStartY;
        private float _resizeStartHeight;
        private bool _resizePushed;
        private int _fillTargetRow;
        private int _fillTargetCol;
        private bool _fillVertical;

        private ClipboardBlock _clipboard;
        private Timer _blink;
        private bool _blinkOn;
        private Timer _autoScroll;
        private Point _lastMouse;
        private readonly Dictionary<string, Font> _fontCache = new Dictionary<string, Font>();

        private sealed class CellEditor : TextBox
        {
            public SheetGrid Owner;
            protected override bool IsInputKey(Keys keyData)
            {
                Keys k = keyData & Keys.KeyCode;
                if (k == Keys.Tab || k == Keys.Enter || k == Keys.Escape) return true;
                return base.IsInputKey(keyData);
            }
            protected override void OnKeyDown(KeyEventArgs e)
            {
                if (Owner != null && Owner.EditorKeyDown(e))
                {
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    return;
                }
                base.OnKeyDown(e);
            }
        }

        private sealed class ClipboardBlock
        {
            public int Rows;
            public int Cols;
            public string[,] Raw;
            public CellStyle[,] Styles;
            public CellValue[,] Values;
            public bool[,] Formulas;
            public bool Cut;
            public Sheet SourceSheet;
            public int R1;
            public int C1;
            public string Text;
        }

        private readonly ToolTip _commentTip = new ToolTip();
        private string _lastComment;

        public SheetGrid()
        {
            _commentTip.InitialDelay = 250;
            _commentTip.ReshowDelay = 100;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw |
                     ControlStyles.Selectable, true);
            try
            {
                int dpi = DeviceDpi;
                if (dpi > 0) _uiScale = Math.Max(1f, dpi / 96f);
            }
            catch { _uiScale = 1f; }
            BackColor = Color.White;
            TabStop = true;
            Font = Theme.Font(10f);

            _h = new HScrollBar();
            _h.Minimum = 0;
            _h.Visible = true;
            _h.Scroll += delegate { if (!_syncingScroll) { _scrollX = _h.Value; Invalidate(); } };
            Controls.Add(_h);
            _v = new VScrollBar();
            _v.Minimum = 0;
            _v.Visible = true;
            _v.Scroll += delegate { if (!_syncingScroll) { _scrollY = _v.Value; Invalidate(); } };
            Controls.Add(_v);

            _blink = new Timer();
            _blink.Interval = 450;
            _blink.Tick += delegate { _blinkOn = !_blinkOn; Invalidate(); };

            _autoScroll = new Timer();
            _autoScroll.Interval = 60;
            _autoScroll.Tick += delegate { AutoScrollTick(); };
        }

        public bool IsEditing { get { return _editing; } }
        public string EditingText { get { return _editing && _editor != null ? _editor.Text : ""; } }
        public int ActiveRow { get { return _activeRow; } }
        public int ActiveCol { get { return _activeCol; } }
        public int SelectionTop { get { return _selTop; } }
        public int SelectionLeft { get { return _selLeft; } }
        public int SelectionBottom { get { return _selBottom; } }
        public int SelectionRight { get { return _selRight; } }
        public float Zoom { get { return _zoom; } }
        public int ScrollX { get { return _scrollX; } }
        public int ScrollY { get { return _scrollY; } }

        public void RestoreView(int activeRow, int activeCol, int selTop, int selLeft,
                                int selBottom, int selRight, int scrollX, int scrollY)
        {
            if (Sheet == null) return;
            activeRow = Math.Max(0, Math.Min(Limits.MaxRows - 1, activeRow));
            activeCol = Math.Max(0, Math.Min(Limits.MaxCols - 1, activeCol));
            int top = Math.Min(selTop, selBottom);
            int bottom = Math.Max(selTop, selBottom);
            int left = Math.Min(selLeft, selRight);
            int right = Math.Max(selLeft, selRight);
            selTop = Math.Max(0, Math.Min(Limits.MaxRows - 1, top));
            selBottom = Math.Max(selTop, Math.Min(Limits.MaxRows - 1, bottom));
            selLeft = Math.Max(0, Math.Min(Limits.MaxCols - 1, left));
            selRight = Math.Max(selLeft, Math.Min(Limits.MaxCols - 1, right));
            _activeRow = activeRow;
            _activeCol = activeCol;
            _anchorRow = activeRow;
            _anchorCol = activeCol;
            _selTop = selTop;
            _selLeft = selLeft;
            _selBottom = selBottom;
            _selRight = selRight;
            RebuildMetrics();
            UpdateScrollbars();
            SetScroll(scrollX, scrollY);
            Invalidate();
            RaiseSelectionChanged();
        }

        public string ActiveCellName { get { return AddressUtil.CellName(_activeRow, _activeCol); } }
        public string SelectionName { get { return AddressUtil.RangeName(_selTop, _selLeft, _selBottom, _selRight); } }
        public Rectangle SelectionBounds { get { return new Rectangle(_selLeft, _selTop, _selRight - _selLeft + 1, _selBottom - _selTop + 1); } }

        public void Attach(Workbook wb, int sheetIndex)
        {
            CancelEdit();
            Workbook = wb;
            if (wb == null) return;
            if (sheetIndex < 0 || sheetIndex >= wb.Sheets.Count) sheetIndex = 0;
            wb.ActiveSheetIndex = sheetIndex;
            Sheet = wb.Sheets[sheetIndex];
            _activeRow = 0; _activeCol = 0;
            _selTop = _selLeft = _selBottom = _selRight = 0;
            _scrollX = 0; _scrollY = 0;
            _clipboard = null;
            _selectedChart = null;
            _blink.Stop();
            RebuildMetrics();
            UpdateScrollbars();
            Invalidate();
        }

        public void EnsureChartVisible(SheetChart chart)
        {
            if (chart == null || Sheet == null) return;
            EnsureVisible(chart.FromRow, chart.FromCol);
            Invalidate();
        }

        public void ApplyValidation(int r1, int c1, int r2, int c2, string[] items, string message,
                                    bool allowBlank, string sourceRange = null)
        {
            if (!EnsureWritable()) return;
            if (Sheet == null || ((items == null || items.Length == 0) &&
                string.IsNullOrEmpty(sourceRange))) return;
            OnBeforeChange();
            Sheet.Validations.RemoveAll(delegate(DataValidationRule r)
            {
                return r.Row1 <= Math.Max(r1, r2) && r.Row2 >= Math.Min(r1, r2) &&
                       r.Col1 <= Math.Max(c1, c2) && r.Col2 >= Math.Min(c1, c2);
            });
            DataValidationRule rule = new DataValidationRule
            {
                Row1 = Math.Min(r1, r2), Col1 = Math.Min(c1, c2),
                Row2 = Math.Max(r1, r2), Col2 = Math.Max(c1, c2),
                SourceRange = sourceRange,
                Message = string.IsNullOrEmpty(message) ? "El valor no está en la lista." : message,
                AllowBlank = allowBlank
            };
            if (items != null) rule.Items.AddRange(items);
            Sheet.Validations.Add(rule);
            OnChanged();
            Invalidate();
        }

        public DataValidationRule ValidationAt(int row, int col)
        {
            if (Sheet == null) return null;
            foreach (DataValidationRule rule in Sheet.Validations)
            {
                if (!rule.Contains(row, col)) continue;
                if (rule.Items.Count == 0 && !string.IsNullOrEmpty(rule.SourceRange))
                {
                    int r1, c1, r2, c2;
                    string range = rule.SourceRange;
                    int bang = range.LastIndexOf('!');
                    if (bang >= 0) range = range.Substring(bang + 1).Replace("$", "");
                    if (AddressUtil.TryParseRange(range, out r1, out c1, out r2, out c2))
                    {
                        for (int r = Math.Min(r1, r2); r <= Math.Max(r1, r2) && rule.Items.Count < 1000; r++)
                            for (int c = Math.Min(c1, c2); c <= Math.Max(c1, c2) && rule.Items.Count < 1000; c++)
                            {
                                Cell cell = Sheet.GetCell(r, c);
                                if (cell == null || cell.Value.Kind == ValueKind.Empty) continue;
                                string text = NumberFormatter.Format(cell.Value, cell.EffectiveStyle);
                                if (!string.IsNullOrEmpty(text)) rule.Items.Add(text);
                            }
                    }
                }
                if (rule.Items.Count > 0) return rule;
            }
            return null;
        }

        public void ShowValidationList()
        {
            if (ReadOnly) return;
            DataValidationRule rule = ValidationAt(_activeRow, _activeCol);
            if (rule == null) return;
            ContextMenuStrip m = new ContextMenuStrip();
            foreach (string item in rule.Items)
            {
                string value = item;
                m.Items.Add(value, null, delegate
                {
                    OnBeforeChange();
                    Sheet.SetInput(_activeRow, _activeCol, value);
                    OnChanged();
                    Invalidate();
                });
            }
            Rectangle r = CellRect(_activeRow, _activeCol);
            m.Show(this, new Point(r.Left, r.Bottom));
        }

        public void SetAutoFilterSelection()
        {
            if (!EnsureWritable()) return;
            if (Sheet == null) return;
            Sheet.AutoFilterRef = AddressUtil.CellName(_selTop, _selLeft) + ":" +
                AddressUtil.CellName(_selBottom, _selRight);
            Sheet.RowsDirty = true;
            Invalidate();
        }

        private bool TryFilterRange(out int r1, out int c1, out int r2, out int c2)
        {
            r1 = c1 = r2 = c2 = 0;
            if (Sheet == null || string.IsNullOrEmpty(Sheet.AutoFilterRef)) return false;
            return AddressUtil.TryParseRange(Sheet.AutoFilterRef, out r1, out c1, out r2, out c2);
        }

        private bool TryFilterRangeForColumn(int column, out int r1, out int c1, out int r2, out int c2)
        {
            if (TryFilterRange(out r1, out c1, out r2, out c2) &&
                column >= c1 && column <= c2) return true;
            if (Sheet != null)
            {
                foreach (TableFormat table in Sheet.TableFormats)
                {
                    if (column < table.Col1 || column > table.Col2 || table.Row2 <= table.Row1)
                        continue;
                    r1 = table.Row1; c1 = table.Col1; r2 = table.Row2; c2 = table.Col2;
                    return true;
                }
            }
            r1 = c1 = r2 = c2 = 0;
            return false;
        }

        private void ApplyFilterAt(int column, string value, bool all)
        {
            if (!EnsureWritable()) return;
            int r1, c1, r2, c2;
            if (!TryFilterRangeForColumn(column, out r1, out c1, out r2, out c2)) return;
            OnBeforeChange();
            for (int r = r1 + 1; r <= r2; r++) Sheet.HiddenRows.Remove(r);
            if (!all)
            {
                for (int r = r1 + 1; r <= r2; r++)
                {
                    Cell cell = Sheet.GetCell(r, column);
                    string text = cell == null ? "" : NumberFormatter.Format(cell.Value, cell.EffectiveStyle);
                    if (!string.Equals(text, value ?? "", StringComparison.CurrentCultureIgnoreCase))
                        Sheet.HiddenRows.Add(r);
                }
            }
            Sheet.FilterSelections.RemoveAll(delegate(FilterSelection f) { return f.Column == column; });
            if (!all)
            {
                FilterSelection selection = new FilterSelection { Column = column };
                selection.Values.Add(value ?? "");
                Sheet.FilterSelections.Add(selection);
            }
            Sheet.RowsDirty = true;
            RebuildMetrics();
            OnChanged();
            Invalidate();
        }

        private Rectangle FilterArrowRect(int col)
        {
            int r1, c1, r2, c2;
            if (!TryFilterRangeForColumn(col, out r1, out c1, out r2, out c2)) return Rectangle.Empty;
            Rectangle cell = CellRect(r1, col);
            if (cell.Width < 18 || cell.Height < 12) return Rectangle.Empty;
            int size = Math.Min(14, Math.Max(10, cell.Height - 4));
            return new Rectangle(cell.Right - size - 2,
                cell.Y + Math.Max(1, (cell.Height - size) / 2), size, size);
        }

        private void DrawFilterButtons(Graphics g, int firstRow, int firstCol, int maxRow, int maxCol)
        {
            for (int c = firstCol; c <= maxCol; c++)
            {
                if (Sheet.HiddenCols.Contains(c)) continue;
                int r1, c1, r2, c2;
                if (!TryFilterRangeForColumn(c, out r1, out c1, out r2, out c2)) continue;
                if (r1 > maxRow) continue;
                if (r1 < firstRow && r1 >= FreezeRows) continue;
                bool filtered = false;
                for (int r = r1 + 1; r <= r2; r++)
                    if (Sheet.HiddenRows.Contains(r)) { filtered = true; break; }
                Rectangle arrow = FilterArrowRect(c);
                if (arrow.IsEmpty) continue;
                using (SolidBrush fill = new SolidBrush(filtered
                    ? Color.FromArgb(222, 235, 247) : Color.FromArgb(245, 245, 245)))
                using (Pen border = new Pen(Color.FromArgb(150, 150, 150)))
                {
                    g.FillRectangle(fill, arrow);
                    g.DrawRectangle(border, arrow);
                    PointF[] tri = new PointF[]
                    {
                        new PointF(arrow.Left + 3, arrow.Top + 4),
                        new PointF(arrow.Right - 3, arrow.Top + 4),
                        new PointF(arrow.Left + arrow.Width / 2f, arrow.Bottom - 3)
                    };
                    using (SolidBrush b = new SolidBrush(Color.FromArgb(80, 80, 80)))
                        g.FillPolygon(b, tri);
                }
            }
        }

        private void ShowFilterMenu(int col)
        {
            int r1, c1, r2, c2;
            if (!TryFilterRangeForColumn(col, out r1, out c1, out r2, out c2)) return;
            ContextMenuStrip menu = new ContextMenuStrip();
            menu.Items.Add("(Todo)", null, delegate { ApplyFilterAt(col, null, true); });
            menu.Items.Add(new ToolStripSeparator());
            HashSet<string> seen = new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);
            for (int r = r1 + 1; r <= r2; r++)
            {
                Cell cell = Sheet.GetCell(r, col);
                string text = cell == null ? "" : NumberFormatter.Format(cell.Value, cell.EffectiveStyle);
                if (!seen.Add(text)) continue;
                string value = text;
                menu.Items.Add(value.Length == 0 ? "(Vacías)" : value, null,
                    delegate { ApplyFilterAt(col, value, false); });
                if (menu.Items.Count > 60) break;
            }
            Rectangle arrowRect = FilterArrowRect(col);
            menu.Show(this, new Point(arrowRect.Left, arrowRect.Bottom));
        }

        public void ApplyFilter(int column, int op, string text)
        {
            if (!EnsureWritable()) return;
            if (Sheet == null) return;
            int c = Math.Max(0, Math.Min(column, _colCount - 1));
            int first = Math.Max(0, _selTop);
            int last = Math.Max(first, Math.Min(_selBottom, _rowCount - 1));
            OnBeforeChange();
            for (int r = first + 1; r <= last; r++) Sheet.HiddenRows.Remove(r);
            for (int r = first + 1; r <= last; r++)
            {
                Cell cell = Sheet.GetCell(r, c);
                string value = cell == null ? "" : NumberFormatter.Format(cell.Value, cell.EffectiveStyle);
                bool match;
                if (op == 2) match = !string.IsNullOrEmpty(value);
                else if (op == 3) match = string.IsNullOrEmpty(value);
                else if (op == 1) match = string.Equals(value, text ?? "",
                    StringComparison.CurrentCultureIgnoreCase);
                else match = value != null && value.IndexOf(text ?? "",
                    StringComparison.CurrentCultureIgnoreCase) >= 0;
                if (!match) Sheet.HiddenRows.Add(r);
            }
            Sheet.RowsDirty = true;
            RebuildMetrics();
            OnChanged();
            Invalidate();
        }

        public void ClearFilter()
        {
            if (!EnsureWritable()) return;
            if (Sheet == null) return;
            OnBeforeChange();
            int first = Math.Max(0, _selTop);
            int last = Math.Max(first, Math.Min(_selBottom, _rowCount - 1));
            for (int r = first + 1; r <= last; r++) Sheet.HiddenRows.Remove(r);
            Sheet.FilterSelections.Clear();
            Sheet.RowsDirty = true;
            RebuildMetrics();
            OnChanged();
            Invalidate();
        }

        public void FreezePanesAtActive()
        {
            if (Sheet == null) return;
            Sheet.FreezeRows = Math.Max(0, Math.Min(_activeRow, _rowCount));
            Sheet.FreezeCols = Math.Max(0, Math.Min(_activeCol, _colCount));
            Invalidate();
        }

        public void UnfreezePanes()
        {
            if (Sheet == null) return;
            Sheet.FreezeRows = 0;
            Sheet.FreezeCols = 0;
            Invalidate();
        }

        public void SelectChart(SheetChart chart)
        {
            _selectedChart = chart;
            EnsureChartVisible(chart);
        }

        public bool HasSelectedImage
        {
            get { return _selectedImage != null && Sheet != null && Sheet.Images.Contains(_selectedImage); }
        }

        public SheetImage SelectedImage
        {
            get { return HasSelectedImage ? _selectedImage : null; }
        }

        public void CropSelectedImage(int left, int top, int right, int bottom)
        {
            if (!EnsureWritable()) return;
            if (!HasSelectedImage || _selectedImage.Image == null) return;
            Image src = _selectedImage.Image;
            int x = Math.Max(0, src.Width * Math.Min(45, Math.Max(0, left)) / 100);
            int y = Math.Max(0, src.Height * Math.Min(45, Math.Max(0, top)) / 100);
            int w = Math.Max(1, src.Width - x - src.Width * Math.Min(45, Math.Max(0, right)) / 100);
            int h = Math.Max(1, src.Height - y - src.Height * Math.Min(45, Math.Max(0, bottom)) / 100);
            if (w <= 1 || h <= 1) return;
            OnBeforeChange();
            Bitmap cropped = new Bitmap(w, h);
            using (Graphics g = Graphics.FromImage(cropped))
                g.DrawImage(src, new Rectangle(0, 0, w, h), new Rectangle(x, y, w, h), GraphicsUnit.Pixel);
            _selectedImage.Image = cropped;
            _selectedImage.WidthPx = Math.Max(20, _selectedImage.WidthPx * w / Math.Max(1, src.Width));
            _selectedImage.HeightPx = Math.Max(20, _selectedImage.HeightPx * h / Math.Max(1, src.Height));
            _selectedImage.Dirty = true;
            Sheet.ImagesDirty = true;
            OnChanged();
            Invalidate();
        }

        public void DeleteSelectedImage()
        {
            if (!EnsureWritable()) return;
            if (!HasSelectedImage) return;
            OnBeforeChange();
            Sheet.DeletedImages.Add(_selectedImage);
            Sheet.Images.Remove(_selectedImage);
            Sheet.ImagesDirty = true;
            _selectedImage = null;
            OnChanged();
            Invalidate();
        }

        public bool HasSelectedChart
        {
            get { return _selectedChart != null && Sheet != null && Sheet.Charts.Contains(_selectedChart); }
        }

        public SheetChart SelectedChart
        {
            get { return HasSelectedChart ? _selectedChart : null; }
        }

        public void DeleteSelectedChart()
        {
            if (!EnsureWritable()) return;
            if (!HasSelectedChart) return;
            OnBeforeChange();
            Sheet.Charts.Remove(_selectedChart);
            _selectedChart = null;
            Sheet.ChartsDirty = true;
            OnChanged();
            Invalidate();
        }

        public void SetSheet(Sheet sheet)
        {
            CancelEdit();
            Sheet = sheet;
            if (Workbook != null)
            {
                int idx = Workbook.Sheets.IndexOf(sheet);
                if (idx >= 0) Workbook.ActiveSheetIndex = idx;
            }
            _activeRow = 0; _activeCol = 0;
            _selTop = _selLeft = _selBottom = _selRight = 0;
            _scrollX = 0; _scrollY = 0;
            _selectedChart = null;
            RebuildMetrics();
            UpdateScrollbars();
            Invalidate();
            RaiseSelectionChanged();
        }

        public void SetZoom(float z)
        {
            z = Math.Max(MinZoom, Math.Min(MaxZoom, z));
            if (Math.Abs(z - _zoom) < 0.001f) return;
            int firstRow = _rowHeight <= 0 ? 0 : _scrollY / _rowHeight;
            int firstCol = ColAtX(_scrollX);
            if (firstCol < 0) firstCol = 0;
            _zoom = z;
            ClearFontCache();
            RebuildMetrics();
            _scrollX = _colLeft[Math.Min(firstCol, _colCount - 1)];
            _scrollY = _rowTop[Math.Min(firstRow, _rowCount - 1)];
            UpdateScrollbars();
            Invalidate();
            if (ZoomChanged != null) ZoomChanged(this, EventArgs.Empty);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_blink != null) { _blink.Stop(); _blink.Dispose(); _blink = null; }
                if (_autoScroll != null) { _autoScroll.Stop(); _autoScroll.Dispose(); _autoScroll = null; }
                ClearFontCache();
            }
            base.Dispose(disposing);
        }

        private void ClearFontCache()
        {
            foreach (Font f in _fontCache.Values) f.Dispose();
            _fontCache.Clear();
        }

        private Font GetFont(CellStyle st)
        {
            return GetFont(st, false);
        }

        private Font GetFont(CellStyle st, bool forceUnderline)
        {
            if (st == null) st = CellStyle.Default;
            int sizeKey = (int)Math.Round(st.FontSize * _zoom * 10);
            string key = (st.FontName ?? "Segoe UI") + "|" + sizeKey + "|" + (st.Bold ? "B" : "") +
                         (st.Italic ? "I" : "") + ((st.Underline || forceUnderline) ? "U" : "") + (forceUnderline ? "L" : "");
            Font f;
            if (_fontCache.TryGetValue(key, out f)) return f;
            FontStyle fs = FontStyle.Regular;
            if (st.Bold) fs |= FontStyle.Bold;
            if (st.Italic) fs |= FontStyle.Italic;
            if (st.Underline || forceUnderline) fs |= FontStyle.Underline;
            float size = Math.Max(5f, st.FontSize * _zoom);
            try { f = new Font(st.FontName ?? "Segoe UI", size, fs, GraphicsUnit.Point); }
            catch { f = new Font("Segoe UI", size, fs, GraphicsUnit.Point); }
            _fontCache[key] = f;
            return f;
        }

        private void RebuildMetrics()
        {
            if (Sheet == null) return;
            _rowHeight = Math.Max(14, (int)Math.Round(BaseRowHeight * _zoom * _uiScale));
            _headerHeight = Math.Max(18, (int)Math.Round(BaseHeaderHeight * _zoom * _uiScale));
            _headerWidth = Math.Max(34, (int)Math.Round(BaseHeaderWidth * _zoom * _uiScale));

            int needRows = Math.Max(1000, Sheet.UsedMaxRow + 100);
            needRows = Math.Max(needRows, _activeRow + 100);
            if (needRows > Limits.MaxRows) needRows = Limits.MaxRows;
            _rowCount = needRows;
            int needCols = Math.Max(256, Sheet.UsedMaxCol + 8);
            needCols = Math.Max(needCols, _activeCol + 8);
            if (needCols > Limits.MaxCols) needCols = Limits.MaxCols;
            _colCount = needCols;

            _colLeft = new int[_colCount + 1];
            int x = 0;
            for (int c = 0; c < _colCount; c++)
            {
                _colLeft[c] = x;
                if (Sheet.HiddenCols.Contains(c)) continue;
                int w = (int)Math.Round(Sheet.GetColWidth(c) * _zoom * _uiScale);
                x += Math.Max(18, w);
            }
            _colLeft[_colCount] = x;

            _rowTop = new int[_rowCount + 1];
            int y = 0;
            for (int r = 0; r < _rowCount; r++)
            {
                _rowTop[r] = y;
                if (Sheet.HiddenRows.Contains(r)) continue;
                float points = Sheet.GetRowHeightPoints(r);
                int h = (int)Math.Round(points * (4.0 / 3.0) * _zoom * _uiScale);
                y += Math.Max(7, h);
            }
            _rowTop[_rowCount] = y;
            UpdateScrollbars();
        }

        private void ExpandTo(int row, int col)
        {
            bool changed = false;
            if (row >= _rowCount - 2)
            {
                int n = Math.Min(Limits.MaxRows, Math.Max(row + 100, _rowCount + 200));
                if (n != _rowCount) { _rowCount = n; changed = true; }
            }
            if (col >= _colCount - 1)
            {
                int n = Math.Min(Limits.MaxCols, Math.Max(col + 10, _colCount + 32));
                if (n != _colCount) { _colCount = n; changed = true; }
            }
            if (changed) RebuildMetrics();
        }

        private int TotalWidth
        {
            get { return _colLeft == null || _colCount < 0 || _colCount >= _colLeft.Length ? 0 : _colLeft[_colCount]; }
        }
        private int TotalHeight
        {
            get { return _rowTop == null || _rowCount < 0 || _rowCount >= _rowTop.Length ? 0 : _rowTop[_rowCount]; }
        }

        public int ViewWidth { get { return _viewWidth; } }
        public int ViewHeight { get { return _viewHeight; } }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            LayoutScrollbars();
        }

        protected override void OnLayout(LayoutEventArgs levent)
        {
            base.OnLayout(levent);
            LayoutScrollbars();
        }

        private void LayoutScrollbars()
        {
            if (_h == null || _v == null) return;
            int hh = SystemInformation.HorizontalScrollBarHeight;
            int vw = SystemInformation.VerticalScrollBarWidth;
            if (hh <= 0) hh = 17;
            if (vw <= 0) vw = 17;
            int w = Math.Max(0, ClientSize.Width - vw);
            int h = Math.Max(0, ClientSize.Height - hh);
            _viewWidth = w;
            _viewHeight = h;
            _v.Bounds = new Rectangle(w, 0, vw, h);
            _h.Bounds = new Rectangle(0, h, w, hh);
            UpdateScrollbars();
        }

        private void UpdateScrollbars()
        {
            if (_h == null || _v == null) return;
            int vw = Math.Max(1, _viewWidth);
            int vh = Math.Max(1, _viewHeight);
            int maxX = Math.Max(0, TotalWidth - vw);
            int maxY = Math.Max(0, TotalHeight - vh);
            if (_scrollX > maxX) _scrollX = maxX;
            if (_scrollY > maxY) _scrollY = maxY;
            _syncingScroll = true;
            try
            {
                _h.Maximum = Math.Max(0, TotalWidth - 1);
                _h.LargeChange = Math.Max(1, Math.Min(vw, Math.Max(1, TotalWidth)));
                _h.SmallChange = Math.Max(8, (int)(20 * _zoom * _uiScale));
                _h.Value = Math.Min(_scrollX, Math.Max(0, _h.Maximum - _h.LargeChange + 1));
                _v.Maximum = Math.Max(0, TotalHeight - 1);
                _v.LargeChange = Math.Max(1, Math.Min(vh, Math.Max(1, TotalHeight)));
                _v.SmallChange = Math.Max(8, _rowHeight);
                _v.Value = Math.Min(_scrollY, Math.Max(0, _v.Maximum - _v.LargeChange + 1));
                _scrollX = _h.Value;
                _scrollY = _v.Value;
            }
            finally
            {
                _syncingScroll = false;
            }
        }

        private void SetScroll(int x, int y)
        {
            int maxX = Math.Max(0, TotalWidth - Math.Max(1, _viewWidth));
            int maxY = Math.Max(0, TotalHeight - Math.Max(1, _viewHeight));
            x = Math.Max(0, Math.Min(maxX, x));
            y = Math.Max(0, Math.Min(maxY, y));
            _scrollX = x;
            _scrollY = y;
            _syncingScroll = true;
            try
            {
                _h.Value = Math.Min(x, Math.Max(0, _h.Maximum - _h.LargeChange + 1));
                _v.Value = Math.Min(y, Math.Max(0, _v.Maximum - _v.LargeChange + 1));
                _scrollX = _h.Value;
                _scrollY = _v.Value;
            }
            finally
            {
                _syncingScroll = false;
            }
            Invalidate();
        }

        private int ColAtX(int contentX)
        {
            if (contentX < 0 || _colCount <= 0) return -1;
            int lo = 0, hi = _colCount - 1;
            while (lo <= hi)
            {
                int mid = (lo + hi) / 2;
                if (contentX < _colLeft[mid]) hi = mid - 1;
                else if (contentX >= _colLeft[mid + 1]) lo = mid + 1;
                else
                {
                    int c = mid;
                    while (c < _colCount - 1 && _colLeft[c + 1] == _colLeft[c]) c++;
                    return c;
                }
            }
            return -1;
        }

        private Rectangle CellRect(int row, int col)
        {
            if (col < 0 || col >= _colCount || row < 0 || row >= _rowCount)
                return Rectangle.Empty;
            int sx = FreezeCols > 0 && col < FreezeCols ? 0 : _scrollX;
            int sy = FreezeRows > 0 && row < FreezeRows ? 0 : _scrollY;
            int x = _headerWidth + _colLeft[col] - sx;
            int y = _headerHeight + _rowTop[row] - sy;
            int w = _colLeft[col + 1] - _colLeft[col];
            int h = _rowTop[row + 1] - _rowTop[row];
            return new Rectangle(x, y, w, h);
        }

        private int RowAtY(int contentY)
        {
            if (contentY < 0) return 0;
            if (_rowCount <= 0) return 0;
            if (contentY >= _rowTop[_rowCount]) return _rowCount - 1;
            int lo = 0, hi = _rowCount - 1;
            while (lo <= hi)
            {
                int mid = (lo + hi) / 2;
                if (contentY < _rowTop[mid]) hi = mid - 1;
                else if (contentY >= _rowTop[mid + 1]) lo = mid + 1;
                else
                {
                    int r = mid;
                    while (r < _rowCount - 1 && _rowTop[r + 1] == _rowTop[r]) r++;
                    return r;
                }
            }
            return Math.Max(0, Math.Min(_rowCount - 1, lo));
        }

        private void UpdateCommentTooltip(Point p)
        {
            if (Sheet == null) return;
            int row, col;
            if (!TryPointToCell(p, out row, out col))
            {
                if (_lastComment != null) { _lastComment = null; _commentTip.SetToolTip(this, ""); }
                return;
            }
            Cell cell = Sheet.GetCell(row, col);
            string text = cell == null ? null : cell.Comment;
            if (text == _lastComment) return;
            _lastComment = text;
            _commentTip.SetToolTip(this, string.IsNullOrEmpty(text) ? "" : text);
        }

        private bool TryPointToCell(Point p, out int row, out int col)
        {
            row = -1; col = -1;
            if (p.X < _headerWidth || p.Y < _headerHeight || p.X >= _viewWidth || p.Y >= _viewHeight) return false;
            int frozenW = FreezeCols > 0 ? _colLeft[FreezeCols] : 0;
            int frozenH = FreezeRows > 0 ? _rowTop[FreezeRows] : 0;
            int cx = p.X - _headerWidth < frozenW ? p.X - _headerWidth
                : _scrollX + (p.X - _headerWidth);
            int cy = p.Y - _headerHeight < frozenH ? p.Y - _headerHeight
                : _scrollY + (p.Y - _headerHeight);
            col = ColAtX(cx);
            row = RowAtY(cy);
            return col >= 0 && row >= 0 && row < _rowCount && col < _colCount;
        }

        private void PointToCellClamped(Point p, out int row, out int col)
        {
            int frozenW = FreezeCols > 0 ? _colLeft[FreezeCols] : 0;
            int frozenH = FreezeRows > 0 ? _rowTop[FreezeRows] : 0;
            int cx = Math.Max(0, p.X - _headerWidth) < frozenW ? Math.Max(0, p.X - _headerWidth)
                : _scrollX + Math.Max(0, p.X - _headerWidth);
            int cy = Math.Max(0, p.Y - _headerHeight) < frozenH ? Math.Max(0, p.Y - _headerHeight)
                : _scrollY + Math.Max(0, p.Y - _headerHeight);
            col = ColAtX(cx);
            if (col < 0) col = 0;
            row = RowAtY(cy);
            if (p.X >= _viewWidth) col = Math.Min(_colCount - 1, ColAtX(Math.Max(0, _scrollX + _viewWidth - _headerWidth - 1)) + 1);
            if (p.Y >= _viewHeight) row = Math.Min(_rowCount - 1, RowAtY(_scrollY + _viewHeight - _headerHeight) + 1);
            if (col < 0) col = 0;
            if (row < 0) row = 0;
        }

        private void EnsureVisible(int row, int col)
        {
            ExpandTo(row, col);
            int top = _headerHeight + _rowTop[row] - _scrollY;
            int bottom = _headerHeight + _rowTop[row + 1] - _scrollY;
            if (top < _headerHeight) _scrollY = Math.Max(0, _scrollY - (_headerHeight - top));
            else if (bottom > _viewHeight - 2) _scrollY += bottom - (_viewHeight - 2);
            int x = _headerWidth + _colLeft[col] - _scrollX;
            int right = x + (_colLeft[col + 1] - _colLeft[col]);
            if (x < _headerWidth) _scrollX = Math.Max(0, _scrollX - (_headerWidth - x));
            else if (right > _viewWidth - 2) _scrollX += right - (_viewWidth - 2);
            SetScroll(_scrollX, _scrollY);
        }

        public void SelectCell(int row, int col)
        {
            SelectRange(row, col, row, col);
        }

        private Rectangle GetDisplayRect(int row, int col)
        {
            if (Sheet != null)
            {
                MergeRange m = Sheet.GetMergeAt(row, col);
                if (m != null)
                    return RectForCells(m.Row1, m.Col1, m.Row2, m.Col2);
            }
            return CellRect(row, col);
        }

        public void MoveActiveCell(int dRow, int dCol)
        {
            MoveActive(dRow, dCol, false, true);
        }

        public void SelectRange(int r1, int c1, int r2, int c2)
        {
            if (Sheet == null) return;
            if (IsEditing) CancelEdit();
            r1 = Math.Max(0, Math.Min(Limits.MaxRows - 1, r1));
            r2 = Math.Max(0, Math.Min(Limits.MaxRows - 1, r2));
            c1 = Math.Max(0, Math.Min(Limits.MaxCols - 1, c1));
            c2 = Math.Max(0, Math.Min(Limits.MaxCols - 1, c2));
            _selTop = Math.Min(r1, r2); _selBottom = Math.Max(r1, r2);
            _selLeft = Math.Min(c1, c2); _selRight = Math.Max(c1, c2);
            _activeRow = r1; _activeCol = c1;
            if (r1 == r2 && c1 == c2)
            {
                MergeRange m = Sheet.GetMergeAt(r1, c1);
                if (m != null)
                {
                    _selTop = m.Row1; _selBottom = m.Row2;
                    _selLeft = m.Col1; _selRight = m.Col2;
                    _activeRow = m.Row1; _activeCol = m.Col1;
                }
            }
            _anchorRow = _activeRow; _anchorCol = _activeCol;
            ExpandTo(_selBottom, _selRight);
            EnsureVisible(_activeRow, _activeCol);
            Invalidate();
            RaiseSelectionChanged();
        }

        private void RaiseSelectionChanged()
        {
            if (SelectionChanged != null) SelectionChanged(this, EventArgs.Empty);
        }

        private void MoveActive(int dr, int dc, bool extend, bool ensure)
        {
            int r = Math.Max(0, Math.Min(Limits.MaxRows - 1, _activeRow + dr));
            int c = Math.Max(0, Math.Min(Limits.MaxCols - 1, _activeCol + dc));
            if (Sheet != null)
            {
                // Al salir de una celda combinada hay que saltar su rango
                // completo; si no, el destino vuelve al vértice y la flecha
                // queda atrapada dentro de la combinación.
                MergeRange source = Sheet.GetMergeAt(_activeRow, _activeCol);
                if (source != null)
                {
                    if (dc > 0) c = Math.Min(Limits.MaxCols - 1, source.Col2 + 1);
                    else if (dc < 0) c = Math.Max(0, source.Col1 - 1);
                    if (dr > 0) r = Math.Min(Limits.MaxRows - 1, source.Row2 + 1);
                    else if (dr < 0) r = Math.Max(0, source.Row1 - 1);
                }
            }
            if (Sheet != null && dr != 0)
            {
                int step = dr > 0 ? 1 : -1;
                while (r > 0 && r < Limits.MaxRows - 1 && Sheet.HiddenRows.Contains(r)) r += step;
            }
            if (Sheet != null && dc != 0)
            {
                int step = dc > 0 ? 1 : -1;
                while (c > 0 && c < Limits.MaxCols - 1 && Sheet.HiddenCols.Contains(c)) c += step;
            }
            if (Sheet != null)
            {
                MergeRange m = Sheet.GetMergeAt(r, c);
                if (m != null) { r = m.Row1; c = m.Col1; }
            }
            if (r == _activeRow && c == _activeCol) return;
            _activeRow = r; _activeCol = c;
            if (extend)
            {
                _selTop = Math.Min(_anchorRow, _activeRow);
                _selBottom = Math.Max(_anchorRow, _activeRow);
                _selLeft = Math.Min(_anchorCol, _activeCol);
                _selRight = Math.Max(_anchorCol, _activeCol);
            }
            else
            {
                _anchorRow = _activeRow; _anchorCol = _activeCol;
                _selTop = _activeRow; _selBottom = _activeRow;
                _selLeft = _activeCol; _selRight = _activeCol;
                MergeRange m = Sheet.GetMergeAt(_activeRow, _activeCol);
                if (m != null)
                {
                    _selTop = m.Row1; _selBottom = m.Row2;
                    _selLeft = m.Col1; _selRight = m.Col2;
                }
            }
            ExpandTo(_activeRow, _activeCol);
            if (ensure) EnsureVisible(_activeRow, _activeCol);
            Invalidate();
            RaiseSelectionChanged();
        }

        protected override bool IsInputKey(Keys keyData)
        {
            Keys k = keyData & Keys.KeyCode;
            if (k == Keys.Tab || k == Keys.Enter || k == Keys.Left || k == Keys.Right || k == Keys.Up ||
                k == Keys.Down || k == Keys.Home || k == Keys.End || k == Keys.PageUp || k == Keys.PageDown)
                return true;
            return base.IsInputKey(keyData);
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            if (IsEditing) return;
            bool shift = e.Shift;
            bool ctrl = e.Control;
            if (ReadOnly && (e.KeyCode == Keys.Delete || e.KeyCode == Keys.Back ||
                e.KeyCode == Keys.F2 || (ctrl && (e.KeyCode == Keys.D || e.KeyCode == Keys.R))))
            {
                e.Handled = true;
                return;
            }
            switch (e.KeyCode)
            {
                case Keys.F4:
                    if (ValidationAt(_activeRow, _activeCol) != null)
                    {
                        ShowValidationList();
                        e.Handled = true;
                    }
                    break;
                case Keys.Left:
                    if (ctrl) MoveActive(0, -FindEdgeCol(-1), shift, true);
                    else MoveActive(0, -1, shift, true);
                    e.Handled = true; break;
                case Keys.Right:
                    if (ctrl) MoveActive(0, FindEdgeCol(1), shift, true);
                    else MoveActive(0, 1, shift, true);
                    e.Handled = true; break;
                case Keys.Up:
                    if (ctrl) MoveActive(-FindEdgeRow(-1), 0, shift, true);
                    else MoveActive(-1, 0, shift, true);
                    e.Handled = true; break;
                case Keys.Down:
                    if (e.Alt && ValidationAt(_activeRow, _activeCol) != null)
                    {
                        ShowValidationList();
                        e.Handled = true;
                    }
                    else if (ctrl) MoveActive(FindEdgeRow(1), 0, shift, true);
                    else MoveActive(1, 0, shift, true);
                    e.Handled = true; break;
                case Keys.Tab:
                    MoveActive(0, shift ? -1 : 1, false, true);
                    e.Handled = true; break;
                case Keys.Enter:
                    MoveActive(shift ? -1 : 1, 0, false, true);
                    e.Handled = true; break;
                case Keys.Home:
                    if (ctrl) SelectCell(0, 0);
                    else MoveActive(0, -_activeCol, shift, true);
                    e.Handled = true; break;
                case Keys.End:
                    if (ctrl) SelectCell(Math.Max(0, Sheet.UsedMaxRow), Math.Max(0, Sheet.UsedMaxCol));
                    else MoveActive(0, _colCount - 1 - _activeCol, shift, true);
                    e.Handled = true; break;
                case Keys.PageUp:
                    MoveActive(-Math.Max(1, (_viewHeight - _headerHeight) / _rowHeight - 1), 0, shift, true);
                    e.Handled = true; break;
                case Keys.PageDown:
                    MoveActive(Math.Max(1, (_viewHeight - _headerHeight) / _rowHeight - 1), 0, shift, true);
                    e.Handled = true; break;
                case Keys.Delete:
                    ClearSelectionContents();
                    e.Handled = true; break;
                case Keys.Back:
                    BeginEdit("");
                    e.Handled = true; break;
                case Keys.F2:
                    BeginEdit(null);
                    e.Handled = true; break;
                case Keys.A:
                    if (ctrl)
                    {
                        if (Sheet.UsedMaxRow >= 0)
                            SelectRange(0, 0, Sheet.UsedMaxRow, Sheet.UsedMaxCol);
                        else SelectRange(0, 0, 99, 25);
                        e.Handled = true;
                    }
                    break;
                case Keys.D:
                    if (ctrl) { FillDown(); e.Handled = true; }
                    break;
                case Keys.R:
                    if (ctrl) { FillRight(); e.Handled = true; }
                    break;
                case Keys.Escape:
                    _clipboard = null;
                    _blink.Stop();
                    Invalidate();
                    break;
            }
        }

        protected override void OnKeyPress(KeyPressEventArgs e)
        {
            base.OnKeyPress(e);
            if (IsEditing) return;
            if (!char.IsControl(e.KeyChar))
            {
                BeginEdit(e.KeyChar.ToString());
                e.Handled = true;
            }
        }

        private int FindEdgeRow(int dir)
        {
            int r = _activeRow;
            int c = _activeCol;
            if (dir > 0)
            {
                bool started = HasCell(r, c);
                while (r < Limits.MaxRows - 1)
                {
                    bool next = HasCell(r + 1, c);
                    if (started && !next) break;
                    if (!started && next) { started = true; }
                    r++;
                    if (started && !HasCell(r, c)) break;
                }
            }
            else
            {
                bool started = HasCell(r, c);
                while (r > 0)
                {
                    bool next = HasCell(r - 1, c);
                    if (started && !next) break;
                    if (!started && next) started = true;
                    r--;
                    if (started && !HasCell(r, c)) break;
                }
            }
            return Math.Abs(r - _activeRow);
        }

        private int FindEdgeCol(int dir)
        {
            int r = _activeRow;
            int c = _activeCol;
            if (dir > 0)
            {
                bool started = HasCell(r, c);
                while (c < Limits.MaxCols - 1)
                {
                    bool next = HasCell(r, c + 1);
                    if (started && !next) break;
                    if (!started && next) started = true;
                    c++;
                    if (started && !HasCell(r, c)) break;
                }
            }
            else
            {
                bool started = HasCell(r, c);
                while (c > 0)
                {
                    bool next = HasCell(r, c - 1);
                    if (started && !next) break;
                    if (!started && next) started = true;
                    c--;
                    if (started && !HasCell(r, c)) break;
                }
            }
            return Math.Abs(c - _activeCol);
        }

        private bool HasCell(int row, int col)
        {
            Cell c = Sheet == null ? null : Sheet.GetCell(row, col);
            return c != null && c.Value.Kind != ValueKind.Empty;
        }

        private void OnBeforeChange()
        {
            if (BeforeChange != null) BeforeChange(this, EventArgs.Empty);
        }

        private void OnChanged()
        {
            if (Changed != null) Changed(this, EventArgs.Empty);
            Invalidate();
        }

        public void BeginEdit(string initial)
        {
            if (_editing || Sheet == null) return;
            if (ReadOnly)
            {
                MessageBox.Show(this, "La hoja está protegida. Quita la protección en Excel para editarla.",
                    "Mini Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            EnsureVisible(_activeRow, _activeCol);
            Rectangle r = GetDisplayRect(_activeRow, _activeCol);
            if (r.Width < 6 || r.Height < 6) return;
            Cell cell = Sheet.GetCell(_activeRow, _activeCol);
            CellStyle st = cell == null ? CellStyle.Default : cell.EffectiveStyle;
            _editRow = _activeRow; _editCol = _activeCol;
            _editOriginal = initial != null ? initial : (cell == null || cell.Raw == null ? "" : cell.Raw);
            if (_editor == null)
            {
                _editor = new CellEditor();
                _editor.Owner = this;
                _editor.BorderStyle = BorderStyle.None;
                _editor.Multiline = false;
                _editor.AutoSize = false;
                _editor.WordWrap = false;
                _editor.BackColor = Color.White;
                _editor.ForeColor = st.FontColor;
                _editor.Visible = false;
                _editor.TextChanged += delegate
                {
                    if (_editing && EditTextChanged != null) EditTextChanged(_editor.Text);
                };
                _editor.Leave += delegate
                {
                    if (_editing && !_closingEditor)
                    {
                        if (!CommitEdit())
                        {
                            BeginInvoke(new MethodInvoker(delegate { if (_editing) { _editor.Visible = true; _editor.Focus(); } }));
                        }
                    }
                };
                Controls.Add(_editor);
            }
            _editor.Font = GetFont(st);
            _editor.ForeColor = st.FontColor;
            _editor.Bounds = new Rectangle(r.X + 1, r.Y + 1, Math.Max(4, r.Width - 2), Math.Max(4, r.Height - 2));
            _editor.Text = _editOriginal;
            _editing = true;
            _editor.Visible = true;
            _editor.BringToFront();
            _editor.Focus();
            _editor.SelectionStart = _editor.TextLength;
            Invalidate();
            if (EditStateChanged != null) EditStateChanged(this, EventArgs.Empty);
            if (EditTextChanged != null) EditTextChanged(_editor.Text);
        }

        private bool EditorKeyDown(KeyEventArgs e)
        {
            if (!_editing) return false;
            if (e.KeyCode == Keys.Escape)
            {
                CancelEdit();
                return true;
            }
            if (e.KeyCode == Keys.Enter)
            {
                if (e.Control)
                {
                    if (CommitEdit()) Focus();
                    return true;
                }
                int dr = e.Shift ? -1 : 1;
                if (CommitAndMove(dr, 0)) return true;
                return true;
            }
            if (e.KeyCode == Keys.Tab)
            {
                int dc = e.Shift ? -1 : 1;
                CommitAndMove(0, dc);
                return true;
            }
            return false;
        }

        public bool CommitEdit()
        {
            if (!_editing || _editor == null) return true;
            string text = _editor.Text;
            if (text.Length > 32767)
            {
                MessageBox.Show(this, "El contenido de una celda no puede superar los 32.767 caracteres.", "Mini Excel",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                if (_editor != null) _editor.Focus();
                return false;
            }
            if (text.StartsWith("=", StringComparison.Ordinal) && FormulaValidator != null)
            {
                string error = FormulaValidator(text);
                if (!string.IsNullOrEmpty(error))
                {
                    MessageBox.Show(this, "La fórmula no es válida:\n\n" + error, "Mini Excel",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    if (_editor != null) _editor.Focus();
                    return false;
                }
            }
            DataValidationRule rule = ValidationAt(_editRow, _editCol);
            if (rule != null)
            {
                bool blank = string.IsNullOrEmpty(text);
                bool found = false;
                foreach (string item in rule.Items)
                    if (string.Equals(item, text, StringComparison.CurrentCultureIgnoreCase)) { found = true; break; }
                if ((!blank || !rule.AllowBlank) && !found)
                {
                    MessageBox.Show(this, rule.Message, "Mini Excel",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    if (_editor != null) _editor.Focus();
                    return false;
                }
            }
            if (string.Equals(text, _editOriginal, StringComparison.Ordinal))
            {
                CloseEditor();
                return true;
            }
            OnBeforeChange();
            try
            {
                Sheet.SetInput(_editRow, _editCol, text);
                if (Sheet.UsedMaxRow >= _editRow || Sheet.UsedMaxCol >= _editCol)
                    ExpandTo(Sheet.UsedMaxRow + 1, Sheet.UsedMaxCol + 1);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "No se pudo guardar la celda:\n" + ex.Message, "Mini Excel",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            CloseEditor();
            OnChanged();
            RaiseSelectionChanged();
            return true;
        }

        private bool CommitAndMove(int dr, int dc)
        {
            if (!CommitEdit()) return false;
            MoveActive(dr, dc, false, true);
            return true;
        }

        private void CloseEditor()
        {
            bool hadFocus = _editor != null && _editor.Focused;
            _closingEditor = true;
            try
            {
                _editing = false;
                if (_editor != null) _editor.Visible = false;
            }
            finally
            {
                _closingEditor = false;
            }
            if (hadFocus) Focus();
            Invalidate();
            if (EditStateChanged != null) EditStateChanged(this, EventArgs.Empty);
        }

        public void CancelEdit()
        {
            if (!_editing) return;
            bool hadFocus = _editor != null && _editor.Focused;
            _closingEditor = true;
            try
            {
                _editing = false;
                if (_editor != null) _editor.Visible = false;
            }
            finally
            {
                _closingEditor = false;
            }
            if (hadFocus) Focus();
            Invalidate();
            if (EditStateChanged != null) EditStateChanged(this, EventArgs.Empty);
            RaiseSelectionChanged();
        }

        public void BeginExternalEdit()
        {
            OnBeforeChange();
        }

        private bool EnsureWritable()
        {
            if (!ReadOnly) return true;
            MessageBox.Show(this, "La hoja está protegida. Quita la protección para editarla.",
                "Mini Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return false;
        }

        public bool ApplyCellText(string text, out string error)
        {
            error = null;
            if (Sheet == null) return false;
            if (ReadOnly) { error = "La hoja está protegida."; return false; }
            if (text == null) text = "";
            if (text.Length > 32767) { error = "El texto es demasiado largo."; return false; }
            if (text.StartsWith("=", StringComparison.Ordinal) && FormulaValidator != null)
            {
                error = FormulaValidator(text);
                if (!string.IsNullOrEmpty(error)) return false;
            }
            Cell cell = Sheet.GetCell(_activeRow, _activeCol);
            string old = cell == null || cell.Raw == null ? "" : cell.Raw;
            if (old == text) return true;
            OnBeforeChange();
            Sheet.SetInput(_activeRow, _activeCol, text);
            ExpandTo(_activeRow, _activeCol);
            OnChanged();
            RaiseSelectionChanged();
            return true;
        }

        public CellStyle GetActiveStyle()
        {
            if (Sheet == null) return new CellStyle();
            Cell c = Sheet.GetCell(_activeRow, _activeCol);
            return c == null ? new CellStyle() : c.EffectiveStyle.Clone();
        }

        public SelectionStats GetSelectionStats()
        {
            if (Sheet == null) return new SelectionStats();
            return SheetStats.Compute(Sheet, _selTop, _selLeft, _selBottom, _selRight);
        }

        public void ApplyStyle(Action<CellStyle> action)
        {
            if (Sheet == null || action == null) return;
            if (ReadOnly) return;
            long area = (long)(_selBottom - _selTop + 1) * (_selRight - _selLeft + 1);
            if (area > 200000)
            {
                MessageBox.Show(this, "La selección es demasiado grande para aplicar formato.", "Mini Excel",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            OnBeforeChange();
            for (int r = _selTop; r <= _selBottom; r++)
            {
                for (int c = _selLeft; c <= _selRight; c++)
                {
                    Cell cell = Sheet.GetCell(r, c);
                    CellStyle st = cell == null ? new CellStyle() : cell.EffectiveStyle.Clone();
                    action(st);
                    if (st.IsDefault()) Sheet.SetStyle(r, c, null);
                    else Sheet.SetStyle(r, c, st);
                }
            }
            OnChanged();
        }

        public void ClearFormatting()
        {
            if (!EnsureWritable()) return;
            if (Sheet == null) return;
            if ((long)(_selBottom - _selTop + 1) * (_selRight - _selLeft + 1) > 200000)
            {
                MessageBox.Show(this, "La selección es demasiado grande para quitar formato.", "Mini Excel",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            OnBeforeChange();
            for (int r = _selTop; r <= _selBottom; r++)
                for (int c = _selLeft; c <= _selRight; c++)
                    Sheet.SetStyle(r, c, null);
            OnChanged();
        }

        public void MergeSelection()
        {
            if (!EnsureWritable()) return;
            if (Sheet == null) return;
            if (_selBottom <= _selTop && _selRight <= _selLeft) return;
            if ((long)(_selBottom - _selTop + 1) * (_selRight - _selLeft + 1) > 200000)
            {
                MessageBox.Show(this, "La selección es demasiado grande para combinar.", "Mini Excel",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            OnBeforeChange();
            for (int r = _selTop; r <= _selBottom; r++)
                for (int c = _selLeft; c <= _selRight; c++)
                    if (r != _selTop || c != _selLeft) Sheet.ClearContents(r, c);
            Sheet.AddMerge(new MergeRange(_selTop, _selLeft, _selBottom, _selRight));
            _activeRow = _selTop; _activeCol = _selLeft;
            _anchorRow = _activeRow; _anchorCol = _activeCol;
            OnChanged();
            RaiseSelectionChanged();
        }

        public void UnmergeSelection()
        {
            if (!EnsureWritable()) return;
            if (Sheet == null) return;
            OnBeforeChange();
            List<MergeRange> remove = new List<MergeRange>();
            foreach (MergeRange m in Sheet.Merges)
                if (m.Intersects(_selTop, _selLeft, _selBottom, _selRight)) remove.Add(m);
            foreach (MergeRange m in remove) Sheet.Merges.Remove(m);
            if (remove.Count > 0)
            {
                Sheet.RebuildMergeLookup();
                Sheet.MergesDirty = true;
            }
            _activeRow = Math.Min(_activeRow, _selTop);
            _activeCol = Math.Min(_activeCol, _selLeft);
            OnChanged();
            RaiseSelectionChanged();
        }

        public void ClearSelectionContents()
        {
            if (!EnsureWritable()) return;
            if (Sheet == null) return;
            if (IsEditing) CancelEdit();
            OnBeforeChange();
            Sheet.ClearRange(_selTop, _selLeft, _selBottom, _selRight);
            Sheet.RecalculateUsedRange();
            OnChanged();
        }

        private void OpenHyperlink(string target, bool internalLink)
        {
            if (string.IsNullOrEmpty(target)) return;
            try
            {
                if (internalLink || target.StartsWith("#", StringComparison.Ordinal))
                {
                    string inner = target.TrimStart('#');
                    int bang = inner.LastIndexOf('!');
                    if (bang > 0)
                    {
                        string sheetName = inner.Substring(0, bang).Trim('\'').Replace("''", "'");
                        string addr = inner.Substring(bang + 1);
                        Sheet targetSheet = Workbook == null ? null : Workbook.FindSheet(sheetName);
                        int r, c, r2, c2;
                        if (targetSheet != null && AddressUtil.TryParseRange(addr, out r, out c, out r2, out c2))
                        {
                            SetSheet(targetSheet);
                            SelectRange(r, c, r2, c2);
                            return;
                        }
                    }
                }
                Uri uri;
                if (!Uri.TryCreate(target, UriKind.Absolute, out uri) ||
                    (uri.Scheme != "http" && uri.Scheme != "https" && uri.Scheme != "mailto"))
                {
                    MessageBox.Show(this,
                        "Por seguridad, Mini Excel solo abre vínculos http, https y mailto.",
                        "Mini Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(uri.AbsoluteUri)
                {
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "No se pudo abrir el vínculo:\n" + ex.Message, "Mini Excel",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        public void CopySelection(bool cut)
        {
            if (Sheet == null) return;
            if (cut && !EnsureWritable()) return;
            if (IsEditing && !CommitEdit()) return;
            int rows = _selBottom - _selTop + 1;
            int cols = _selRight - _selLeft + 1;
            if ((long)rows * cols > 2000000)
            {
                MessageBox.Show(this, "La selección es demasiado grande para copiarla.", "Mini Excel",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            ClipboardBlock block = new ClipboardBlock();
            block.Rows = rows; block.Cols = cols;
            block.Raw = new string[rows, cols];
            block.Styles = new CellStyle[rows, cols];
            block.Values = new CellValue[rows, cols];
            block.Formulas = new bool[rows, cols];
            block.Cut = cut;
            block.SourceSheet = Sheet;
            block.R1 = _selTop; block.C1 = _selLeft;
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    Cell cell = Sheet.GetCell(_selTop + r, _selLeft + c);
                    string raw = cell == null || cell.Raw == null ? "" : cell.Raw;
                    block.Raw[r, c] = raw;
                    block.Values[r, c] = cell == null ? CellValue.Empty : cell.Value;
                    block.Formulas[r, c] = cell != null && cell.IsFormula;
                    block.Styles[r, c] = cell == null || cell.Style == null ? null : cell.Style.Clone();
                    string clip = raw;
                    if (cell != null && !cell.IsFormula && cell.Value.Kind == ValueKind.Number)
                        clip = cell.Value.Number.ToString("G17", System.Globalization.CultureInfo.InvariantCulture);
                    if (c > 0) sb.Append('\t');
                    sb.Append(QuoteTsv(clip));
                }
                sb.Append("\r\n");
            }
            block.Text = sb.ToString();
            _clipboard = block;
            try { Clipboard.SetText(block.Text); } catch { }
            _blinkOn = true;
            _blink.Start();
            Invalidate();
        }

        private static string QuoteTsv(string s)
        {
            if (s.IndexOf('\t') >= 0 || s.IndexOf('\r') >= 0 || s.IndexOf('\n') >= 0 || s.IndexOf('"') >= 0)
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }

        public void PasteClipboard()
        {
            if (!EnsureWritable()) return;
            if (Sheet == null) return;
            if (IsEditing && !CommitEdit()) return;
            string text = null;
            try { text = Clipboard.GetText(); } catch { }
            bool useInternal = _clipboard != null && (string.IsNullOrEmpty(text) || text == _clipboard.Text);
            if (useInternal)
            {
                PasteBlock(_clipboard);
                return;
            }
            if (text == null) return;
            List<List<string>> rows = ParseTsv(text);
            if (rows.Count == 0) return;
            int maxCols = 0;
            foreach (List<string> row in rows) if (row.Count > maxCols) maxCols = row.Count;
            if (rows.Count * maxCols > 500000)
            {
                MessageBox.Show(this, "El portapapeles contiene demasiadas celdas.", "Mini Excel",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            OnBeforeChange();
            for (int r = 0; r < rows.Count; r++)
            {
                for (int c = 0; c < rows[r].Count; c++)
                {
                    int tr = _activeRow + r;
                    int tc = _activeCol + c;
                    if (tr >= Limits.MaxRows || tc >= Limits.MaxCols) continue;
                    Sheet.SetInput(tr, tc, rows[r][c]);
                }
            }
            int bottom = Math.Min(Limits.MaxRows - 1, _activeRow + rows.Count - 1);
            int right = Math.Min(Limits.MaxCols - 1, _activeCol + maxCols - 1);
            _selTop = _activeRow; _selLeft = _activeCol; _selBottom = bottom; _selRight = right;
            _anchorRow = _activeRow; _anchorCol = _activeCol;
            ExpandTo(bottom, right);
            EnsureVisible(_activeRow, _activeCol);
            OnChanged();
            RaiseSelectionChanged();
        }

        private void PasteBlock(ClipboardBlock block)
        {
            if (block == null) return;
            int dr = _activeRow - block.R1;
            int dc = _activeCol - block.C1;
            Sheet sourceSheet = block.SourceSheet ?? Sheet;
            if (block.Cut && sourceSheet == Sheet && dr == 0 && dc == 0)
            {
                _clipboard = null;
                _blink.Stop();
                Invalidate();
                return;
            }
            OnBeforeChange();
            if (block.Cut)
            {
                for (int r = 0; r < block.Rows; r++)
                    for (int c = 0; c < block.Cols; c++)
                    {
                        int sr = block.R1 + r;
                        int sc = block.C1 + c;
                        Cell src = sourceSheet.GetCell(sr, sc);
                        if (src != null) { src.ClearContents(); src.Style = null; }
                        sourceSheet.Cells.Remove(Sheet.Key(sr, sc));
                    }
                int cutBottom = block.R1 + block.Rows - 1;
                int cutRight = block.C1 + block.Cols - 1;
                sourceSheet.Merges.RemoveAll(delegate(MergeRange m)
                {
                    return m.Intersects(block.R1, block.C1, cutBottom, cutRight);
                });
                sourceSheet.RebuildMergeLookup();
                sourceSheet.RecalculateUsedRange();
            }
            int targetBottom = _activeRow;
            int targetRight = _activeCol;
            for (int r = 0; r < block.Rows; r++)
            {
                for (int c = 0; c < block.Cols; c++)
                {
                    int tr = _activeRow + r;
                    int tc = _activeCol + c;
                    if (tr >= Limits.MaxRows || tc >= Limits.MaxCols) continue;
                    string raw = block.Raw[r, c];
                    bool formula = block.Formulas != null && block.Formulas[r, c];
                    CellValue value = block.Values == null ? CellValue.Empty : block.Values[r, c];
                    if (formula)
                    {
                        if (!block.Cut && !string.IsNullOrEmpty(raw) && raw[0] == '=')
                            raw = FormulaEngine.TranslateFormula(raw, tr - (block.R1 + r), tc - (block.C1 + c));
                        Sheet.SetInput(tr, tc, raw);
                    }
                    else if (value.Kind == ValueKind.Number) Sheet.SetNumber(tr, tc, value.Number);
                    else if (value.Kind == ValueKind.Text) Sheet.SetTextLiteral(tr, tc, value.Text);
                    else if (value.Kind == ValueKind.Bool) Sheet.SetInput(tr, tc, value.Bool ? "VERDADERO" : "FALSO");
                    else if (value.Kind == ValueKind.Error) Sheet.SetError(tr, tc, value.Error);
                    else Sheet.ClearContents(tr, tc);
                    if (block.Styles[r, c] != null) Sheet.SetStyle(tr, tc, block.Styles[r, c].Clone());
                    if (tr > targetBottom) targetBottom = tr;
                    if (tc > targetRight) targetRight = tc;
                }
            }
            _selTop = _activeRow; _selLeft = _activeCol;
            _selBottom = targetBottom; _selRight = targetRight;
            _clipboard = null;
            _blink.Stop();
            ExpandTo(targetBottom, targetRight);
            EnsureVisible(_activeRow, _activeCol);
            OnChanged();
            RaiseSelectionChanged();
        }

        private static List<List<string>> ParseTsv(string text)
        {
            List<List<string>> rows = new List<List<string>>();
            List<string> row = new List<string>();
            System.Text.StringBuilder field = new System.Text.StringBuilder();
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
                    else if (ch == '\t') { row.Add(field.ToString()); field.Length = 0; }
                    else if (ch == '\r')
                    {
                        if (i + 1 < text.Length && text[i + 1] == '\n') i++;
                        row.Add(field.ToString()); field.Length = 0;
                        rows.Add(row); row = new List<string>();
                    }
                    else if (ch == '\n')
                    {
                        row.Add(field.ToString()); field.Length = 0;
                        rows.Add(row); row = new List<string>();
                    }
                    else field.Append(ch);
                }
            }
            if (field.Length > 0 || row.Count > 0) { row.Add(field.ToString()); rows.Add(row); }
            return rows;
        }

        public void FillDown()
        {
            if (!EnsureWritable()) return;
            if (Sheet == null) return;
            if (_selBottom <= _selTop)
            {
                if (_activeRow <= 0) return;
                SelectRange(_activeRow - 1, _activeCol, _activeRow, _activeCol);
            }
            if ((long)(_selBottom - _selTop + 1) * (_selRight - _selLeft + 1) > 200000)
            {
                MessageBox.Show(this, "La selección es demasiado grande para rellenar.", "Mini Excel",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            OnBeforeChange();
            for (int r = _selTop + 1; r <= _selBottom; r++)
                for (int c = _selLeft; c <= _selRight; c++)
                    CopyCell(_selTop, c, r, c, r - _selTop, 0);
            OnChanged();
        }

        public void FillRight()
        {
            if (!EnsureWritable()) return;
            if (Sheet == null) return;
            if (_selRight <= _selLeft)
            {
                if (_activeCol <= 0) return;
                SelectRange(_activeRow, _activeCol - 1, _activeRow, _activeCol);
            }
            if ((long)(_selBottom - _selTop + 1) * (_selRight - _selLeft + 1) > 200000)
            {
                MessageBox.Show(this, "La selección es demasiado grande para rellenar.", "Mini Excel",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            OnBeforeChange();
            for (int c = _selLeft + 1; c <= _selRight; c++)
                for (int r = _selTop; r <= _selBottom; r++)
                    CopyCell(r, _selLeft, r, c, 0, c - _selLeft);
            OnChanged();
        }

        private void CopyCell(int srcRow, int srcCol, int dstRow, int dstCol, int dRow, int dCol)
        {
            Cell src = Sheet.GetCell(srcRow, srcCol);
            if (src == null)
            {
                Sheet.ClearContents(dstRow, dstCol);
                return;
            }
            if (src.IsFormula)
            {
                string formula = src.Raw;
                if (!string.IsNullOrEmpty(formula) && formula[0] == '=')
                    formula = FormulaEngine.TranslateFormula(formula, dRow, dCol);
                Sheet.SetInput(dstRow, dstCol, formula);
            }
            else if (src.Value.Kind == ValueKind.Number) Sheet.SetNumber(dstRow, dstCol, src.Value.Number);
            else if (src.Value.Kind == ValueKind.Text) Sheet.SetTextLiteral(dstRow, dstCol, src.Value.Text);
            else if (src.Value.Kind == ValueKind.Bool) Sheet.SetInput(dstRow, dstCol, src.Value.Bool ? "VERDADERO" : "FALSO");
            else if (src.Value.Kind == ValueKind.Error) Sheet.SetError(dstRow, dstCol, src.Value.Error);
            else Sheet.ClearContents(dstRow, dstCol);
            if (src.Style != null) Sheet.SetStyle(dstRow, dstCol, src.Style.Clone());
        }

        public void AutoSum()
        {
            if (!EnsureWritable()) return;
            if (Sheet == null) return;
            int selRows = _selBottom - _selTop + 1;
            int selCols = _selRight - _selLeft + 1;
            if (selRows > 1 && selCols == 1)
            {
                int target = _selBottom + 1;
                if (target < Limits.MaxRows)
                {
                    OnBeforeChange();
                    Sheet.SetInput(target, _selLeft, "=SUMA(" +
                        AddressUtil.CellName(_selTop, _selLeft) + ":" +
                        AddressUtil.CellName(_selBottom, _selLeft) + ")");
                    _activeRow = target; _activeCol = _selLeft;
                    _selTop = _selBottom = target;
                    _selLeft = _selRight = _selLeft;
                    _anchorRow = target; _anchorCol = _selLeft;
                    OnChanged();
                    RaiseSelectionChanged();
                    return;
                }
            }
            if (selCols > 1 && selRows == 1)
            {
                int target = _selRight + 1;
                if (target < Limits.MaxCols)
                {
                    OnBeforeChange();
                    Sheet.SetInput(_selTop, target, "=SUMA(" +
                        AddressUtil.CellName(_selTop, _selLeft) + ":" +
                        AddressUtil.CellName(_selTop, _selRight) + ")");
                    _activeRow = _selTop; _activeCol = target;
                    _selTop = _selBottom = _selTop;
                    _selLeft = _selRight = target;
                    _anchorRow = _selTop; _anchorCol = target;
                    OnChanged();
                    RaiseSelectionChanged();
                    return;
                }
            }
            int row = _activeRow;
            int col = _activeCol;
            int first = -1, last = -1;
            for (int r = row - 1; r >= 0; r--)
            {
                Cell c = Sheet.GetCell(r, col);
                if (c == null || c.Value.Kind == ValueKind.Empty) break;
                if (c.Value.Kind == ValueKind.Number || c.IsFormula)
                {
                    first = r;
                    last = row - 1;
                    // Un subtotal anterior marca el inicio del bloque: no se
                    // deben volver a sumar las operaciones que ya contiene.
                    if (c.IsFormula && LooksLikeTotalFormula(c.Formula)) break;
                }
                else break;
            }
            if (first < 0)
            {
                for (int c = col - 1; c >= 0; c--)
                {
                    Cell cell = Sheet.GetCell(row, c);
                    if (cell == null || cell.Value.Kind == ValueKind.Empty) break;
                    if (cell.Value.Kind == ValueKind.Number || cell.IsFormula)
                    {
                        first = c;
                        last = col - 1;
                        if (cell.IsFormula && LooksLikeTotalFormula(cell.Formula)) break;
                    }
                    else break;
                }
                if (first >= 0)
                {
                    string formula = "=SUMA(" + AddressUtil.CellName(row, first) + ":" + AddressUtil.CellName(row, last) + ")";
                    ApplyCellToActive(formula);
                    return;
                }
                return;
            }
            string f = "=SUMA(" + AddressUtil.CellName(first, col) + ":" + AddressUtil.CellName(last, col) + ")";
            ApplyCellToActive(f);
        }

        private static bool LooksLikeTotalFormula(string formula)
        {
            if (string.IsNullOrEmpty(formula)) return false;
            string f = formula.Trim();
            if (f.StartsWith("=", StringComparison.Ordinal)) f = f.Substring(1).TrimStart();
            int paren = f.IndexOf('(');
            if (paren <= 0) return false;
            string name = f.Substring(0, paren).Trim();
            return string.Equals(name, "SUMA", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(name, "SUM", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(name, "SUBTOTAL", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(name, "SUBTOTALES", StringComparison.OrdinalIgnoreCase);
        }

        private void ApplyCellToActive(string formula)
        {
            OnBeforeChange();
            Sheet.SetInput(_activeRow, _activeCol, formula);
            OnChanged();
            RaiseSelectionChanged();
        }

        public void SortSelection(bool ascending)
        {
            if (!EnsureWritable()) return;
            if (Sheet == null) return;
            int top, left, bottom, right;
            if (_selBottom > _selTop)
            {
                top = _selTop; left = _selLeft; bottom = _selBottom; right = _selRight;
            }
            else
            {
                FindCurrentRegion(out top, out left, out bottom, out right);
            }
            if (bottom - top < 1) return;
            if (bottom - top > 10000)
            {
                MessageBox.Show(this, "El rango es demasiado grande para ordenar.", "Mini Excel",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            List<RowData> rows = new List<RowData>();
            for (int r = top; r <= bottom; r++)
            {
                RowData rd = new RowData();
                rd.Cells = new Cell[right - left + 1];
                rd.Key = CellValue.Empty;
                for (int c = left; c <= right; c++)
                {
                    Cell cell = Sheet.GetCell(r, c);
                    rd.Cells[c - left] = cell == null ? null : cell.Clone();
                    if (c == left && cell != null) rd.Key = cell.Value;
                }
                rows.Add(rd);
            }
            rows.Sort(delegate(RowData a, RowData b) { return CompareRows(a.Key, b.Key, ascending); });
            OnBeforeChange();
            for (int r = top; r <= bottom; r++)
                for (int c = left; c <= right; c++)
                    Sheet.Cells.Remove(Sheet.Key(r, c));
            for (int i = 0; i < rows.Count; i++)
            {
                int r = top + i;
                for (int j = 0; j < rows[i].Cells.Length; j++)
                {
                    Cell cell = rows[i].Cells[j];
                    if (cell == null) continue;
                    Sheet.Cells[Sheet.Key(r, left + j)] = cell;
                }
            }
            Sheet.RecalculateUsedRange();
            OnChanged();
            RaiseSelectionChanged();
        }

        private sealed class RowData
        {
            public Cell[] Cells;
            public CellValue Key;
        }

        private static int CompareRows(CellValue a, CellValue b, bool ascending)
        {
            if (a.Kind == ValueKind.Empty && b.Kind == ValueKind.Empty) return 0;
            if (a.Kind == ValueKind.Empty) return 1;
            if (b.Kind == ValueKind.Empty) return -1;
            int cmp;
            if (a.Kind == ValueKind.Number && b.Kind == ValueKind.Number) cmp = a.Number.CompareTo(b.Number);
            else if (a.Kind == ValueKind.Text && b.Kind == ValueKind.Text)
                cmp = string.Compare(a.Text ?? "", b.Text ?? "", StringComparison.CurrentCultureIgnoreCase);
            else if (a.Kind == ValueKind.Number) cmp = -1;
            else if (b.Kind == ValueKind.Number) cmp = 1;
            else cmp = string.Compare(a.ToDisplayString(), b.ToDisplayString(), StringComparison.CurrentCultureIgnoreCase);
            return ascending ? cmp : -cmp;
        }

        private void FindCurrentRegion(out int top, out int left, out int bottom, out int right)
        {
            top = _activeRow; bottom = _activeRow; left = _activeCol; right = _activeCol;
            int maxRow = Math.Max(Sheet.UsedMaxRow, _activeRow);
            int maxCol = Math.Max(Sheet.UsedMaxCol, _activeCol);
            while (top > 0 && RowHasContent(top - 1, left, right)) top--;
            while (bottom < maxRow && RowHasContent(bottom + 1, left, right)) bottom++;
            while (left > 0 && ColHasContent(left - 1, top, bottom)) left--;
            while (right < maxCol && ColHasContent(right + 1, top, bottom)) right++;
            if (right == left)
            {
                int extra = Math.Min(maxCol, left + 8);
                right = extra;
            }
        }

        private bool RowHasContent(int row, int c1, int c2)
        {
            for (int c = c1; c <= c2; c++)
                if (HasCell(row, c)) return true;
            return false;
        }

        private bool ColHasContent(int col, int r1, int r2)
        {
            for (int r = r1; r <= r2; r++)
                if (HasCell(r, col)) return true;
            return false;
        }

        public void InsertRow()
        {
            InsertRows(_selTop, _selBottom - _selTop + 1);
        }

        public void InsertColumn()
        {
            InsertCols(_selLeft, _selRight - _selLeft + 1);
        }

        public void DeleteRow()
        {
            DeleteRows(_selTop, _selBottom - _selTop + 1);
        }

        public void DeleteColumn()
        {
            DeleteCols(_selLeft, _selRight - _selLeft + 1);
        }

        private void InsertRows(int at, int count)
        {
            if (!EnsureWritable()) return;
            if (Sheet == null || count < 1) return;
            if (Sheet.UsedMaxRow + count >= Limits.MaxRows - 1)
            {
                MessageBox.Show(this, "No hay espacio para insertar más filas.", "Mini Excel");
                return;
            }
            OnBeforeChange();
            Sheet.StructureDirty = true;
            Sheet.RowsDirty = true;
            Sheet.MergesDirty = true;
            ShiftRowMetrics(at, count, true);
            Sheet.ShiftMergesForInsert(true, at, count);
            Sheet.ShiftStructuredRangesForInsert(true, at, count);
            AdjustFormulas(true, at, count, true);
            List<KeyValuePair<long, Cell>> move = new List<KeyValuePair<long, Cell>>();
            foreach (KeyValuePair<long, Cell> kv in Sheet.Cells)
            {
                int row = (int)(kv.Key >> 14);
                if (row >= at) move.Add(kv);
            }
            foreach (KeyValuePair<long, Cell> kv in move) Sheet.Cells.Remove(kv.Key);
            foreach (KeyValuePair<long, Cell> kv in move)
            {
                int row = (int)(kv.Key >> 14);
                int col = (int)(kv.Key & 0x3FFF);
                Sheet.Cells[Sheet.Key(row + count, col)] = kv.Value;
            }
            Sheet.RecalculateUsedRange();
            _activeRow = at; _activeCol = Math.Min(_activeCol, Limits.MaxCols - 1);
            _selTop = _selBottom = at; _selLeft = _selRight = _activeCol;
            _anchorRow = at; _anchorCol = _activeCol;
            OnChanged();
            RaiseSelectionChanged();
        }

        private void InsertCols(int at, int count)
        {
            if (!EnsureWritable()) return;
            if (Sheet == null || count < 1) return;
            if (Sheet.UsedMaxCol + count >= Limits.MaxCols - 1)
            {
                MessageBox.Show(this, "No hay espacio para insertar más columnas.", "Mini Excel");
                return;
            }
            OnBeforeChange();
            Sheet.StructureDirty = true;
            Sheet.WidthsDirty = true;
            Sheet.MergesDirty = true;
            ShiftColumnMetrics(at, count, true);
            Sheet.ShiftMergesForInsert(false, at, count);
            Sheet.ShiftStructuredRangesForInsert(false, at, count);
            AdjustFormulas(false, at, count, true);
            List<KeyValuePair<long, Cell>> move = new List<KeyValuePair<long, Cell>>();
            foreach (KeyValuePair<long, Cell> kv in Sheet.Cells)
            {
                int col = (int)(kv.Key & 0x3FFF);
                if (col >= at) move.Add(kv);
            }
            foreach (KeyValuePair<long, Cell> kv in move) Sheet.Cells.Remove(kv.Key);
            foreach (KeyValuePair<long, Cell> kv in move)
            {
                int row = (int)(kv.Key >> 14);
                int col = (int)(kv.Key & 0x3FFF);
                Sheet.Cells[Sheet.Key(row, col + count)] = kv.Value;
            }
            Sheet.RecalculateUsedRange();
            _activeCol = at;
            _selTop = _selBottom = _activeRow; _selLeft = _selRight = at;
            _anchorRow = _activeRow; _anchorCol = at;
            OnChanged();
            RaiseSelectionChanged();
        }

        private void DeleteRows(int at, int count)
        {
            if (!EnsureWritable()) return;
            if (Sheet == null || count < 1) return;
            if (count > Limits.MaxRows - at) count = Limits.MaxRows - at;
            OnBeforeChange();
            Sheet.StructureDirty = true;
            Sheet.RowsDirty = true;
            Sheet.MergesDirty = true;
            ShiftRowMetrics(at, count, false);
            Sheet.ShiftMergesForDelete(true, at, count);
            Sheet.ShiftStructuredRangesForDelete(true, at, count);
            AdjustFormulas(true, at, count, false);
            List<KeyValuePair<long, Cell>> keep = new List<KeyValuePair<long, Cell>>();
            foreach (KeyValuePair<long, Cell> kv in Sheet.Cells)
            {
                int row = (int)(kv.Key >> 14);
                if (row < at) keep.Add(kv);
                else if (row >= at + count)
                {
                    int col = (int)(kv.Key & 0x3FFF);
                    keep.Add(new KeyValuePair<long, Cell>(Sheet.Key(row - count, col), kv.Value));
                }
            }
            Sheet.Cells.Clear();
            foreach (KeyValuePair<long, Cell> kv in keep) Sheet.Cells[kv.Key] = kv.Value;
            Sheet.RecalculateUsedRange();
            _activeRow = Math.Min(at, Math.Max(0, Sheet.UsedMaxRow));
            _selTop = _selBottom = _activeRow; _selLeft = _selRight = _activeCol;
            _anchorRow = _activeRow; _anchorCol = _activeCol;
            OnChanged();
            RaiseSelectionChanged();
        }

        private void DeleteCols(int at, int count)
        {
            if (!EnsureWritable()) return;
            if (Sheet == null || count < 1) return;
            if (count > Limits.MaxCols - at) count = Limits.MaxCols - at;
            OnBeforeChange();
            Sheet.StructureDirty = true;
            Sheet.WidthsDirty = true;
            Sheet.MergesDirty = true;
            ShiftColumnMetrics(at, count, false);
            Sheet.ShiftMergesForDelete(false, at, count);
            Sheet.ShiftStructuredRangesForDelete(false, at, count);
            AdjustFormulas(false, at, count, false);
            List<KeyValuePair<long, Cell>> keep = new List<KeyValuePair<long, Cell>>();
            foreach (KeyValuePair<long, Cell> kv in Sheet.Cells)
            {
                int col = (int)(kv.Key & 0x3FFF);
                if (col < at) keep.Add(kv);
                else if (col >= at + count)
                {
                    int row = (int)(kv.Key >> 14);
                    keep.Add(new KeyValuePair<long, Cell>(Sheet.Key(row, col - count), kv.Value));
                }
            }
            Sheet.Cells.Clear();
            foreach (KeyValuePair<long, Cell> kv in keep) Sheet.Cells[kv.Key] = kv.Value;
            Sheet.RecalculateUsedRange();
            _activeCol = Math.Min(at, Math.Max(0, Sheet.UsedMaxCol));
            _selTop = _selBottom = _activeRow; _selLeft = _selRight = _activeCol;
            _anchorRow = _activeRow; _anchorCol = _activeCol;
            OnChanged();
            RaiseSelectionChanged();
        }

        private void ShiftRowMetrics(int at, int count, bool insert)
        {
            Dictionary<int, float> heights = new Dictionary<int, float>(Sheet.RowHeights);
            Sheet.RowHeights.Clear();
            foreach (KeyValuePair<int, float> kv in heights)
            {
                int r = kv.Key;
                if (insert)
                {
                    if (r >= at) r += count;
                    Sheet.RowHeights[r] = kv.Value;
                }
                else
                {
                    if (r >= at && r < at + count) continue;
                    if (r >= at + count) r -= count;
                    Sheet.RowHeights[r] = kv.Value;
                }
            }
            HashSet<int> hidden = new HashSet<int>(Sheet.HiddenRows);
            Sheet.HiddenRows.Clear();
            foreach (int r in hidden)
            {
                int nr = r;
                if (insert)
                {
                    if (r >= at) nr = r + count;
                    Sheet.HiddenRows.Add(nr);
                }
                else
                {
                    if (r >= at && r < at + count) continue;
                    if (r >= at + count) nr = r - count;
                    Sheet.HiddenRows.Add(nr);
                }
            }
        }

        private void ShiftColumnMetrics(int at, int count, bool insert)
        {
            Dictionary<int, float> widths = new Dictionary<int, float>(Sheet.ColWidths);
            Sheet.ColWidths.Clear();
            foreach (KeyValuePair<int, float> kv in widths)
            {
                int c = kv.Key;
                if (insert)
                {
                    if (c >= at) c += count;
                    Sheet.ColWidths[c] = kv.Value;
                }
                else
                {
                    if (c >= at && c < at + count) continue;
                    if (c >= at + count) c -= count;
                    Sheet.ColWidths[c] = kv.Value;
                }
            }
            HashSet<int> hidden = new HashSet<int>(Sheet.HiddenCols);
            Sheet.HiddenCols.Clear();
            foreach (int c0 in hidden)
            {
                int c = c0;
                if (insert)
                {
                    if (c0 >= at) c = c0 + count;
                    Sheet.HiddenCols.Add(c);
                }
                else
                {
                    if (c0 >= at && c0 < at + count) continue;
                    if (c0 >= at + count) c = c0 - count;
                    Sheet.HiddenCols.Add(c);
                }
            }
        }

        private void AdjustFormulas(bool rows, int at, int count, bool insert)
        {
            foreach (Sheet sheet in Workbook.Sheets)
            {
                foreach (Cell cell in sheet.Cells.Values)
                {
                    if (!cell.IsFormula) continue;
                    string f = FormulaEngine.ShiftFormula(cell.Formula, sheet.Name, Sheet.Name, insert, rows, at, count);
                    if (!string.Equals(f, cell.Formula, StringComparison.Ordinal))
                    {
                        cell.Formula = f;
                        cell.Raw = f;
                        cell.Ast = null;
                        cell.AstReady = false;
                        cell.FormulaDirty = true;
                        cell.Dirty = true;
                    }
                }
            }
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            Focus();
            if (_editing)
            {
                if (!CommitEdit()) return;
            }
            if (e.Button != MouseButtons.Left && e.Button != MouseButtons.Right) return;
            if (e.Button == MouseButtons.Left)
            {
                string slicerValue;
                SheetSlicer slicer = SlicerItemAt(e.Location, out slicerValue);
                if (slicer != null)
                {
                    if (!EnsureWritable()) return;
                    PivotRange pivot;
                    int field;
                    SlicerItems(slicer, out pivot, out field);
                    slicer.SelectedValue = slicerValue;
                    if (pivot != null && field >= 0)
                    {
                        OnBeforeChange();
                        PivotEngine.RefreshFiltered(Sheet, pivot, field, slicerValue);
                    }
                    Invalidate();
                    return;
                }
                SheetShape shape = ShapeAt(e.Location);
                if (shape != null && (!string.IsNullOrEmpty(shape.Hyperlink) ||
                    !string.IsNullOrEmpty(shape.HyperlinkInternal)))
                {
                    string target = !string.IsNullOrEmpty(shape.HyperlinkInternal)
                        ? shape.HyperlinkInternal : shape.Hyperlink;
                    OpenHyperlink(target, !string.IsNullOrEmpty(shape.HyperlinkInternal));
                    return;
                }
                bool resize;
                SheetChart hit = ChartAt(e.Location, out resize);
                if (hit != null)
                {
                    _selectedChart = hit;
                    _chartDrag = resize ? ChartDragMode.Resize : ChartDragMode.Move;
                    _chartDragStart = e.Location;
                    _chartPushed = false;
                    Rectangle cr;
                    TryChartRect(hit, out cr);
                    _chartStartRect = cr;
                    _chartStartWidth = Math.Max(40, cr.Width);
                    _chartStartHeight = Math.Max(30, cr.Height);
                    Capture = true;
                    Cursor = resize ? Cursors.SizeNWSE : Cursors.SizeAll;
                    Invalidate();
                    return;
                }
                bool imageResize;
                SheetImage image = ImageAt(e.Location, out imageResize);
                if (image != null)
                {
                    _selectedImage = image;
                    _selectedChart = null;
                    _imageDrag = true;
                    _imageResize = imageResize;
                    _imageDragStart = e.Location;
                    TryImageRect(image, out _imageStartRect);
                    _imageStartFromColOff = image.FromColOff;
                    _imageStartFromRowOff = image.FromRowOff;
                    _imageStartToColOff = image.ToColOff;
                    _imageStartToRowOff = image.ToRowOff;
                    _imageStartWidthPx = image.WidthPx;
                    _imageStartHeightPx = image.HeightPx;
                    _imagePushed = false;
                    Capture = true;
                    Cursor = imageResize ? Cursors.SizeNWSE : Cursors.SizeAll;
                    Invalidate();
                    return;
                }
                if (_selectedChart != null)
                {
                    _selectedChart = null;
                    Invalidate();
                }
                if (_selectedImage != null)
                {
                    _selectedImage = null;
                    Invalidate();
                }
            }
            if (e.Button == MouseButtons.Left)
            {
                for (int fcol = 0; fcol < _colCount; fcol++)
                {
                    Rectangle filterArrow = FilterArrowRect(fcol);
                    if (filterArrow.IsEmpty || !filterArrow.Contains(e.Location)) continue;
                    ShowFilterMenu(fcol);
                    return;
                }
            }
            if (e.Button == MouseButtons.Left && (ModifierKeys & Keys.Control) != 0)
            {
                int hr, hc;
                if (TryPointToCell(e.Location, out hr, out hc))
                {
                    Cell hcCell = Sheet.GetCell(hr, hc);
                    if (hcCell != null && !string.IsNullOrEmpty(hcCell.Hyperlink))
                    {
                        OpenHyperlink(hcCell.Hyperlink, hcCell.HyperlinkInternal);
                        return;
                    }
                }
            }
            if (e.Button == MouseButtons.Right)
            {
                int rr, cc;
                if (TryPointToCell(e.Location, out rr, out cc))
                {
                    if (rr < _selTop || rr > _selBottom || cc < _selLeft || cc > _selRight)
                        SelectCell(rr, cc);
                }
                if (ContextMenuRequested != null) ContextMenuRequested(this, EventArgs.Empty);
                return;
            }

            if (InFillHandle(e.Location))
            {
                _drag = DragMode.Fill;
                _fillTargetRow = _selBottom;
                _fillTargetCol = _selRight;
                _fillVertical = true;
                Capture = true;
                return;
            }

            int borderCol = HitColumnBorder(e.Location);
            if (borderCol >= 0)
            {
                _drag = DragMode.ResizeColumn;
                _resizeCol = borderCol;
                _resizeStartX = e.X;
                _resizeStartWidth = Sheet.GetColWidth(borderCol);
                _resizePushed = false;
                Capture = true;
                Cursor = Cursors.SizeWE;
                return;
            }
            int borderRow = HitRowBorder(e.Location);
            if (borderRow >= 0)
            {
                _drag = DragMode.ResizeRow;
                _resizeRow = borderRow;
                _resizeStartY = e.Y;
                _resizeStartHeight = Sheet.GetRowHeightPoints(borderRow);
                _resizePushed = false;
                Capture = true;
                Cursor = Cursors.SizeNS;
                return;
            }

            if (e.Y <= _headerHeight && e.X > _headerWidth)
            {
                int cx = _scrollX + e.X - _headerWidth;
                int col = ColAtX(cx);
                if (col >= 0)
                {
                    _drag = DragMode.Columns;
                    _anchorCol = col;
                    int bottom = Math.Max(Sheet.UsedMaxRow, Math.Min(Limits.MaxRows - 1, 199));
                    SelectRange(0, col, bottom, col);
                    Capture = true;
                }
                return;
            }
            if (e.X <= _headerWidth && e.Y > _headerHeight)
            {
                int row = RowAtY(_scrollY + e.Y - _headerHeight);
                _drag = DragMode.Rows;
                _anchorRow = row;
                int right = Math.Max(Sheet.UsedMaxCol, Math.Min(Limits.MaxCols - 1, 49));
                SelectRange(row, 0, row, right);
                Capture = true;
                return;
            }
            if (e.X <= _headerWidth && e.Y <= _headerHeight)
            {
                if (Sheet.UsedMaxRow >= 0) SelectRange(0, 0, Sheet.UsedMaxRow, Sheet.UsedMaxCol);
                else SelectRange(0, 0, 99, 25);
                return;
            }

            int r, c;
            if (TryPointToCell(e.Location, out r, out c))
            {
                if ((ModifierKeys & Keys.Shift) != 0)
                {
                    MergeRange shiftMerge = Sheet.GetMergeAt(r, c);
                    if (shiftMerge != null) { r = shiftMerge.Row1; c = shiftMerge.Col1; }
                    _activeRow = r; _activeCol = c;
                    _selTop = Math.Min(_anchorRow, _activeRow);
                    _selBottom = Math.Max(_anchorRow, _activeRow);
                    _selLeft = Math.Min(_anchorCol, _activeCol);
                    _selRight = Math.Max(_anchorCol, _activeCol);
                    Invalidate();
                    RaiseSelectionChanged();
                }
                else
                {
                    SelectCell(r, c);
                }
                _drag = DragMode.Cells;
                Capture = true;
            }
        }

        private bool InFillHandle(Point p)
        {
            if (_selBottom < _selTop || _selRight < _selLeft) return false;
            Rectangle r = CellRect(_selBottom, _selRight);
            if (r.IsEmpty) return false;
            Rectangle h = new Rectangle(r.Right - 7, r.Bottom - 7, 8, 8);
            return h.Contains(p);
        }

        private int HitColumnBorder(Point p)
        {
            if (p.Y > _headerHeight || p.X <= _headerWidth) return -1;
            int cx = _scrollX + p.X - _headerWidth;
            int col = ColAtX(cx);
            if (col < 0 || col >= _colCount - 1) return -1;
            int borderX = _headerWidth + _colLeft[col + 1] - _scrollX;
            if (Math.Abs(p.X - borderX) <= 4) return col;
            return -1;
        }

        private int HitRowBorder(Point p)
        {
            if (p.X > _headerWidth || p.Y <= _headerHeight) return -1;
            int contentY = _scrollY + p.Y - _headerHeight;
            if (FreezeRows > 0 && p.Y - _headerHeight < _rowTop[FreezeRows])
                contentY = p.Y - _headerHeight;
            int row = RowAtY(contentY);
            if (row < 0 || row >= _rowCount) return -1;
            int sy = FreezeRows > 0 && row < FreezeRows ? 0 : _scrollY;
            int topY = _headerHeight + _rowTop[row] - sy;
            if (row > 0 && Math.Abs(p.Y - topY) <= 4) return row - 1;
            if (row >= _rowCount - 1) return -1;
            int borderY = _headerHeight + _rowTop[row + 1] - sy;
            if (Math.Abs(p.Y - borderY) <= 4) return row;
            return -1;
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            _lastMouse = e.Location;
            UpdateCommentTooltip(e.Location);
            if (_imageDrag && _selectedImage != null)
            {
                if (!_imagePushed) { OnBeforeChange(); _imagePushed = true; }
                float z = Math.Max(0.1f, _zoom * _uiScale);
                int dx = e.X - _imageDragStart.X;
                int dy = e.Y - _imageDragStart.Y;
                if (_imageResize)
                {
                    _selectedImage.ToRow = -1;
                    _selectedImage.ToCol = -1;
                    _selectedImage.WidthPx = Math.Max(20,
                        (int)(Math.Max(20, _imageStartWidthPx + dx / z)));
                    _selectedImage.HeightPx = Math.Max(20,
                        (int)(Math.Max(20, _imageStartHeightPx + dy / z)));
                }
                else
                {
                    long ex = (long)(dx / z * 9525);
                    long ey = (long)(dy / z * 9525);
                    _selectedImage.FromColOff = _imageStartFromColOff + ex;
                    _selectedImage.FromRowOff = _imageStartFromRowOff + ey;
                    if (_selectedImage.ToRow >= 0 && _selectedImage.ToCol >= 0)
                    {
                        _selectedImage.ToColOff = _imageStartToColOff + ex;
                        _selectedImage.ToRowOff = _imageStartToRowOff + ey;
                    }
                }
                _selectedImage.Dirty = true;
                Sheet.ImagesDirty = true;
                Invalidate();
                return;
            }
            if (_drag == DragMode.ResizeRow)
            {
                int dy = e.Y - _resizeStartY;
                if (Math.Abs(dy) > 1)
                {
                    if (!_resizePushed)
                    {
                        OnBeforeChange();
                        _resizePushed = true;
                    }
                    float points = _resizeStartHeight + dy * 3f / (4f * _zoom * _uiScale);
                    Sheet.SetRowHeightPoints(_resizeRow, points);
                    RebuildMetrics();
                    Invalidate();
                }
                return;
            }
            if (_chartDrag != ChartDragMode.None && _selectedChart != null)
            {
                if (!_chartPushed)
                {
                    OnBeforeChange();
                    _chartPushed = true;
                }
                float z = Math.Max(0.1f, _zoom * _uiScale);
                int dx = e.X - _chartDragStart.X;
                int dy = e.Y - _chartDragStart.Y;
                if (_chartDrag == ChartDragMode.Move)
                {
                    int targetX = _chartStartRect.X + dx;
                    int targetY = _chartStartRect.Y + dy;
                    int contentX = _scrollX + targetX - _headerWidth;
                    int contentY = _scrollY + targetY - _headerHeight;
                    int col = ColAtX(contentX);
                    int row = RowAtY(contentY);
                    if (col < 0) col = 0;
                    if (row < 0) row = 0;
                    _selectedChart.FromCol = col;
                    _selectedChart.FromRow = row;
                    _selectedChart.FromColOff = (long)((contentX - _colLeft[col]) / z * 9525);
                    _selectedChart.FromRowOff = (long)((contentY - _rowTop[row]) / z * 9525);
                    _selectedChart.ToRow = -1;
                    _selectedChart.ToCol = -1;
                    _selectedChart.WidthPx = Math.Max(60, (int)(_chartStartWidth / z));
                    _selectedChart.HeightPx = Math.Max(40, (int)(_chartStartHeight / z));
                }
                else
                {
                    _selectedChart.ToRow = -1;
                    _selectedChart.ToCol = -1;
                    _selectedChart.WidthPx = Math.Max(60, (int)(Math.Max(60, _chartStartWidth + dx) / z));
                    _selectedChart.HeightPx = Math.Max(40, (int)(Math.Max(40, _chartStartHeight + dy) / z));
                }
                _selectedChart.Dirty = true;
                Sheet.ChartsDirty = true;
                Invalidate();
                return;
            }
            if (_drag == DragMode.ResizeColumn)
            {
                int dx = e.X - _resizeStartX;
                if (Math.Abs(dx) > 1)
                {
                    if (!_resizePushed)
                    {
                        OnBeforeChange();
                        _resizePushed = true;
                    }
                    float w = _resizeStartWidth + dx / (_zoom * _uiScale);
                    Sheet.SetColWidth(_resizeCol, w);
                    RebuildMetrics();
                    Invalidate();
                }
                return;
            }
            if (_drag == DragMode.Fill)
            {
                int rr, cc;
                PointToCellClamped(e.Location, out rr, out cc);
                int dRow = rr - _selBottom;
                int dCol = cc - _selRight;
                if (dRow <= 0 && dCol <= 0) { _fillTargetRow = _selBottom; _fillTargetCol = _selRight; Invalidate(); return; }
                if (Math.Abs(dRow) >= Math.Abs(dCol) && dRow > 0)
                {
                    _fillVertical = true;
                    _fillTargetRow = rr;
                    _fillTargetCol = _selRight;
                }
                else if (dCol > 0)
                {
                    _fillVertical = false;
                    _fillTargetRow = _selBottom;
                    _fillTargetCol = cc;
                }
                Invalidate();
                return;
            }
            if (_drag == DragMode.Cells || _drag == DragMode.Columns || _drag == DragMode.Rows)
            {
                int r, c;
                PointToCellClamped(e.Location, out r, out c);
                if (_drag == DragMode.Columns)
                {
                    _selLeft = Math.Min(_anchorCol, c);
                    _selRight = Math.Max(_anchorCol, c);
                    _activeCol = c;
                }
                else if (_drag == DragMode.Rows)
                {
                    _selTop = Math.Min(_anchorRow, r);
                    _selBottom = Math.Max(_anchorRow, r);
                    _activeRow = r;
                }
                else
                {
                    MergeRange targetMerge = Sheet.GetMergeAt(r, c);
                    if (targetMerge != null)
                    {
                        // Arrastrar dentro de la misma combinación no debe
                        // cambiar la celda activa a una celda interior.
                        if (_anchorRow == targetMerge.Row1 && _anchorCol == targetMerge.Col1)
                        {
                            r = _anchorRow;
                            c = _anchorCol;
                        }
                        else
                        {
                            r = targetMerge.Row1;
                            c = targetMerge.Col1;
                        }
                    }
                    _activeRow = r; _activeCol = c;
                    _selTop = Math.Min(_anchorRow, r);
                    _selBottom = Math.Max(_anchorRow, r);
                    _selLeft = Math.Min(_anchorCol, c);
                    _selRight = Math.Max(_anchorCol, c);
                    if (targetMerge != null && _anchorRow == targetMerge.Row1 &&
                        _anchorCol == targetMerge.Col1)
                    {
                        _selTop = targetMerge.Row1;
                        _selBottom = targetMerge.Row2;
                        _selLeft = targetMerge.Col1;
                        _selRight = targetMerge.Col2;
                    }
                    if (_activeRow < _anchorRow || _activeCol < _anchorCol)
                    {
                        _selTop = Math.Min(_anchorRow, r);
                        _selBottom = Math.Max(_anchorRow, r);
                    }
                }
                ExpandTo(_selBottom, _selRight);
                Invalidate();
                RaiseSelectionChanged();
                if (e.X < 0 || e.Y < 0 || e.X >= _viewWidth || e.Y >= _viewHeight || e.Y > _viewHeight - 20)
                    _autoScroll.Start();
                else _autoScroll.Stop();
                return;
            }
            bool chartResize;
            SheetChart hoverChart = ChartAt(e.Location, out chartResize);
            if (hoverChart != null) Cursor = chartResize ? Cursors.SizeNWSE : Cursors.SizeAll;
            else if (e.Y <= _headerHeight && HitColumnBorder(e.Location) >= 0) Cursor = Cursors.SizeWE;
            else if (HitRowBorder(e.Location) >= 0) Cursor = Cursors.SizeNS;
            else if (InFillHandle(e.Location)) Cursor = Cursors.Cross;
            else Cursor = Cursors.Default;
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            _autoScroll.Stop();
            if (_imageDrag)
            {
                if (_imagePushed) OnChanged();
                _imageDrag = false;
                _imagePushed = false;
                Capture = false;
                Cursor = Cursors.Default;
                Invalidate();
                return;
            }
            if (_chartDrag != ChartDragMode.None)
            {
                if (_chartPushed) OnChanged();
                _chartDrag = ChartDragMode.None;
                _chartPushed = false;
                Capture = false;
                Cursor = Cursors.Default;
                Invalidate();
                return;
            }
            if (_drag == DragMode.Fill) ApplyFill();
            if ((_drag == DragMode.ResizeColumn || _drag == DragMode.ResizeRow) && _resizePushed)
            {
                OnChanged();
            }
            _drag = DragMode.None;
            Capture = false;
            Cursor = Cursors.Default;
            Invalidate();
        }

        private void ApplyFill()
        {
            if (Sheet == null) return;
            if (_fillTargetRow <= _selBottom && _fillTargetCol <= _selRight) return;
            long targetRows = Math.Max(_selBottom, _fillTargetRow) - _selTop + 1;
            long targetCols = Math.Max(_selRight, _fillTargetCol) - _selLeft + 1;
            if (targetRows * targetCols > 200000)
            {
                MessageBox.Show(this, "El área de relleno es demasiado grande.", "Mini Excel",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            OnBeforeChange();
            if (_fillVertical)
            {
                int h = _selBottom - _selTop + 1;
                for (int r = _selBottom + 1; r <= _fillTargetRow; r++)
                {
                    int srcRow = _selTop + ((r - _selTop) % h);
                    for (int c = _selLeft; c <= _selRight; c++)
                        CopyCell(srcRow, c, r, c, r - srcRow, 0);
                }
                _selBottom = _fillTargetRow;
            }
            else
            {
                int w = _selRight - _selLeft + 1;
                for (int c = _selRight + 1; c <= _fillTargetCol; c++)
                {
                    int srcCol = _selLeft + ((c - _selLeft) % w);
                    for (int r = _selTop; r <= _selBottom; r++)
                        CopyCell(r, srcCol, r, c, 0, c - srcCol);
                }
                _selRight = _fillTargetCol;
            }
            OnChanged();
            RaiseSelectionChanged();
        }

        protected override void OnMouseDoubleClick(MouseEventArgs e)
        {
            base.OnMouseDoubleClick(e);
            int borderCol = HitColumnBorder(e.Location);
            if (borderCol >= 0)
            {
                AutoFitColumn(borderCol);
                return;
            }
            int r, c;
            if (TryPointToCell(e.Location, out r, out c))
            {
                SelectCell(r, c);
                BeginEdit(null);
            }
        }

        private void AutoFitColumn(int col)
        {
            int max = 40;
            using (Graphics g = CreateGraphics())
            {
                for (int r = 0; r <= Math.Min(Sheet.UsedMaxRow, 5000); r++)
                {
                    Cell cell = Sheet.GetCell(r, col);
                    if (cell == null || cell.Value.Kind == ValueKind.Empty) continue;
                    string text = ShowFormulas && cell.IsFormula ? cell.Raw : NumberFormatter.Format(cell.Value, cell.EffectiveStyle);
                    Size sz = TextRenderer.MeasureText(g, text, GetFont(cell.EffectiveStyle));
                    if (sz.Width + 12 > max) max = sz.Width + 12;
                }
            }
            max = Math.Max(36, Math.Min(500, max));
            OnBeforeChange();
            Sheet.SetColWidth(col, max / (_zoom * _uiScale));
            RebuildMetrics();
            OnChanged();
        }

        private void AutoScrollTick()
        {
            if (_drag == DragMode.None) { _autoScroll.Stop(); return; }
            int dx = 0, dy = 0;
            if (_lastMouse.X >= _viewWidth - 10) dx = (int)(20 * _zoom);
            else if (_lastMouse.X <= _headerWidth + 4) dx = -(int)(20 * _zoom);
            if (_lastMouse.Y >= _viewHeight - 10) dy = _rowHeight;
            else if (_lastMouse.Y <= _headerHeight + 4) dy = -_rowHeight;
            if (dx == 0 && dy == 0) { _autoScroll.Stop(); return; }
            SetScroll(_scrollX + dx, _scrollY + dy);
            int r, c;
            PointToCellClamped(new Point(Math.Max(_headerWidth + 2, Math.Min(_viewWidth - 2, _lastMouse.X)),
                                         Math.Max(_headerHeight + 2, Math.Min(_viewHeight - 2, _lastMouse.Y))), out r, out c);
            _activeRow = r; _activeCol = c;
            _selTop = Math.Min(_anchorRow, r);
            _selBottom = Math.Max(_anchorRow, r);
            _selLeft = Math.Min(_anchorCol, c);
            _selRight = Math.Max(_anchorCol, c);
            Invalidate();
            RaiseSelectionChanged();
        }

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            base.OnMouseWheel(e);
            if ((ModifierKeys & Keys.Control) != 0)
            {
                float z = _zoom + (e.Delta > 0 ? 0.1f : -0.1f);
                SetZoom(z);
                return;
            }
            int lines = SystemInformation.MouseWheelScrollLines;
            if (lines <= 0) lines = 3;
            int dy = (e.Delta / 120) * lines * _rowHeight;
            if ((ModifierKeys & Keys.Shift) != 0)
            {
                SetScroll(_scrollX - dy, _scrollY);
            }
            else
            {
                SetScroll(_scrollX, _scrollY - dy);
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.None;
            g.PixelOffsetMode = PixelOffsetMode.Half;
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            using (SolidBrush b = new SolidBrush(Color.White)) g.FillRectangle(b, ClientRectangle);
            if (Sheet == null) return;

            Rectangle viewport = new Rectangle(_headerWidth, _headerHeight,
                Math.Max(0, _viewWidth - _headerWidth), Math.Max(0, _viewHeight - _headerHeight));
            if (viewport.Width <= 0 || viewport.Height <= 0) return;

            int firstRow = RowAtY(_scrollY);
            int firstCol = ColAtX(_scrollX);
            if (firstCol < 0) firstCol = 0;
            int maxRow = firstRow;
            while (maxRow < _rowCount - 1 &&
                   _rowTop[maxRow + 1] - _scrollY + _headerHeight < _viewHeight)
                maxRow++;
            maxRow = Math.Min(_rowCount - 1, maxRow + 2);
            int maxCol = Math.Min(_colCount - 1, firstCol + 80);

            int selRows = _selBottom - _selTop + 1;
            int selCols = _selRight - _selLeft + 1;

            // Las líneas se dibujan antes de los fondos: así una celda con relleno
            // (o una combinación) tapa las líneas interiores, como en Excel.
            if (ShowGridLines && Sheet.ShowGridLines)
            {
                using (Pen p = new Pen(Theme.GridLine, 1f))
                {
                    for (int r = firstRow; r <= maxRow + 1 && r <= _rowCount; r++)
                    {
                        if (r < _rowCount && _rowTop[r + 1] == _rowTop[r]) continue;
                        int y = _headerHeight + _rowTop[r] - _scrollY;
                        if (y >= _headerHeight && y <= _viewHeight) g.DrawLine(p, _headerWidth, y, _viewWidth, y);
                    }
                    for (int c = firstCol; c <= maxCol + 1 && c <= _colCount; c++)
                    {
                        if (c < _colCount && _colLeft[c + 1] == _colLeft[c]) continue;
                        int x = _headerWidth + _colLeft[c] - _scrollX;
                        if (x >= _headerWidth && x <= _viewWidth) g.DrawLine(p, x, _headerHeight, x, _viewHeight);
                    }
                }
            }

            // Fondos de celdas
            for (int r = firstRow; r <= maxRow; r++)
            {
                for (int c = firstCol; c <= maxCol; c++)
                {
                    MergeRange mm = Sheet.GetMergeAt(r, c);
                    if (mm != null && (mm.Row1 != r || mm.Col1 != c)) continue;
                    Cell cell = Sheet.GetCell(r, c);
                    Color fill = Color.Empty;
                    if (cell != null && cell.Style != null) fill = cell.Style.FillColor;
                    TableFormat tf = TableFormatting.Find(Sheet, r, c);
                    Color tableFill = tf == null ? Color.Empty : TableFormatting.Fill(tf, r, c);
                    if (!tableFill.IsEmpty) fill = tableFill;
                    PivotRange pivot = PivotFormatting.Find(Sheet, r, c);
                    Color pivotFill = pivot == null ? Color.Empty : PivotFormatting.Fill(pivot, r, c);
                    if (!pivotFill.IsEmpty) fill = pivotFill;
                    ConditionalRule cf = ConditionalEvaluator.Find(Sheet, r, c);
                    if (cf != null)
                    {
                        if (!cf.ScaleFill.IsEmpty) fill = cf.ScaleFill;
                        else if (cf.Style != null && !cf.Style.FillColor.IsEmpty)
                            fill = cf.Style.FillColor;
                    }
                    if (!fill.IsEmpty)
                    {
                        Rectangle cr = mm == null ? CellRect(r, c) : RectForCells(mm.Row1, mm.Col1, mm.Row2, mm.Col2);
                        using (SolidBrush b = new SolidBrush(fill)) g.FillRectangle(b, cr);
                    }
                    if (cell != null && cell.Value.Kind == ValueKind.Number)
                    {
                        foreach (ConditionalFormat cfBar in Sheet.ConditionalFormats)
                        {
                            if (r < cfBar.Row1 || r > cfBar.Row2 || c < cfBar.Col1 || c > cfBar.Col2) continue;
                            foreach (ConditionalRule bar in cfBar.Rules)
                            {
                                if (!bar.IsDataBar) continue;
                                double value = cell.Value.Number;
                                if (value <= 0) break;
                                if (value > 1) value = 1;
                                Rectangle cr = mm == null ? CellRect(r, c) : RectForCells(mm.Row1, mm.Col1, mm.Row2, mm.Col2);
                                int w = Math.Max(2, (int)(cr.Width * value));
                                using (SolidBrush b = new SolidBrush(Color.FromArgb(110, bar.DataBarColor)))
                                    g.FillRectangle(b, cr.X + 1, cr.Y + 2, w - 2, Math.Max(1, cr.Height - 5));
                                break;
                            }
                        }
                    }
                }
            }

            if (selRows > 1 || selCols > 1)
            {
                Rectangle sel = RectForCells(_selTop, _selLeft, _selBottom, _selRight);
                Rectangle clipped = Rectangle.Intersect(sel, viewport);
                if (!clipped.IsEmpty)
                {
                    using (SolidBrush b = new SolidBrush(Theme.SelectionFill)) g.FillRectangle(b, clipped);
                }
                Rectangle ac = GetDisplayRect(_activeRow, _activeCol);
                if (ac.IntersectsWith(viewport))
                {
                    Cell cell = Sheet.GetCell(_activeRow, _activeCol);
                    Color fill = cell == null || cell.Style == null || cell.Style.FillColor.IsEmpty ? Color.White : cell.Style.FillColor;
                    using (SolidBrush b = new SolidBrush(fill))
                        g.FillRectangle(b, Rectangle.Intersect(ac, viewport));
                }
            }

            g.SetClip(viewport);
            for (int r = firstRow; r <= maxRow; r++)
            {
                for (int c = firstCol; c <= maxCol; c++)
                {
                    MergeRange mm = Sheet.GetMergeAt(r, c);
                    if (mm != null && (mm.Row1 != r || mm.Col1 != c)) continue;
                    Rectangle cr = mm == null ? CellRect(r, c) : RectForCells(mm.Row1, mm.Col1, mm.Row2, mm.Col2);
                    if (cr.Width <= 0 || cr.Height <= 0) continue;
                    if (cr.Right < _headerWidth || cr.Left > _viewWidth) continue;
                    DrawCell(g, r, c, cr);
                }
            }
            g.ResetClip();

            // Bordes personales
            g.SetClip(viewport);
            for (int r = firstRow; r <= maxRow; r++)
            {
                for (int c = firstCol; c <= maxCol; c++)
                {
                    Cell cell = Sheet.GetCell(r, c);
                    if (cell == null || cell.Style == null || cell.Style.Borders == BorderFlags.None) continue;
                    ConditionalRule cfBorder = ConditionalEvaluator.Find(Sheet, r, c);
                    if (cfBorder != null &&
                        (!cfBorder.ScaleFill.IsEmpty ||
                         (cfBorder.Style != null && !cfBorder.Style.FillColor.IsEmpty)))
                        continue;
                    MergeRange mm = Sheet.GetMergeAt(r, c);
                    if (mm != null && (mm.Row1 != r || mm.Col1 != c)) continue;
                    Rectangle cr = mm == null ? CellRect(r, c) : RectForCells(mm.Row1, mm.Col1, mm.Row2, mm.Col2);
                    if (cr.Width <= 0 || cr.Height <= 0) continue;
                    using (Pen p = new Pen(Color.FromArgb(110, 110, 110), 1f))
                    {
                        if ((cell.Style.Borders & BorderFlags.Left) != 0) g.DrawLine(p, cr.Left, cr.Top, cr.Left, cr.Bottom - 1);
                        if ((cell.Style.Borders & BorderFlags.Top) != 0) g.DrawLine(p, cr.Left, cr.Top, cr.Right - 1, cr.Top);
                        if ((cell.Style.Borders & BorderFlags.Right) != 0) g.DrawLine(p, cr.Right - 1, cr.Top, cr.Right - 1, cr.Bottom - 1);
                        if ((cell.Style.Borders & BorderFlags.Bottom) != 0) g.DrawLine(p, cr.Left, cr.Bottom - 1, cr.Right - 1, cr.Bottom - 1);
                    }
                }
            }
            g.ResetClip();

            DrawFrozenPanes(g, viewport, firstRow, firstCol, maxRow, maxCol);
            DrawTableBorders(g, viewport);
            DrawImages(g, viewport);
            DrawCharts(g, viewport);
            DrawShapes(g, viewport);
            DrawSlicers(g, viewport);
            DrawChartSelection(g);
            DrawImageSelection(g);
            DrawFilterButtons(g, firstRow, firstCol, maxRow, maxCol);
            DrawHeaders(g, firstRow, firstCol, maxRow, maxCol);
            DrawFrozenHeaders(g, FreezeRows, FreezeCols);

            // Borde de selección
            g.SetClip(viewport);
            Rectangle selRect = RectForCells(_selTop, _selLeft, _selBottom, _selRight);
            selRect.Inflate(-1, -1);
            Rectangle visible = Rectangle.Intersect(selRect, viewport);
            if (!visible.IsEmpty)
            {
                using (Pen p = new Pen(Theme.Green, 2f))
                    g.DrawRectangle(p, selRect.X, selRect.Y, selRect.Width, selRect.Height);
                Rectangle handle = new Rectangle(selRect.Right - 3, selRect.Bottom - 3, 7, 7);
                if (handle.IntersectsWith(viewport))
                {
                    using (SolidBrush b = new SolidBrush(Theme.Green)) g.FillRectangle(b, handle);
                    using (Pen wp = new Pen(Color.White, 1f)) g.DrawRectangle(wp, handle);
                }
            }

            if (_clipboard != null)
            {
                Rectangle copy = RectForCells(_clipboard.R1, _clipboard.C1, _clipboard.R1 + _clipboard.Rows - 1, _clipboard.C1 + _clipboard.Cols - 1);
                if (copy.IntersectsWith(viewport))
                {
                    using (Pen p = new Pen(_blinkOn ? Color.FromArgb(70, 150, 100) : Color.FromArgb(110, 180, 135), 1.5f))
                    {
                        p.DashStyle = DashStyle.Dash;
                        p.DashOffset = _blinkOn ? 0 : 3;
                        g.DrawRectangle(p, copy.X, copy.Y, copy.Width, copy.Height);
                    }
                }
            }

            if (_drag == DragMode.Fill && (_fillTargetRow > _selBottom || _fillTargetCol > _selRight))
            {
                int bottom = Math.Max(_selBottom, _fillTargetRow);
                int right = Math.Max(_selRight, _fillTargetCol);
                Rectangle fr = RectForCells(_selTop, _selLeft, bottom, right);
                using (Pen p = new Pen(Color.FromArgb(90, 120, 100), 1.5f))
                {
                    p.DashStyle = DashStyle.Dash;
                    g.DrawRectangle(p, fr.X, fr.Y, fr.Width, fr.Height);
                }
            }

            if (_editing)
            {
                Rectangle er = GetDisplayRect(_editRow, _editCol);
                using (Pen p = new Pen(Theme.Green, 2f)) g.DrawRectangle(p, er.X, er.Y, er.Width - 1, er.Height - 1);
            }
            g.ResetClip();

            // Esquina inferior de las barras
            using (SolidBrush b = new SolidBrush(Color.FromArgb(243, 243, 243)))
                g.FillRectangle(b, _viewWidth, _viewHeight, ClientSize.Width - _viewWidth, ClientSize.Height - _viewHeight);
        }

        private Rectangle RectForCells(int r1, int c1, int r2, int c2)
        {
            int sx = FreezeCols > 0 && c1 < FreezeCols ? 0 : _scrollX;
            int sy = FreezeRows > 0 && r1 < FreezeRows ? 0 : _scrollY;
            int x = _headerWidth + _colLeft[Math.Max(0, Math.Min(c1, _colCount - 1))] - sx;
            int y = _headerHeight + _rowTop[Math.Max(0, Math.Min(r1, _rowCount - 1))] - sy;
            int x2 = _headerWidth + _colLeft[Math.Max(0, Math.Min(c2 + 1, _colCount))] - _scrollX;
            int y2 = _headerHeight + _rowTop[Math.Max(0, Math.Min(r2 + 1, _rowCount))] - _scrollY;
            return new Rectangle(x, y, x2 - x, y2 - y);
        }

        private bool DrawAccounting(Graphics g, string text, Font font, Color color, Rectangle cr, CellStyle st)
        {
            string code = st.CustomFormat ?? "";
            if (code.IndexOf('*') < 0 && code.IndexOf("_(", StringComparison.Ordinal) < 0) return false;
            string symbol = null;
            int at = -1;
            string[] symbols = new string[] { "$", "\u20AC", "\u00A3", "\u00A5", "S/" };
            foreach (string s in symbols)
            {
                int p = text.IndexOf(s, StringComparison.Ordinal);
                if (p >= 0 && (at < 0 || p < at)) { at = p; symbol = s; }
            }
            if (at < 0) return false;
            string amount = text.Substring(at + symbol.Length).Trim();
            int pad = 3;
            Rectangle sign = new Rectangle(cr.X + pad, cr.Y, 20, cr.Height);
            TextRenderer.DrawText(g, symbol, font, sign, color,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix);
            Rectangle value = new Rectangle(sign.Right + 1, cr.Y, Math.Max(1, cr.Width - (sign.Width + pad * 2 + 1)), cr.Height);
            TextRenderer.DrawText(g, amount, font, value, color,
                TextFormatFlags.Right | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding |
                TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis);
            return true;
        }

        private void DrawCell(Graphics g, int row, int col, Rectangle cr)
        {
            Cell cell = Sheet.GetCell(row, col);
            if (cell == null) return;
            bool isError = false;
            bool forceText = false;
            string text = GetDisplayText(cell, out isError, out forceText);
            if (string.IsNullOrEmpty(text)) return;
            // Excel guarda descripciones de accesibilidad detrás de segmentaciones
            // e imágenes; si no dibujamos esos objetos, mostrarlas como texto
            // produce cajas gigantes que no existen en Excel/LibreOffice.
            if (Sheet.HasSlicerPlaceholders &&
                text.StartsWith("La segmentación de datos", StringComparison.OrdinalIgnoreCase)) return;
            if (Sheet.HasPicturePlaceholders &&
                text.StartsWith("La tira de luces", StringComparison.OrdinalIgnoreCase)) return;
            if (Sheet.HasChartPlaceholders &&
                text.StartsWith("El gráfico de barras", StringComparison.OrdinalIgnoreCase)) return;
            CellStyle st = cell.EffectiveStyle;
            Color cellFill = st.FillColor;
            TableFormat tfCell = TableFormatting.Find(Sheet, row, col);
            if (tfCell != null)
            {
                Color tfFill = TableFormatting.Fill(tfCell, row, col);
                if (!tfFill.IsEmpty) cellFill = tfFill;
            }
            PivotRange pivotCell = PivotFormatting.Find(Sheet, row, col);
            if (pivotCell != null)
            {
                Color pf = PivotFormatting.Fill(pivotCell, row, col);
                if (!pf.IsEmpty) cellFill = pf;
            }
            ConditionalRule cfFill = ConditionalEvaluator.Find(Sheet, row, col);
            if (cfFill != null)
            {
                if (!cfFill.ScaleFill.IsEmpty) cellFill = cfFill.ScaleFill;
                else if (cfFill.Style != null && !cfFill.Style.FillColor.IsEmpty)
                    cellFill = cfFill.Style.FillColor;
            }
            if (!cellFill.IsEmpty)
                using (SolidBrush fb = new SolidBrush(cellFill)) g.FillRectangle(fb, cr);
            if (cell.Value.Kind == ValueKind.Number)
            {
                foreach (ConditionalFormat cfBar in Sheet.ConditionalFormats)
                {
                    if (row < cfBar.Row1 || row > cfBar.Row2 || col < cfBar.Col1 || col > cfBar.Col2) continue;
                    foreach (ConditionalRule bar in cfBar.Rules)
                    {
                        if (!bar.IsDataBar) continue;
                        double value = cell.Value.Number;
                        if (value <= 0) break;
                        if (value > 1) value = 1;
                        int w = Math.Max(2, (int)(cr.Width * value));
                        using (SolidBrush b = new SolidBrush(Color.FromArgb(100, bar.DataBarColor)))
                            g.FillRectangle(b, cr.X + 1, cr.Y + 2, w - 2, Math.Max(1, cr.Height - 5));
                        break;
                    }
                }
            }
            if (!string.IsNullOrEmpty(cell.Comment))
            {
                Point[] triangle = new Point[]
                {
                    new Point(cr.Right - 1, cr.Y + 1),
                    new Point(cr.Right - 7, cr.Y + 1),
                    new Point(cr.Right - 1, cr.Y + 7)
                };
                using (SolidBrush commentBrush = new SolidBrush(Color.FromArgb(192, 0, 0)))
                    g.FillPolygon(commentBrush, triangle);
            }
            if (cfFill != null && !string.IsNullOrEmpty(cfFill.IconGlyph) && cr.Width >= 16)
            {
                using (Font iconFont = new Font("Segoe UI Symbol",
                    Math.Max(7f, 9f * _zoom * _uiScale), FontStyle.Bold))
                using (SolidBrush iconBrush = new SolidBrush(
                    cfFill.IconColor.IsEmpty ? Color.Gray : cfFill.IconColor))
                {
                    SizeF isz = g.MeasureString(cfFill.IconGlyph, iconFont);
                    g.DrawString(cfFill.IconGlyph, iconFont, iconBrush,
                        cr.Right - isz.Width - 2, cr.Y + Math.Max(0, (cr.Height - isz.Height) / 2));
                }
            }
            bool link = !string.IsNullOrEmpty(cell.Hyperlink);
            Font font = GetFont(st, link);
            Color color = isError ? Theme.CellError : (link ? Color.FromArgb(5, 99, 193) : st.FontColor);
            ConditionalRule cfText = ConditionalEvaluator.Find(Sheet, row, col);
            if (cfText != null && cfText.Style != null && !cfText.Style.FontColor.IsEmpty)
                color = cfText.Style.FontColor;
            int align = NumberFormatter.EffectiveAlign(cell.Value, st);
            int pad = 3;
            if (!forceText && cell.Value.Kind == ValueKind.Number && st.Align < 0 &&
                DrawAccounting(g, text, font, color, cr, st)) return;
            Size measured = TextRenderer.MeasureText(g, text, font, new Size(int.MaxValue, int.MaxValue),
                TextFormatFlags.NoPadding | TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix);
            int avail = cr.Width - pad * 2;
            if (align == 0 && measured.Width > avail)
            {
                int endX = cr.Right - 2;
                for (int cc = col + 1; cc < _colCount && cc < col + 30; cc++)
                {
                    Rectangle nr = CellRect(row, cc);
                    if (nr.Left >= _viewWidth) break;
                    Cell neighbor = Sheet.GetCell(row, cc);
                    bool block = neighbor != null && neighbor.Value.Kind != ValueKind.Empty;
                    if (block) break;
                    endX = nr.Right - 2;
                }
                avail = Math.Max(avail, endX - (cr.X + pad));
            }
            bool dateLikeFormat = st.Format == NumberFormatKind.Date ||
                (st.Format == NumberFormatKind.Custom && !string.IsNullOrEmpty(st.CustomFormat) &&
                 (st.CustomFormat.IndexOf('y') >= 0 || st.CustomFormat.IndexOf('d') >= 0 ||
                  st.CustomFormat.IndexOf("mmm", StringComparison.Ordinal) >= 0));
            if (measured.Width > avail && cell.Value.Kind == ValueKind.Number && !forceText &&
                dateLikeFormat)
            {
                // Excel reduce la fecha en lugar de mostrar #### en columnas
                // estrechas; hacemos lo mismo hasta 5 pt.
                float size = font.SizeInPoints;
                for (float factor = 0.9f; factor >= 0.55f; factor -= 0.1f)
                {
                    using (Font small = new Font(font.FontFamily,
                        Math.Max(5f, size * factor), font.Style, GraphicsUnit.Point))
                    {
                        Size m = TextRenderer.MeasureText(g, text, small,
                            new Size(int.MaxValue, int.MaxValue),
                            TextFormatFlags.NoPadding | TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix);
                        if (m.Width > avail) continue;
                        TextFormatFlags df = TextFormatFlags.NoPadding | TextFormatFlags.SingleLine |
                            TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix;
                        if (align == 1) df |= TextFormatFlags.HorizontalCenter;
                        else if (align == 2) df |= TextFormatFlags.Right;
                        TextRenderer.DrawText(g, text, small, cr, color, df);
                        return;
                    }
                }
            }
            if (measured.Width > avail && cell.Value.Kind == ValueKind.Number && !forceText)
            {
                text = "#####";
                measured = TextRenderer.MeasureText(g, text, font);
            }
            if (st.WrapText && !isError)
            {
                Rectangle wrap = new Rectangle(cr.X + pad, cr.Y + 1,
                    Math.Max(1, cr.Width - pad * 2), Math.Max(1, cr.Height - 2));
                TextFormatFlags wf = TextFormatFlags.NoPadding | TextFormatFlags.WordBreak |
                    TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix;
                if (align == 1) wf |= TextFormatFlags.HorizontalCenter;
                else if (align == 2) wf |= TextFormatFlags.Right;
                else wf |= TextFormatFlags.Left;
                TextRenderer.DrawText(g, text, font, wrap, color, wf);
                return;
            }
            Rectangle tr = new Rectangle(cr.X + pad, cr.Y, Math.Max(1, avail), cr.Height);
            TextFormatFlags flags = TextFormatFlags.NoPadding | TextFormatFlags.SingleLine |
                                    TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis;
            if (align == 1) flags |= TextFormatFlags.HorizontalCenter;
            else if (align == 2) flags |= TextFormatFlags.Right;
            else flags |= TextFormatFlags.Left;
            if (align == 0 && measured.Width > cr.Width - pad * 2)
                flags &= ~TextFormatFlags.EndEllipsis;
            TextRenderer.DrawText(g, text, font, tr, color, flags);
        }

        private string GetDisplayText(Cell cell, out bool isError, out bool forceText)
        {
            isError = false;
            forceText = false;
            if (cell == null) return "";
            if (ShowFormulas && cell.IsFormula) { forceText = true; return Builtins.DisplayFormula(cell.Raw ?? ""); }
            CellValue v = cell.Value;
            if (v.Kind == ValueKind.Error) { isError = true; return v.Error; }
            if (v.Kind == ValueKind.Text) { forceText = true; return v.Text; }
            if (v.Kind == ValueKind.Empty) return "";
            return NumberFormatter.Format(v, cell.EffectiveStyle);
        }

        private bool TryImageRect(SheetImage image, out Rectangle rect)
        {
            rect = Rectangle.Empty;
            if (image == null || image.Image == null || Sheet == null ||
                image.FromRow < 0 || image.FromCol < 0 ||
                image.FromRow >= _rowCount || image.FromCol >= _colCount) return false;
            Rectangle from = GetDisplayRect(image.FromRow, image.FromCol);
            float z = _zoom * _uiScale;
            int x = from.Left + (int)(image.FromColOff / 9525.0 * z);
            int y = from.Top + (int)(image.FromRowOff / 9525.0 * z);
            int w, h;
            if (image.ToRow >= 0 && image.ToCol >= 0 &&
                image.ToRow < _rowCount && image.ToCol < _colCount)
            {
                Rectangle to = GetDisplayRect(image.ToRow, image.ToCol);
                w = to.Left + (int)(image.ToColOff / 9525.0 * z) - x;
                h = to.Top + (int)(image.ToRowOff / 9525.0 * z) - y;
            }
            else
            {
                w = (int)(image.WidthPx * z);
                h = (int)(image.HeightPx * z);
            }
            if (w < 4 || h < 4) return false;
            rect = new Rectangle(x, y, w, h);
            return true;
        }

        private void DrawImages(Graphics g, Rectangle viewport)
        {
            if (Sheet == null || Sheet.Images.Count == 0) return;
            foreach (SheetImage image in Sheet.Images)
            {
                Rectangle rect;
                if (!TryImageRect(image, out rect) || !rect.IntersectsWith(viewport)) continue;
                GraphicsState state = g.Save();
                try
                {
                    g.SetClip(viewport);
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    g.DrawImage(image.Image, rect);
                }
                finally { g.Restore(state); }
            }
        }

        private void DrawShapes(Graphics g, Rectangle viewport)
        {
            if (Sheet == null || Sheet.Shapes.Count == 0) return;
            foreach (SheetShape shape in Sheet.Shapes)
            {
                Rectangle rect;
                if (!TryShapeRect(shape, out rect) || !rect.IntersectsWith(viewport)) continue;
                GraphicsState state = g.Save();
                try
                {
                    g.SetClip(viewport);
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                    using (Pen border = new Pen(shape.BorderColor, 1f))
                    using (SolidBrush fill = new SolidBrush(shape.FillColor))
                    {
                        g.FillRectangle(fill, rect);
                        g.DrawRectangle(border, rect);
                    }
                    if (!string.IsNullOrEmpty(shape.Text) && rect.Width > 10 && rect.Height > 8)
                    {
                        using (Font font = new Font("Segoe UI", shape.FontSize * _zoom * _uiScale,
                            FontStyle.Regular))
                        using (SolidBrush text = new SolidBrush(shape.FontColor))
                        using (StringFormat format = new StringFormat())
                        {
                            format.Alignment = StringAlignment.Center;
                            format.LineAlignment = StringAlignment.Center;
                            format.Trimming = StringTrimming.EllipsisCharacter;
                            g.DrawString(shape.Text, font, text, rect, format);
                        }
                    }
                }
                finally { g.Restore(state); }
            }
        }

        private void DrawSlicers(Graphics g, Rectangle viewport)
        {
            if (Sheet == null || Sheet.Slicers.Count == 0) return;
            foreach (SheetSlicer slicer in Sheet.Slicers)
            {
                Rectangle rect;
                if (!TrySlicerRect(slicer, out rect) || !rect.IntersectsWith(viewport)) continue;
                GraphicsState state = g.Save();
                try
                {
                    g.SetClip(viewport);
                    using (SolidBrush panel = new SolidBrush(Color.White))
                        g.FillRectangle(panel, rect);
                    using (Pen border = new Pen(Color.FromArgb(166, 166, 166)))
                        g.DrawRectangle(border, rect);
                    int titleH = Math.Min(24, Math.Max(18, rect.Height / 6));
                    using (SolidBrush title = new SolidBrush(Color.FromArgb(68, 114, 196)))
                        g.FillRectangle(title, rect.Left + 1, rect.Top + 1, rect.Width - 2, titleH - 1);
                    using (Font font = new Font("Segoe UI", Math.Max(8f, 9f * _zoom * _uiScale),
                        FontStyle.Bold))
                    using (SolidBrush text = new SolidBrush(Color.White))
                    using (StringFormat sf = new StringFormat())
                    {
                        sf.Alignment = StringAlignment.Near;
                        sf.LineAlignment = StringAlignment.Center;
                        sf.Trimming = StringTrimming.EllipsisCharacter;
                        Rectangle titleRect = new Rectangle(rect.Left + 4, rect.Top + 1,
                            rect.Width - 8, titleH - 1);
                        g.DrawString(slicer.Name ?? "Segmentador", font, text, titleRect, sf);
                    }
                    PivotRange pivot;
                    int field;
                    List<string> values = SlicerItems(slicer, out pivot, out field);
                    int itemH = Math.Max(14, (int)(18 * _zoom * _uiScale));
                    int maxItems = Math.Max(1, (rect.Height - titleH - 6) / itemH);
                    int count = Math.Min(maxItems, values.Count + 1);
                    for (int i = 0; i < count; i++)
                    {
                        string value = i == 0 ? null : values[i - 1];
                        bool selected = i == 0 ? string.IsNullOrEmpty(slicer.SelectedValue) :
                            string.Equals(slicer.SelectedValue, value, StringComparison.OrdinalIgnoreCase);
                        Rectangle itemRect = new Rectangle(rect.Left + 3, rect.Top + titleH + 3 + i * itemH,
                            Math.Max(1, rect.Width - 6), itemH - 1);
                        if (selected)
                            using (SolidBrush sel = new SolidBrush(Color.FromArgb(222, 235, 247)))
                                g.FillRectangle(sel, itemRect);
                        using (Font font = new Font("Segoe UI", Math.Max(7f, 8f * _zoom * _uiScale)))
                        using (SolidBrush text = new SolidBrush(Color.FromArgb(50, 50, 50)))
                        using (StringFormat sf = new StringFormat())
                        {
                            sf.Alignment = StringAlignment.Near;
                            sf.LineAlignment = StringAlignment.Center;
                            sf.Trimming = StringTrimming.EllipsisCharacter;
                            g.DrawString(i == 0 ? "Todos" : value, font, text,
                                new Rectangle(itemRect.Left + 3, itemRect.Top, itemRect.Width - 6,
                                    itemRect.Height), sf);
                        }
                    }
                }
                finally { g.Restore(state); }
            }
        }

        private List<string> SlicerItems(SheetSlicer slicer, out PivotRange pivot, out int field)
        {
            pivot = null; field = -1;
            List<string> values = new List<string>();
            if (Sheet == null || slicer == null) return values;
            foreach (PivotRange p in Sheet.PivotRanges)
            {
                if (p.Cache == null) continue;
                int index = p.Cache.FieldNames.FindIndex(delegate(string name)
                {
                    return string.Equals(name, slicer.FieldName ?? slicer.Name,
                        StringComparison.OrdinalIgnoreCase);
                });
                if (index < 0) continue;
                pivot = p;
                field = index;
                HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (List<CellValue> row in p.Cache.Rows)
                {
                    if (index >= row.Count) continue;
                    string value = row[index].ToDisplayString();
                    if (seen.Add(value)) values.Add(value);
                }
                break;
            }
            return values;
        }

        private SheetSlicer SlicerItemAt(Point p, out string value)
        {
            value = null;
            if (Sheet == null) return null;
            foreach (SheetSlicer slicer in Sheet.Slicers)
            {
                Rectangle rect;
                if (!TrySlicerRect(slicer, out rect) || !rect.Contains(p)) continue;
                int titleH = Math.Min(24, Math.Max(18, rect.Height / 6));
                int itemH = Math.Max(14, (int)(18 * _zoom * _uiScale));
                int maxItems = Math.Max(1, (rect.Height - titleH - 6) / itemH);
                PivotRange pivot;
                int field;
                List<string> values = SlicerItems(slicer, out pivot, out field);
                int count = Math.Min(maxItems, values.Count + 1);
                for (int i = 0; i < count; i++)
                {
                    Rectangle itemRect = new Rectangle(rect.Left + 3, rect.Top + titleH + 3 + i * itemH,
                        Math.Max(1, rect.Width - 6), itemH - 1);
                    if (itemRect.Contains(p))
                    {
                        value = i == 0 ? null : values[i - 1];
                        return slicer;
                    }
                }
            }
            return null;
        }

        private bool TrySlicerRect(SheetSlicer slicer, out Rectangle rect)
        {
            rect = Rectangle.Empty;
            if (slicer == null || Sheet == null || slicer.FromRow < 0 || slicer.FromCol < 0 ||
                slicer.FromRow >= _rowCount || slicer.FromCol >= _colCount) return false;
            Rectangle from = GetDisplayRect(slicer.FromRow, slicer.FromCol);
            float z = _zoom * _uiScale;
            int x = from.Left + (int)(slicer.FromColOff / 9525.0 * z);
            int y = from.Top + (int)(slicer.FromRowOff / 9525.0 * z);
            int w, h;
            if (slicer.ToRow >= 0 && slicer.ToCol >= 0 &&
                slicer.ToRow < _rowCount && slicer.ToCol < _colCount)
            {
                Rectangle to = GetDisplayRect(slicer.ToRow, slicer.ToCol);
                w = to.Left + (int)(slicer.ToColOff / 9525.0 * z) - x;
                h = to.Top + (int)(slicer.ToRowOff / 9525.0 * z) - y;
            }
            else return false;
            if (w < 20 || h < 20) return false;
            rect = new Rectangle(x, y, w, h);
            return true;
        }

        private bool TryShapeRect(SheetShape shape, out Rectangle rect)
        {
            rect = Rectangle.Empty;
            if (shape == null || Sheet == null || shape.FromRow < 0 || shape.FromCol < 0 ||
                shape.FromRow >= _rowCount || shape.FromCol >= _colCount) return false;
            Rectangle from = GetDisplayRect(shape.FromRow, shape.FromCol);
            float z = _zoom * _uiScale;
            int x = from.Left + (int)(shape.FromColOff / 9525.0 * z);
            int y = from.Top + (int)(shape.FromRowOff / 9525.0 * z);
            int w, h;
            if (shape.ToRow >= 0 && shape.ToCol >= 0 &&
                shape.ToRow < _rowCount && shape.ToCol < _colCount)
            {
                Rectangle to = GetDisplayRect(shape.ToRow, shape.ToCol);
                w = to.Left + (int)(shape.ToColOff / 9525.0 * z) - x;
                h = to.Top + (int)(shape.ToRowOff / 9525.0 * z) - y;
            }
            else { w = 120; h = 28; }
            if (w < 4 || h < 4) return false;
            rect = new Rectangle(x, y, w, h);
            return true;
        }

        private SheetShape ShapeAt(Point p)
        {
            if (Sheet == null) return null;
            for (int i = Sheet.Shapes.Count - 1; i >= 0; i--)
            {
                Rectangle r;
                if (TryShapeRect(Sheet.Shapes[i], out r) && r.Contains(p)) return Sheet.Shapes[i];
            }
            return null;
        }

        private SheetImage ImageAt(Point p, out bool resize)
        {
            resize = false;
            if (Sheet == null) return null;
            for (int i = Sheet.Images.Count - 1; i >= 0; i--)
            {
                Rectangle r;
                if (!TryImageRect(Sheet.Images[i], out r) || !r.Contains(p)) continue;
                resize = p.X >= r.Right - 10 && p.Y >= r.Bottom - 10;
                return Sheet.Images[i];
            }
            return null;
        }

        private void DrawImageSelection(Graphics g)
        {
            if (!HasSelectedImage) return;
            Rectangle r;
            if (!TryImageRect(_selectedImage, out r)) return;
            using (Pen p = new Pen(Theme.Green, 2f))
                g.DrawRectangle(p, r.X, r.Y, r.Width, r.Height);
            Rectangle handle = new Rectangle(r.Right - 4, r.Bottom - 4, 8, 8);
            using (SolidBrush b = new SolidBrush(Theme.Green)) g.FillRectangle(b, handle);
        }

        private bool TryChartRect(SheetChart chart, out Rectangle rect)
        {
            rect = Rectangle.Empty;
            if (chart == null || Sheet == null || chart.FromRow < 0 || chart.FromCol < 0 ||
                chart.FromRow >= _rowCount || chart.FromCol >= _colCount) return false;
            Rectangle from = GetDisplayRect(chart.FromRow, chart.FromCol);
            float z = _zoom * _uiScale;
            int x = from.Left + (int)(chart.FromColOff / 9525.0 * z);
            int y = from.Top + (int)(chart.FromRowOff / 9525.0 * z);
            int w, h;
            if (chart.ToRow >= 0 && chart.ToCol >= 0 && chart.ToRow < _rowCount && chart.ToCol < _colCount)
            {
                Rectangle to = GetDisplayRect(chart.ToRow, chart.ToCol);
                w = to.Left + (int)(chart.ToColOff / 9525.0 * z) - x;
                h = to.Top + (int)(chart.ToRowOff / 9525.0 * z) - y;
            }
            else
            {
                w = (int)(chart.WidthPx * z);
                h = (int)(chart.HeightPx * z);
            }
            if (w < 12 || h < 12) return false;
            rect = new Rectangle(x, y, w, h);
            return true;
        }

        private SheetChart ChartAt(Point p, out bool resize)
        {
            resize = false;
            if (Sheet == null) return null;
            for (int i = Sheet.Charts.Count - 1; i >= 0; i--)
            {
                SheetChart chart = Sheet.Charts[i];
                Rectangle r;
                if (!TryChartRect(chart, out r) || !r.Contains(p)) continue;
                Rectangle handle = new Rectangle(r.Right - 8, r.Bottom - 8, 12, 12);
                resize = handle.Contains(p);
                return chart;
            }
            return null;
        }

        private void DrawCharts(Graphics g, Rectangle viewport)
        {
            if (Sheet == null || Sheet.Charts.Count == 0) return;
            foreach (SheetChart chart in Sheet.Charts)
            {
                Rectangle rect;
                if (!TryChartRect(chart, out rect)) continue;
                if (!rect.IntersectsWith(viewport)) continue;
                Rectangle clip = Rectangle.Intersect(viewport, rect);
                if (clip.IsEmpty) continue;
                GraphicsState state = g.Save();
                try
                {
                    g.SetClip(clip);
                    ChartRenderer.Draw(g, rect, chart, Sheet.Workbook, Sheet);
                }
                finally { g.Restore(state); }
            }
        }

        private void DrawChartSelection(Graphics g)
        {
            if (_selectedChart == null || Sheet == null || !Sheet.Charts.Contains(_selectedChart)) return;
            Rectangle r;
            if (!TryChartRect(_selectedChart, out r)) return;
            using (Pen p = new Pen(Theme.Green, 2f))
                g.DrawRectangle(p, r.X, r.Y, r.Width - 1, r.Height - 1);
            using (SolidBrush b = new SolidBrush(Color.White))
            using (Pen p = new Pen(Theme.Green, 1f))
            {
                int[] xs = new int[] { r.Left, r.Left + r.Width / 2, r.Right - 6 };
                int[] ys = new int[] { r.Top, r.Top + r.Height / 2, r.Bottom - 6 };
                for (int i = 0; i < 3; i++)
                    for (int j = 0; j < 3; j++)
                    {
                        if (i == 1 && j == 1) continue;
                        Rectangle h = new Rectangle(xs[j], ys[i], 6, 6);
                        g.FillRectangle(b, h);
                        g.DrawRectangle(p, h);
                    }
            }
        }

        private void DrawTableBorders(Graphics g, Rectangle viewport)
        {
            if (Sheet == null || Sheet.TableFormats.Count == 0) return;
            using (Pen pen = new Pen(Color.FromArgb(105, 105, 105), 1f))
            {
                foreach (TableFormat table in Sheet.TableFormats)
                {
                    Rectangle rect = RectForCells(table.Row1, table.Col1, table.Row2, table.Col2);
                    if (!rect.IntersectsWith(viewport)) continue;
                    g.SetClip(viewport);
                    g.DrawRectangle(pen, rect.X, rect.Y, rect.Width - 1, rect.Height - 1);
                    int clipLeft = _headerWidth + (FreezeCols > 0 ? _colLeft[FreezeCols] : 0);
                    int clipTop = _headerHeight + (FreezeRows > 0 ? _rowTop[FreezeRows] : 0);
                    Rectangle innerClip = Rectangle.Intersect(viewport,
                        new Rectangle(clipLeft, clipTop,
                            Math.Max(0, _viewWidth - clipLeft), Math.Max(0, _viewHeight - clipTop)));
                    g.SetClip(innerClip);
                    for (int r = table.Row1; r < table.Row2; r++)
                    {
                        int y = CellRect(r, table.Col1).Bottom;
                        if (y < innerClip.Top || y > innerClip.Bottom) continue;
                        g.DrawLine(pen, rect.X, y, rect.Right - 1, y);
                    }
                    for (int c = table.Col1; c < table.Col2; c++)
                    {
                        int x = CellRect(table.Row1, c).Right;
                        if (x < innerClip.Left || x > innerClip.Right) continue;
                        g.DrawLine(pen, x, rect.Y, x, rect.Bottom - 1);
                    }
                    g.ResetClip();
                }
            }
        }

        private void DrawFrozenPanes(Graphics g, Rectangle viewport, int firstRow, int firstCol,
                                     int maxRow, int maxCol)
        {
            int fr = FreezeRows;
            int fc = FreezeCols;
            if (fr == 0 && fc == 0) return;
            int saveX = _scrollX;
            int saveY = _scrollY;
            int fw = fc > 0 ? _colLeft[fc] : 0;
            int fh = fr > 0 ? _rowTop[fr] : 0;
            try
            {
                if (fr > 0)
                {
                    Rectangle clip = new Rectangle(_headerWidth, _headerHeight,
                        Math.Max(0, _viewWidth - _headerWidth),
                        Math.Min(fh, Math.Max(0, _viewHeight - _headerHeight)));
                    if (clip.Height > 0)
                    {
                        g.SetClip(Rectangle.Intersect(clip, viewport));
                        using (SolidBrush bg = new SolidBrush(Color.White)) g.FillRectangle(bg, clip);
                        _scrollY = 0;
                        for (int r = 0; r < fr; r++)
                            for (int c = firstCol; c <= maxCol; c++)
                                DrawCell(g, r, c, CellRect(r, c));
                        DrawFrozenGrid(g, 0, fr - 1, firstCol, maxCol, clip);
                        _scrollY = saveY;
                    }
                }
                if (fc > 0)
                {
                    Rectangle clip = new Rectangle(_headerWidth, _headerHeight,
                        Math.Min(fw, Math.Max(0, _viewWidth - _headerWidth)),
                        Math.Max(0, _viewHeight - _headerHeight));
                    if (clip.Width > 0)
                    {
                        g.SetClip(Rectangle.Intersect(clip, viewport));
                        using (SolidBrush bg = new SolidBrush(Color.White)) g.FillRectangle(bg, clip);
                        _scrollX = 0;
                        for (int c = 0; c < fc; c++)
                            for (int r = firstRow; r <= maxRow; r++)
                                DrawCell(g, r, c, CellRect(r, c));
                        DrawFrozenGrid(g, firstRow, maxRow, 0, fc - 1, clip);
                        _scrollX = saveX;
                    }
                }
                if (fr > 0 && fc > 0)
                {
                    Rectangle clip = new Rectangle(_headerWidth, _headerHeight,
                        Math.Min(fw, Math.Max(0, _viewWidth - _headerWidth)),
                        Math.Min(fh, Math.Max(0, _viewHeight - _headerHeight)));
                    g.SetClip(Rectangle.Intersect(clip, viewport));
                    using (SolidBrush bg = new SolidBrush(Color.White)) g.FillRectangle(bg, clip);
                    _scrollX = 0; _scrollY = 0;
                    for (int r = 0; r < fr; r++)
                        for (int c = 0; c < fc; c++)
                            DrawCell(g, r, c, CellRect(r, c));
                    DrawFrozenGrid(g, 0, fr - 1, 0, fc - 1, clip);
                    _scrollX = saveX; _scrollY = saveY;
                }
                g.ResetClip();
            }
            finally
            {
                _scrollX = saveX;
                _scrollY = saveY;
                g.ResetClip();
            }
        }

        private void DrawFrozenGrid(Graphics g, int r1, int r2, int c1, int c2, Rectangle clip)
        {
            if (!ShowGridLines || !Sheet.ShowGridLines) return;
            using (Pen p = new Pen(Theme.GridLine, 1f))
            {
                for (int r = r1; r <= r2 && r < _rowCount; r++)
                {
                    int y = _headerHeight + _rowTop[r] - _scrollY;
                    if (y >= clip.Top && y <= clip.Bottom) g.DrawLine(p, clip.Left, y, clip.Right, y);
                }
                for (int c = c1; c <= c2 && c < _colCount; c++)
                {
                    int x = _headerWidth + _colLeft[c] - _scrollX;
                    if (x >= clip.Left && x <= clip.Right) g.DrawLine(p, x, clip.Top, x, clip.Bottom);
                }
            }
        }

        private void DrawFrozenHeaders(Graphics g, int fr, int fc)
        {
            using (SolidBrush hb = new SolidBrush(Theme.GridHeaderBack))
            using (SolidBrush tb = new SolidBrush(Theme.GridHeaderText))
            using (Font f = Theme.Font(8f))
            using (StringFormat sf = new StringFormat())
            {
                sf.Alignment = StringAlignment.Center;
                sf.LineAlignment = StringAlignment.Center;
                for (int r = 0; r < fr; r++)
                {
                    Rectangle rr = new Rectangle(0, _headerHeight + _rowTop[r] - 0,
                        _headerWidth, _rowTop[r + 1] - _rowTop[r]);
                    g.FillRectangle(hb, rr);
                    g.DrawString((r + 1).ToString(), f, tb, rr, sf);
                }
                for (int c = 0; c < fc; c++)
                {
                    Rectangle cr = new Rectangle(_headerWidth + _colLeft[c],
                        0, _colLeft[c + 1] - _colLeft[c], _headerHeight);
                    g.FillRectangle(hb, cr);
                    g.DrawString(AddressUtil.ColumnName(c), f, tb, cr, sf);
                }
                if (fr > 0 && fc > 0)
                    g.FillRectangle(hb, 0, 0, _headerWidth, _headerHeight);
            }
        }

        private void DrawHeaders(Graphics g, int firstRow, int firstCol, int maxRow, int maxCol)
        {
            using (SolidBrush hb = new SolidBrush(Theme.GridHeaderBack))
            {
                g.FillRectangle(hb, 0, 0, _viewWidth, _headerHeight);
                g.FillRectangle(hb, 0, 0, _headerWidth, _viewHeight);
            }
            Font hf = Theme.Font(Math.Max(7.5f, 8.4f * _zoom));
            Font hfb = Theme.Font(Math.Max(7.5f, 8.4f * _zoom), true);
            for (int c = firstCol; c <= maxCol + 1 && c < _colCount; c++)
            {
                int x = _headerWidth + _colLeft[c] - _scrollX;
                int w = _colLeft[c + 1] - _colLeft[c];
                if (w <= 0) continue;
                if (x + w < _headerWidth) continue;
                if (x > _viewWidth) break;
                bool selected = c >= _selLeft && c <= _selRight;
                bool active = c == _activeCol;
                if (selected)
                {
                    using (SolidBrush b = new SolidBrush(active ? Theme.GridHeaderActive : Theme.GridHeaderSelected))
                        g.FillRectangle(b, x, 0, w, _headerHeight);
                }
                using (Pen p = new Pen(Theme.GridHeaderBorder))
                {
                    g.DrawLine(p, x + w - 1, 0, x + w - 1, _headerHeight);
                    g.DrawLine(p, x, _headerHeight - 1, x + w, _headerHeight - 1);
                }
                Rectangle tr = new Rectangle(x, 0, w, _headerHeight);
                TextRenderer.DrawText(g, AddressUtil.ColumnName(c), selected ? hfb : hf, tr,
                    Theme.GridHeaderText,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
            }
            for (int r = firstRow; r <= maxRow + 1 && r < _rowCount; r++)
            {
                int y = _headerHeight + _rowTop[r] - _scrollY;
                int rh = _rowTop[r + 1] - _rowTop[r];
                if (rh <= 0) continue;
                if (y + rh < _headerHeight) continue;
                if (y > _viewHeight) break;
                bool selected = r >= _selTop && r <= _selBottom;
                bool active = r == _activeRow;
                if (selected)
                {
                    using (SolidBrush b = new SolidBrush(active ? Theme.GridHeaderActive : Theme.GridHeaderSelected))
                        g.FillRectangle(b, 0, y, _headerWidth, rh);
                }
                using (Pen p = new Pen(Theme.GridHeaderBorder))
                {
                    g.DrawLine(p, 0, y + rh - 1, _headerWidth, y + rh - 1);
                    g.DrawLine(p, _headerWidth - 1, y, _headerWidth - 1, y + rh);
                }
                Rectangle tr = new Rectangle(0, y, _headerWidth - 4, rh);
                TextRenderer.DrawText(g, (r + 1).ToString(CultureInfo.InvariantCulture), selected ? hfb : hf, tr,
                    Theme.GridHeaderText,
                    TextFormatFlags.Right | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
            }
            using (SolidBrush b = new SolidBrush(Theme.GridHeaderBack)) g.FillRectangle(b, 0, 0, _headerWidth, _headerHeight);
            using (Pen p = new Pen(Theme.GridHeaderBorder))
            {
                g.DrawLine(p, 0, _headerHeight - 1, _headerWidth, _headerHeight - 1);
                g.DrawLine(p, _headerWidth - 1, 0, _headerWidth - 1, _headerHeight);
            }
            using (SolidBrush b = new SolidBrush(Color.FromArgb(150, 150, 150)))
            {
                Point[] pts = new Point[]
                {
                    new Point(_headerWidth - 14, _headerHeight - 5),
                    new Point(_headerWidth - 5, _headerHeight - 5),
                    new Point(_headerWidth - 5, _headerHeight - 14)
                };
                g.FillPolygon(b, pts);
            }
            int fr1, fc1, fr2, fc2;
            if (TryFilterRange(out fr1, out fc1, out fr2, out fc2))
            {
                using (SolidBrush b = new SolidBrush(Color.FromArgb(90, 90, 90)))
                {
                    for (int c = Math.Max(fc1, firstCol); c <= Math.Min(fc2, maxCol + 1) && c < _colCount; c++)
                    {
                        int x = _headerWidth + _colLeft[c] - _scrollX;
                        int w = _colLeft[c + 1] - _colLeft[c];
                        if (x + w < _headerWidth || x > _viewWidth) continue;
                        float ax = x + w - 10;
                        g.FillPolygon(b, new PointF[]
                        {
                            new PointF(ax, 5), new PointF(ax + 7, 5), new PointF(ax + 3.5f, 10)
                        });
                    }
                }
            }
            hf.Dispose();
            hfb.Dispose();
        }

        public int ReplaceAll(string find, string replace, bool matchCase, bool wholeCell, bool includeFormulas)
        {
            if (Sheet == null || string.IsNullOrEmpty(find)) return 0;
            if (!EnsureWritable()) return 0;
            List<long> keys = new List<long>(Sheet.Cells.Keys);
            keys.Sort();
            StringComparison cmp = matchCase ? StringComparison.Ordinal : StringComparison.CurrentCultureIgnoreCase;
            List<long> changed = new List<long>();
            Dictionary<long, string> values = new Dictionary<long, string>();
            foreach (long key in keys)
            {
                Cell cell = Sheet.Cells[key];
                if (cell == null || cell.Value.Kind == ValueKind.Empty) continue;
                bool formula = cell.IsFormula;
                if (formula && !includeFormulas) continue;
                if (!formula && cell.Value.Kind == ValueKind.Number) continue;
                string source = formula ? (cell.Raw ?? "") : (cell.Value.Text ?? "");
                string hitText = wholeCell
                    ? (string.Equals(source, find, cmp) ? source : null)
                    : source;
                if (hitText == null) continue;
                bool hit = wholeCell || source.IndexOf(find, cmp) >= 0;
                if (!hit) continue;
                string replaced = wholeCell ? (replace ?? "")
                    : ReplaceIgnoreCase(source, find, replace ?? "", cmp);
                if (string.Equals(replaced, source, StringComparison.Ordinal)) continue;
                changed.Add(key);
                values[key] = replaced;
            }
            if (changed.Count == 0) return 0;
            OnBeforeChange();
            foreach (long key in changed)
            {
                int row = (int)(key >> 14);
                int col = (int)(key & 0x3FFF);
                Cell cell = Sheet.Cells[key];
                if (cell.IsFormula) Sheet.SetInput(row, col, values[key]);
                else Sheet.SetTextLiteral(row, col, values[key]);
            }
            OnChanged();
            Invalidate();
            return changed.Count;
        }

        private static string ReplaceIgnoreCase(string source, string find, string replace, StringComparison cmp)
        {
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            int pos = 0;
            while (pos < source.Length)
            {
                int hit = source.IndexOf(find, pos, cmp);
                if (hit < 0) { sb.Append(source, pos, source.Length - pos); break; }
                sb.Append(source, pos, hit - pos);
                sb.Append(replace);
                pos = hit + find.Length;
            }
            return sb.ToString();
        }

        public bool FindNext(string text, bool matchCase, bool wholeCell, bool includeFormulas,
                             int startRow, int startCol, out int foundRow, out int foundCol)
        {
            foundRow = -1; foundCol = -1;
            if (Sheet == null || text == null || text.Length == 0) return false;
            List<long> keys = new List<long>();
            foreach (KeyValuePair<long, Cell> kv in Sheet.Cells)
                if (kv.Value.Value.Kind != ValueKind.Empty) keys.Add(kv.Key);
            if (keys.Count == 0) return false;
            keys.Sort();
            long startKey = Sheet.Key(Math.Max(0, startRow), Math.Max(0, startCol));
            int startIndex = 0;
            while (startIndex < keys.Count && keys[startIndex] <= startKey) startIndex++;
            if (startIndex >= keys.Count) startIndex = 0;
            StringComparison cmp = matchCase ? StringComparison.Ordinal : StringComparison.CurrentCultureIgnoreCase;
            for (int k = 0; k < keys.Count; k++)
            {
                long key = keys[(startIndex + k) % keys.Count];
                Cell cell = Sheet.Cells[key];
                string display = NumberFormatter.Format(cell.Value, cell.EffectiveStyle);
                string raw = cell.Raw ?? display;
                bool hit;
                if (wholeCell)
                    hit = string.Equals(display, text, cmp) || (includeFormulas && string.Equals(raw, text, cmp));
                else
                    hit = display.IndexOf(text, cmp) >= 0 || (includeFormulas && raw.IndexOf(text, cmp) >= 0);
                if (hit)
                {
                    foundRow = (int)(key >> 14);
                    foundCol = (int)(key & 0x3FFF);
                    return true;
                }
            }
            return false;
        }

        public void ClearClipboardState()
        {
            _clipboard = null;
            _blink.Stop();
            Invalidate();
        }
    }
}
