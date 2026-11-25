using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace CalcRapido
{
    public sealed class MainForm : Form
    {
        private Workbook _workbook;
        private string _filePath;
        private string _displayName = "Libro1";
        private bool _dirty;
        private bool _syncing;
        private bool _closing;
        private bool _formulaBarBusy;
        private Workbook _pendingCompatibilityWarning;

        private sealed class GridViewState
        {
            public string SheetName;
            public int SheetIndex;
            public int ActiveRow;
            public int ActiveCol;
            public int SelTop;
            public int SelLeft;
            public int SelBottom;
            public int SelRight;
            public int ScrollX;
            public int ScrollY;
        }
        private readonly UndoManager _undo = new UndoManager();

        private RibbonBar _ribbon;
        private SheetGrid _grid;
        private SheetTabs _tabs;
        private Panel _formulaPanel;
        private TextBox _nameBox;
        private TextBox _formulaBox;
        private Label _statusLabel;
        private Label _statsLabel;
        private Label _zoomLabel;
        private FindDialog _findDialog;
        private readonly List<string> _recentFiles = new List<string>();
        private Timer _recoveryTimer;
        private DateTime _lastRecoveryAttempt = DateTime.MinValue;
        private volatile bool _recoveryBusy;
        private bool _largeRecoveryWarned;
        private bool _recoveryErrorWarned;
        private const long MaxRecoveryCells = 300000;
        private bool _startupComplete;

        public MainForm(string initialPath)
        {
            Text = "Libro1 - Mini Excel";
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(1280, 800);
            MinimumSize = new Size(940, 560);
            BackColor = Color.White;
            Font = Theme.Font(9f);
            KeyPreview = true;
            AutoScaleMode = AutoScaleMode.Dpi;
            // El registro OLE de DragDrop puede fallar en algunas sesiones
            // Windows. Se activa después de crear la ventana y, si falla, se
            // desactiva sin impedir el arranque de la aplicación.
            AllowDrop = false;
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }

            BuildInterface();
            LoadRecentFiles();
            _ribbon.SetRecentFiles(_recentFiles);
            NewBook(false);
            if (!string.IsNullOrEmpty(initialPath))
            {
                if (File.Exists(initialPath)) OpenPath(initialPath);
                else MessageBox.Show(this, "No se encontró el archivo:\n" + initialPath, "Mini Excel",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            StartRecoveryTimer();
            FormClosing += MainFormClosing;
            DragEnter += delegate(object s, DragEventArgs e)
            {
                if (e.Data != null && e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
            };
            DragDrop += delegate(object s, DragEventArgs e)
            {
                string[] files = e.Data.GetData(DataFormats.FileDrop) as string[];
                if (files != null && files.Length > 0) OpenPath(files[0]);
            };
            Shown += delegate
            {
                TryRestoreRecoveredDocument();
                _startupComplete = true;
                _grid.Focus();
                try { AllowDrop = true; }
                catch { AllowDrop = false; }
                if (_pendingCompatibilityWarning != null)
                {
                    Workbook pending = _pendingCompatibilityWarning;
                    _pendingCompatibilityWarning = null;
                    ShowCompatibilityWarning(pending);
                }
            };
        }

        private void BuildInterface()
        {
            _grid = new SheetGrid();
            _grid.Dock = DockStyle.Fill;
            _grid.FormulaValidator = ValidateFormula;
            _grid.BeforeChange += delegate { PushUndo(); };
            _grid.Changed += delegate { GridChanged(); };
            _grid.SelectionChanged += delegate { UpdateSelectionUi(); };
            _grid.EditStateChanged += delegate { UpdateEditState(); };
            _grid.EditTextChanged += delegate(string text)
            {
                if (!_formulaBox.Focused) SetFormulaText(text);
            };
            _grid.ZoomChanged += delegate { UpdateZoomLabel(); };
            _grid.ContextMenuRequested += delegate { ShowGridContextMenu(); };

            _tabs = new SheetTabs();
            _tabs.Dock = DockStyle.Bottom;
            _tabs.SheetSelected += delegate(int idx) { SelectSheet(idx); };
            _tabs.AddRequested += delegate { AddSheet(); };
            _tabs.RenameRequested += delegate(int idx) { RenameSheet(idx); };
            _tabs.DeleteRequested += delegate(int idx) { DeleteSheet(idx); };

            Panel status = new Panel();
            status.Dock = DockStyle.Bottom;
            status.Height = 25;
            status.BackColor = Theme.StatusBack;
            status.Paint += delegate(object s, PaintEventArgs e)
            {
                using (Pen p = new Pen(Theme.Border)) e.Graphics.DrawLine(p, 0, 0, status.Width, 0);
            };
            TableLayoutPanel statusLayout = new TableLayoutPanel();
            statusLayout.Dock = DockStyle.Fill;
            statusLayout.ColumnCount = 3;
            statusLayout.RowCount = 1;
            statusLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 118));
            statusLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            statusLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 146));
            statusLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            statusLayout.BackColor = Theme.StatusBack;
            status.Controls.Add(statusLayout);

            _statusLabel = new Label();
            _statusLabel.Text = "Listo";
            _statusLabel.Dock = DockStyle.Fill;
            _statusLabel.TextAlign = ContentAlignment.MiddleLeft;
            _statusLabel.Padding = new Padding(8, 0, 0, 0);
            _statusLabel.ForeColor = Color.FromArgb(75, 75, 75);
            _statusLabel.Font = Theme.Font(8.6f);
            statusLayout.Controls.Add(_statusLabel, 0, 0);

            _statsLabel = new Label();
            _statsLabel.Text = "";
            _statsLabel.Dock = DockStyle.Fill;
            _statsLabel.TextAlign = ContentAlignment.MiddleLeft;
            _statsLabel.ForeColor = Color.FromArgb(75, 75, 75);
            _statsLabel.Font = Theme.Font(8.6f);
            statusLayout.Controls.Add(_statsLabel, 1, 0);

            Panel zoomPanel = new Panel();
            zoomPanel.Dock = DockStyle.Fill;
            zoomPanel.BackColor = Theme.StatusBack;
            Button zout = new Button();
            zout.Text = "\u2212";
            zout.SetBounds(2, 2, 24, 21);
            zout.FlatStyle = FlatStyle.Flat;
            zout.FlatAppearance.BorderSize = 0;
            zout.Font = Theme.Font(11f, true);
            zout.ForeColor = Color.FromArgb(80, 80, 80);
            zout.Click += delegate { _grid.SetZoom(_grid.Zoom - 0.1f); };
            zoomPanel.Controls.Add(zout);
            _zoomLabel = new Label();
            _zoomLabel.Text = "100%";
            _zoomLabel.SetBounds(26, 0, 52, 24);
            _zoomLabel.TextAlign = ContentAlignment.MiddleCenter;
            _zoomLabel.ForeColor = Color.FromArgb(75, 75, 75);
            _zoomLabel.Font = Theme.Font(8.6f);
            zoomPanel.Controls.Add(_zoomLabel);
            Button zin = new Button();
            zin.Text = "+";
            zin.SetBounds(78, 2, 24, 21);
            zin.FlatStyle = FlatStyle.Flat;
            zin.FlatAppearance.BorderSize = 0;
            zin.Font = Theme.Font(11f, true);
            zin.ForeColor = Color.FromArgb(80, 80, 80);
            zin.Click += delegate { _grid.SetZoom(_grid.Zoom + 0.1f); };
            zoomPanel.Controls.Add(zin);
            Button z100 = new Button();
            z100.Text = "100%";
            z100.SetBounds(104, 2, 40, 21);
            z100.FlatStyle = FlatStyle.Flat;
            z100.FlatAppearance.BorderSize = 0;
            z100.Font = Theme.Font(7.8f);
            z100.ForeColor = Theme.Green;
            z100.Click += delegate { _grid.SetZoom(1f); };
            zoomPanel.Controls.Add(z100);
            statusLayout.Controls.Add(zoomPanel, 2, 0);

            _formulaPanel = new Panel();
            _formulaPanel.Dock = DockStyle.Top;
            _formulaPanel.Height = 34;
            _formulaPanel.BackColor = Color.White;
            _formulaPanel.Paint += delegate(object s, PaintEventArgs e)
            {
                using (Pen p = new Pen(Theme.Border)) e.Graphics.DrawLine(p, 0, _formulaPanel.Height - 1, _formulaPanel.Width, _formulaPanel.Height - 1);
            };
            _nameBox = new TextBox();
            _nameBox.SetBounds(5, 5, 120, 24);
            _nameBox.BorderStyle = BorderStyle.FixedSingle;
            _nameBox.Font = Theme.Font(9f);
            _nameBox.CharacterCasing = CharacterCasing.Upper;
            _nameBox.KeyDown += NameBoxKeyDown;
            _nameBox.GotFocus += delegate { _nameBox.SelectAll(); };
            _nameBox.LostFocus += delegate { if (!_syncing && _grid != null) _nameBox.Text = _grid.ActiveCellName; };
            _formulaPanel.Controls.Add(_nameBox);

            Button cancel = new Button();
            cancel.Text = "\u2715";
            cancel.SetBounds(128, 4, 24, 25);
            cancel.FlatStyle = FlatStyle.Flat;
            cancel.FlatAppearance.BorderSize = 0;
            cancel.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
            cancel.ForeColor = Color.FromArgb(176, 58, 58);
            cancel.Cursor = Cursors.Hand;
            cancel.TabStop = false;
            cancel.Click += delegate
            {
                if (_grid.IsEditing) _grid.CancelEdit();
                SetFormulaText(GetActiveRaw());
                _grid.Focus();
            };
            _formulaPanel.Controls.Add(cancel);

            Button accept = new Button();
            accept.Text = "\u2713";
            accept.SetBounds(152, 4, 24, 25);
            accept.FlatStyle = FlatStyle.Flat;
            accept.FlatAppearance.BorderSize = 0;
            accept.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
            accept.ForeColor = Theme.Green;
            accept.Cursor = Cursors.Hand;
            accept.TabStop = false;
            accept.Click += delegate { ApplyFormulaBar(false); };
            _formulaPanel.Controls.Add(accept);

            Button fx = new Button();
            fx.Text = "fx";
            fx.SetBounds(178, 4, 26, 25);
            fx.FlatStyle = FlatStyle.Flat;
            fx.FlatAppearance.BorderSize = 0;
            fx.Font = new Font("Georgia", 10.5f, FontStyle.Italic);
            fx.ForeColor = Theme.Green;
            fx.Cursor = Cursors.Hand;
            fx.Click += delegate { InsertFunctionDialog(); };
            _formulaPanel.Controls.Add(fx);

            _formulaBox = new TextBox();
            _formulaBox.SetBounds(208, 5, 800, 24);
            _formulaBox.BorderStyle = BorderStyle.FixedSingle;
            _formulaBox.Font = Theme.Font(10f);
            _formulaBox.KeyDown += FormulaBoxKeyDown;
            _formulaBox.Leave += FormulaBoxLeave;
            _formulaBox.GotFocus += delegate { _formulaBox.SelectAll(); };
            _formulaPanel.Controls.Add(_formulaBox);
            _formulaPanel.Resize += delegate { LayoutFormulaBar(); };

            _ribbon = new RibbonBar();
            _ribbon.Dock = DockStyle.Top;
            _ribbon.Command += Execute;
            _ribbon.RecentFileSelected += OpenRecentFile;
            _ribbon.FontSizeSelected += delegate(float size)
            {
                float s = Math.Max(6f, Math.Min(72f, size));
                _grid.ApplyStyle(delegate(CellStyle st) { st.FontSize = s; });
            };
            _ribbon.FontNameSelected += delegate(string name)
            {
                if (string.IsNullOrEmpty(name)) return;
                _grid.ApplyStyle(delegate(CellStyle st) { st.FontName = name; });
            };

            Controls.Add(_grid);
            Controls.Add(_tabs);
            Controls.Add(status);
            Controls.Add(_formulaPanel);
            Controls.Add(_ribbon);
            LayoutFormulaBar();
        }

        private void LayoutFormulaBar()
        {
            if (_formulaPanel == null || _formulaBox == null) return;
            int w = _formulaPanel.ClientSize.Width - _formulaBox.Left - 8;
            _formulaBox.Width = Math.Max(80, w);
        }

        private void NewBook(bool confirm)
        {
            if (confirm && !ConfirmSave()) return;
            if (confirm) ClearRecovery();
            _workbook = new Workbook();
            _workbook.Recalculate();
            _filePath = null;
            _displayName = "Libro1";
            _dirty = false;
            _undo.Clear();
            BindWorkbook();
            _statusLabel.Text = "Listo";
        }

        private GridViewState CaptureView()
        {
            if (_grid == null || _grid.Sheet == null || _workbook == null) return null;
            GridViewState s = new GridViewState();
            s.SheetName = _grid.Sheet.Name;
            s.SheetIndex = _workbook.Sheets.IndexOf(_grid.Sheet);
            s.ActiveRow = _grid.ActiveRow;
            s.ActiveCol = _grid.ActiveCol;
            s.SelTop = _grid.SelectionTop;
            s.SelLeft = _grid.SelectionLeft;
            s.SelBottom = _grid.SelectionBottom;
            s.SelRight = _grid.SelectionRight;
            s.ScrollX = _grid.ScrollX;
            s.ScrollY = _grid.ScrollY;
            return s;
        }

        private void RestoreView(GridViewState state)
        {
            if (state == null || _grid == null || _workbook == null) return;
            int index = -1;
            if (!string.IsNullOrEmpty(state.SheetName))
            {
                for (int i = 0; i < _workbook.Sheets.Count; i++)
                    if (string.Equals(_workbook.Sheets[i].Name, state.SheetName, StringComparison.Ordinal))
                    {
                        index = i;
                        break;
                    }
            }
            if (index < 0 && state.SheetIndex >= 0 && state.SheetIndex < _workbook.Sheets.Count)
                index = state.SheetIndex;
            if (index < 0) index = 0;
            _workbook.ActiveSheetIndex = index;
            BindWorkbook();
            _grid.RestoreView(state.ActiveRow, state.ActiveCol,
                state.SelTop, state.SelLeft, state.SelBottom, state.SelRight,
                state.ScrollX, state.ScrollY);
            _tabs.ActiveIndex = index;
            _tabs.RefreshTabs();
        }

        private void BindWorkbook()
        {
            ConditionalEvaluator.Reset();
            ChartData.ThemePalette = _workbook.ThemePalette;
            if (_workbook.Sheets.Count == 0) _workbook.AddSheet("Hoja1");
            if (_workbook.ActiveSheetIndex < 0 || _workbook.ActiveSheetIndex >= _workbook.Sheets.Count)
                _workbook.ActiveSheetIndex = 0;
            _grid.Attach(_workbook, _workbook.ActiveSheetIndex);
            _tabs.Workbook = _workbook;
            _tabs.ActiveIndex = _workbook.ActiveSheetIndex;
            _tabs.RefreshTabs();
            _syncing = true;
            try
            {
                _nameBox.Text = _grid.ActiveCellName;
                SetFormulaText(GetActiveRaw());
            }
            finally { _syncing = false; }
            UpdateTitle();
            UpdateSelectionUi();
            UpdateQuickState();
        }

        private static string[] RecentFilePaths
        {
            get
            {
                return new string[]
                {
                    Path.Combine(Application.StartupPath, "recent.txt"),
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                        "MiniExcel", "recent.txt")
                };
            }
        }

        private void LoadRecentFiles()
        {
            _recentFiles.Clear();
            try
            {
                string file = null;
                foreach (string candidate in RecentFilePaths)
                    if (File.Exists(candidate)) { file = candidate; break; }
                if (file == null) return;
                foreach (string line in File.ReadAllLines(file))
                {
                    string path = (line ?? "").Trim();
                    if (path.Length == 0 || _recentFiles.Count >= 10) continue;
                    bool exists = false;
                    foreach (string old in _recentFiles)
                        if (string.Equals(old, path, StringComparison.OrdinalIgnoreCase)) { exists = true; break; }
                    if (!exists) _recentFiles.Add(path);
                }
            }
            catch { }
        }

        private void SaveRecentFiles()
        {
            foreach (string file in RecentFilePaths)
            {
                try
                {
                    string dir = Path.GetDirectoryName(file);
                    if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
                    File.WriteAllLines(file, _recentFiles.ToArray());
                    return;
                }
                catch { }
            }
        }

        private void AddRecentFile(string path)
        {
            if (string.IsNullOrEmpty(path)) return;
            try { path = Path.GetFullPath(path); } catch { return; }
            for (int i = _recentFiles.Count - 1; i >= 0; i--)
                if (string.Equals(_recentFiles[i], path, StringComparison.OrdinalIgnoreCase))
                    _recentFiles.RemoveAt(i);
            _recentFiles.Insert(0, path);
            while (_recentFiles.Count > 10) _recentFiles.RemoveAt(_recentFiles.Count - 1);
            SaveRecentFiles();
            if (_ribbon != null) _ribbon.SetRecentFiles(_recentFiles);
        }

        private void OpenRecentFile(string path)
        {
            if (string.IsNullOrEmpty(path)) return;
            if (!File.Exists(path))
            {
                MessageBox.Show(this, "El archivo ya no existe:\n\n" + path, "Mini Excel",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                for (int i = _recentFiles.Count - 1; i >= 0; i--)
                    if (string.Equals(_recentFiles[i], path, StringComparison.OrdinalIgnoreCase))
                        _recentFiles.RemoveAt(i);
                SaveRecentFiles();
                if (_ribbon != null) _ribbon.SetRecentFiles(_recentFiles);
                return;
            }
            OpenPath(path);
        }

        private void StartRecoveryTimer()
        {
            _recoveryTimer = new Timer();
            _recoveryTimer.Interval = 30000;
            _recoveryTimer.Tick += delegate { RecoveryTick(); };
            _recoveryTimer.Start();
            FormClosed += delegate
            {
                if (_recoveryTimer != null)
                {
                    _recoveryTimer.Stop();
                    _recoveryTimer.Dispose();
                    _recoveryTimer = null;
                }
            };
        }

        private void RecoveryTick()
        {
            if (_recoveryBusy || _closing || _syncing || _workbook == null || _grid == null) return;
            if (!string.IsNullOrEmpty(RecoveryManager.LastError) && !_recoveryErrorWarned)
            {
                _recoveryErrorWarned = true;
                _statusLabel.Text = "La recuperación automática falló: " + RecoveryManager.LastError;
            }
            if (!_dirty || _grid.IsEditing) return;
            if ((DateTime.UtcNow - _lastRecoveryAttempt).TotalSeconds < 12) return;
            long cells = 0;
            foreach (Sheet sheet in _workbook.Sheets) cells += sheet.Cells.Count;
            if (cells > MaxRecoveryCells)
            {
                if (!_largeRecoveryWarned)
                {
                    _largeRecoveryWarned = true;
                    _statusLabel.Text = "Libro grande: recuperación automática desactivada (>300.000 celdas). Guarda manualmente.";
                }
                return;
            }
            SaveRecoverySnapshot();
        }

        private void SaveRecoverySnapshot()
        {
            if (_workbook == null || _recoveryBusy || _closing) return;
            _recoveryBusy = true;
            _lastRecoveryAttempt = DateTime.UtcNow;
            Workbook snapshot;
            try
            {
                snapshot = _workbook.Clone();
            }
            catch
            {
                _recoveryBusy = false;
                return;
            }
            string path = _filePath;
            string name = _displayName;
            System.Threading.ThreadPool.QueueUserWorkItem(delegate
            {
                try
                {
                    if (!_closing) RecoveryManager.Save(snapshot, path, name);
                }
                catch (Exception ex)
                {
                    RecoveryManager.LastError = ex.GetType().Name + ": " + ex.Message;
                }
                finally { _recoveryBusy = false; }
            });
        }

        private void ClearRecovery()
        {
            try { RecoveryManager.Clear(); } catch { }
        }

        private void ClearRecovery(RecoveryInfo info)
        {
            try { RecoveryManager.Clear(info); } catch { }
        }

        private void TryRestoreRecoveredDocument()
        {
            RecoveryInfo info;
            if (!RecoveryManager.TryRead(out info))
            {
                ClearRecovery(info);
                return;
            }
            string where = string.IsNullOrEmpty(info.OriginalPath)
                ? info.DisplayName : info.OriginalPath;
            DialogResult answer = MessageBox.Show(this,
                "Mini Excel encontró cambios sin guardar de una sesión anterior.\n\n" +
                where + "\nGuardado: " + info.SavedAt.ToString("g", CultureInfo.CurrentCulture) +
                "\n\n¿Deseas recuperar el documento?",
                "Recuperar documento", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (answer != DialogResult.Yes)
            {
                ClearRecovery(info);
                return;
            }
            try
            {
                Workbook recovered = RecoveryManager.LoadSnapshot(info);
                _workbook = recovered;
                _filePath = string.IsNullOrEmpty(info.OriginalPath) ? null : info.OriginalPath;
                _displayName = string.IsNullOrEmpty(info.DisplayName) ? "Libro recuperado" : info.DisplayName;
                _dirty = true;
                _undo.Clear();
                BindWorkbook();
                _statusLabel.Text = "Documento recuperado; guárdalo para conservar los cambios";
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "No se pudo recuperar el documento:\n\n" + ex.Message,
                    "Mini Excel", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                ClearRecovery(info);
            }
        }

        private void OpenPath(string path)
        {
            if (!ConfirmSave()) return;
            try
            {
                Workbook wb = FileIO.Open(path);
                wb.Recalculate();
                _workbook = wb;
                _filePath = path;
                _displayName = Path.GetFileNameWithoutExtension(path);
                _dirty = false;
                _undo.Clear();
                BindWorkbook();
                if (_startupComplete) ClearRecovery();
                AddRecentFile(path);
                _statusLabel.Text = "Listo";
                if (wb.LoadWarnings.Count > 0 || wb.UnsupportedFunctions.Count > 0)
                {
                    if (IsHandleCreated && Visible)
                        BeginInvoke(new MethodInvoker(delegate { ShowCompatibilityWarning(wb); }));
                    else
                        _pendingCompatibilityWarning = wb;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "No se pudo abrir el archivo:\n\n" + ex.Message, "Mini Excel",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ShowCompatibilityWarning(Workbook wb)
        {
            if (wb == null || IsDisposed) return;
            StringBuilder message = new StringBuilder();
            message.AppendLine("Mini Excel abrió este documento en modo de compatibilidad limitada.");
            message.AppendLine();
            message.AppendLine("Esta versión es una hoja de cálculo básica y ultraligera. " +
                "Puede abrir, editar y guardar datos, fórmulas básicas y parte del formato, " +
                "y visualizar algunos formatos simples. No es un reemplazo completo de Excel.");
            message.AppendLine();
            message.AppendLine("Se detectó:");
            int lines = 0;
            foreach (string warning in wb.LoadWarnings)
            {
                if (lines++ >= 10) break;
                message.AppendLine("• " + warning);
            }
            if (wb.UnsupportedFunctions.Count > 0)
            {
                StringBuilder list = new StringBuilder();
                int n = 0;
                foreach (string fn in wb.UnsupportedFunctions)
                {
                    if (n++ > 0) list.Append(", ");
                    list.Append(fn);
                    if (n >= 12) { list.Append(", …"); break; }
                }
                message.AppendLine("• Funciones no calculadas por esta versión: " + list + ".");
                message.AppendLine("  Si existen valores guardados por Excel, se muestran esos valores; " +
                    "al editar una fórmula no soportada puede aparecer un error.");
            }
            message.AppendLine();
            message.AppendLine("Los gráficos, imágenes y objetos avanzados pueden no mostrarse o perderse " +
                "al guardar. Guarda una copia antes de editar este archivo.");
            MessageBox.Show(this, message.ToString(), "Compatibilidad limitada",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private bool SaveBook()
        {
            if (string.IsNullOrEmpty(_filePath)) return SaveBookAs();
            if (string.Equals(Path.GetExtension(_filePath), ".xls", StringComparison.OrdinalIgnoreCase))
            {
                _statusLabel.Text = "Los .xls se guardan como .xlsx";
                return SaveBookAs();
            }
            try
            {
                _workbook.Recalculate();
                AtomicIO.Save(_filePath, delegate(string temp)
                {
                    FileIO.Save(_workbook, temp);
                });
                _dirty = false;
                ClearRecovery();
                AddRecentFile(_filePath);
                UpdateTitle();
                _statusLabel.Text = "Guardado";
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "No se pudo guardar el archivo:\n\n" + ex.Message, "Mini Excel",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private bool SaveBookAs()
        {
            using (SaveFileDialog dlg = new SaveFileDialog())
            {
                dlg.Title = "Guardar como";
                dlg.Filter = "Libro de Excel (*.xlsx)|*.xlsx|Libro de Mini Excel (*.crb)|*.crb|Archivo CSV (*.csv)|*.csv";
                dlg.DefaultExt = "xlsx";
                dlg.AddExtension = true;
                dlg.FileName = _displayName;
                dlg.RestoreDirectory = true;
                if (dlg.ShowDialog(this) != DialogResult.OK) return false;
                if (string.Equals(Path.GetExtension(dlg.FileName), ".xls", StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show(this,
                        "Mini Excel no escribe .xls. Elija un nombre .xlsx.",
                        "Mini Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return false;
                }
                string oldPath = _filePath;
                string oldName = _displayName;
                _filePath = dlg.FileName;
                _displayName = Path.GetFileNameWithoutExtension(dlg.FileName);
                if (SaveBook()) return true;
                _filePath = oldPath;
                _displayName = oldName;
                UpdateTitle();
                return false;
            }
        }

        private bool SaveStrictAs()
        {
            using (SaveFileDialog dlg = new SaveFileDialog())
            {
                dlg.Title = "Guardar copia OOXML strict";
                dlg.Filter = "Libro de Excel strict (*.xlsx)|*.xlsx";
                dlg.DefaultExt = "xlsx";
                dlg.AddExtension = true;
                dlg.FileName = _displayName + "-strict";
                dlg.RestoreDirectory = true;
                if (dlg.ShowDialog(this) != DialogResult.OK) return false;
                try
                {
                    _workbook.Recalculate();
                    AtomicIO.Save(dlg.FileName, delegate(string temp)
                    {
                        XlsxIO.SaveStrict(_workbook, temp);
                    });
                    _statusLabel.Text = "Copia OOXML strict guardada";
                    return true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "No se pudo guardar la copia strict:\n\n" + ex.Message,
                        "Mini Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }
        }

        private bool ExportCsv()
        {
            using (SaveFileDialog dlg = new SaveFileDialog())
            {
                dlg.Title = "Exportar a CSV";
                dlg.Filter = "Archivo CSV (*.csv)|*.csv";
                dlg.DefaultExt = "csv";
                dlg.AddExtension = true;
                dlg.FileName = _displayName;
                dlg.RestoreDirectory = true;
                if (dlg.ShowDialog(this) != DialogResult.OK) return false;
                try
                {
                    AtomicIO.Save(dlg.FileName, delegate(string temp)
                    {
                        FileIO.SaveCsv(_workbook, temp);
                    });
                    _statusLabel.Text = "CSV exportado";
                    return true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "No se pudo exportar:\n\n" + ex.Message, "Mini Excel",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }
        }

        private void InsertChart()
        {
            if (_workbook == null) return;
            if (_grid.IsEditing && !_grid.CommitEdit()) return;
            int r1 = _grid.SelectionTop, c1 = _grid.SelectionLeft;
            int r2 = _grid.SelectionBottom, c2 = _grid.SelectionRight;
            if (r2 <= r1 && c2 <= c1)
            {
                MessageBox.Show(this, "Selecciona un rango con datos para crear el gráfico.",
                    "Insertar gráfico", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            Sheet sheet = _workbook.ActiveSheet;
            using (ChartInsertDialog dlg = new ChartInsertDialog("Gráfico"))
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                PushUndo();
                bool header = r2 > r1 && IsTextCell(sheet, r1, c1) && IsNumberCell(sheet, r1 + 1, c1);
                int dataRow1 = header ? r1 + 1 : r1;
                SheetChart chart = new SheetChart();
                chart.ThemePalette = _workbook.ThemePalette == null
                    ? null : (Color[])_workbook.ThemePalette.Clone();
                chart.Kind = dlg.Kind;
                chart.Title = dlg.ChartTitle;
                chart.ShowLegend = dlg.ShowLegend;
                chart.ShowAxes = dlg.ShowAxes;
                chart.Stacked = dlg.Stacked;
                chart.FromRow = r1;
                chart.FromCol = Math.Min(Limits.MaxCols - 1, c2 + 1);
                if (chart.FromCol <= c2) chart.FromCol = c1;
                chart.WidthPx = 420;
                chart.HeightPx = 260;
                chart.Dirty = true;

                string catFormula = RangeFormula(sheet, dataRow1, c1, r2, c1);
                if (c2 >= c1 + 1)
                {
                    for (int c = c1 + 1; c <= c2; c++)
                    {
                        ChartSeries s = new ChartSeries();
                        if (dlg.Kind == ChartKind.Scatter)
                        {
                            s.XValuesFormula = RangeFormula(sheet, dataRow1, c1, r2, c1);
                            s.ValuesFormula = RangeFormula(sheet, dataRow1, c, r2, c);
                        }
                        else
                        {
                            s.CategoriesFormula = catFormula;
                            s.ValuesFormula = RangeFormula(sheet, dataRow1, c, r2, c);
                        }
                        if (header) s.NameFormula = RangeFormula(sheet, r1, c, r1, c);
                        chart.Series.Add(s);
                    }
                }
                else
                {
                    ChartSeries s = new ChartSeries();
                    s.ValuesFormula = RangeFormula(sheet, dataRow1, c1, r2, c1);
                    chart.Series.Add(s);
                }
                if (chart.Series.Count == 0) return;
                sheet.Charts.Add(chart);
                sheet.ChartsDirty = true;
                _dirty = true;
                _grid.SelectChart(chart);
                UpdateTitle();
                _statusLabel.Text = "Gráfico insertado";
            }
        }

        private void CropSelectedImage()
        {
            if (!_grid.HasSelectedImage) return;
            using (CropImageDialog dlg = new CropImageDialog())
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                _grid.CropSelectedImage(dlg.LeftPercent, dlg.TopPercent,
                    dlg.RightPercent, dlg.BottomPercent);
                _dirty = true;
                UpdateTitle();
            }
        }

        private void ShowValidation()
        {
            using (ValidationDialog dlg = new ValidationDialog())
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                if (dlg.Items.Length == 0 && string.IsNullOrEmpty(dlg.SourceRange))
                {
                    MessageBox.Show(this, "Escriba valores o un rango de origen.", "Mini Excel",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                _grid.ApplyValidation(_grid.SelectionTop, _grid.SelectionLeft,
                    _grid.SelectionBottom, _grid.SelectionRight, dlg.Items, dlg.Message,
                    dlg.AllowBlank, dlg.SourceRange);
                _dirty = true;
                UpdateTitle();
                _statusLabel.Text = "Validación aplicada";
            }
        }

        private void ShowFilter()
        {
            if (_grid.SelectionTop != _grid.SelectionBottom)
            {
                _grid.SetAutoFilterSelection();
                string name = AddressUtil.ColumnName(_grid.SelectionLeft);
                using (FilterDialog dlg = new FilterDialog(name))
                {
                    if (dlg.ShowDialog(this) != DialogResult.OK) return;
                    if (dlg.ClearRequested) _grid.ClearFilter();
                    else _grid.ApplyFilter(_grid.SelectionLeft, dlg.OperatorIndex, dlg.FilterText);
                }
            }
            else
            {
                MessageBox.Show(this, "Seleccione primero el rango de datos, incluida la fila de encabezados.",
                    "Mini Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void EditChartSeries()
        {
            SheetChart chart = _grid.SelectedChart;
            if (chart == null) return;
            using (ChartSeriesDialog dlg = new ChartSeriesDialog(chart))
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                PushUndo();
                chart.Series.Clear();
                foreach (ChartSeries s in dlg.Result) chart.Series.Add(s);
                chart.Dirty = true;
                _workbook.ActiveSheet.ChartsDirty = true;
                _dirty = true;
                _grid.Invalidate();
                UpdateTitle();
                _statusLabel.Text = "Series actualizadas";
            }
        }

        private void EditSelectedChart()
        {
            SheetChart chart = _grid.SelectedChart;
            if (chart == null) return;
            using (ChartInsertDialog dlg = new ChartInsertDialog(chart.Title, chart.Kind,
                chart.ShowLegend, chart.ShowAxes, chart.Stacked))
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                PushUndo();
                chart.Kind = dlg.Kind;
                chart.Title = dlg.ChartTitle;
                chart.ShowLegend = dlg.ShowLegend;
                chart.ShowAxes = dlg.ShowAxes;
                chart.Stacked = dlg.Stacked;
                chart.Dirty = true;
                _workbook.ActiveSheet.ChartsDirty = true;
                _dirty = true;
                _grid.Invalidate();
                UpdateTitle();
                _statusLabel.Text = "Gráfico actualizado";
            }
        }

        private static bool IsTextCell(Sheet sheet, int row, int col)
        {
            Cell c = sheet.GetCell(row, col);
            return c != null && c.Value.Kind == ValueKind.Text;
        }

        private static bool IsNumberCell(Sheet sheet, int row, int col)
        {
            Cell c = sheet.GetCell(row, col);
            return c != null && c.Value.Kind == ValueKind.Number;
        }

        private static string RangeFormula(Sheet sheet, int r1, int c1, int r2, int c2)
        {
            string name = sheet.Name ?? "";
            name = name.Replace("'", "''");
            return "'" + name + "'!$" + AddressUtil.ColumnName(c1) + "$" + (r1 + 1) +
                   (r1 == r2 && c1 == c2 ? "" :
                    ":$" + AddressUtil.ColumnName(c2) + "$" + (r2 + 1));
        }

        private void ShowPrintDialog()
        {
            if (_workbook == null) return;
            if (_grid.IsEditing && !_grid.CommitEdit()) return;
            using (PrintOptionsDialog dlg = new PrintOptionsDialog(_workbook, _workbook.ActiveSheetIndex))
            {
                if (dlg.ShowDialog(this) != DialogResult.OK || dlg.Options == null) return;
                PrintOptions options = dlg.Options;
                using (WorkbookPrintDocument doc = new WorkbookPrintDocument(
                    _workbook, options, _workbook.ActiveSheetIndex,
                    _grid.SelectionTop, _grid.SelectionLeft, _grid.SelectionBottom, _grid.SelectionRight))
                {
                    try
                    {
                        if (options.PreviewRequested)
                        {
                            using (PrintPreviewForm preview = new PrintPreviewForm(doc))
                                preview.ShowDialog(this);
                        }
                        else
                        {
                            using (PrintDialog pd = new PrintDialog())
                            {
                                pd.Document = doc;
                                pd.UseEXDialog = true;
                                if (pd.ShowDialog(this) == DialogResult.OK) doc.Print();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this, "No se pudo imprimir:\n\n" + ex.Message,
                            "Mini Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private bool ConfirmSave()
        {
            if (!_dirty) return true;
            DialogResult r = MessageBox.Show(this, "¿Deseas guardar los cambios en «" + _displayName + "»?",
                "Mini Excel", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (r == DialogResult.Cancel) return false;
            if (r == DialogResult.Yes) return SaveBook();
            return true;
        }

        private void MainFormClosing(object sender, FormClosingEventArgs e)
        {
            if (_closing) return;
            if (!ConfirmSave())
            {
                e.Cancel = true;
                return;
            }
            _closing = true;
            if (_recoveryTimer != null) _recoveryTimer.Stop();
            ClearRecovery();
            if (_findDialog != null && !_findDialog.IsDisposed) _findDialog.Close();
        }

        private void UpdateTitle()
        {
            Text = _displayName + (_dirty ? " *" : "") + " - Mini Excel";
        }

        private void GridChanged()
        {
            if (_workbook == null) return;
            if (HasFormulas(_workbook))
            {
                try { _workbook.Recalculate(); }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Error al recalcular:\n" + ex.Message, "Mini Excel",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            _dirty = true;
            _grid.Invalidate();
            UpdateTitle();
            UpdateStats();
            UpdateRibbonState();
            UpdateQuickState();
        }

        private static bool HasFormulas(Workbook wb)
        {
            foreach (Sheet sheet in wb.Sheets)
                foreach (Cell cell in sheet.Cells.Values)
                    if (cell.IsFormula) return true;
            return false;
        }

        private void UpdateSelectionUi()
        {
            if (_workbook == null || _grid == null) return;
            _syncing = true;
            try
            {
                if (!_nameBox.Focused) _nameBox.Text = _grid.ActiveCellName;
                if (!_formulaBox.Focused && !_grid.IsEditing) SetFormulaText(GetActiveRaw());
                if (!_grid.IsEditing) _statusLabel.Text = "Listo";
            }
            finally { _syncing = false; }
            UpdateStats();
            UpdateRibbonState();
            UpdateQuickState();
        }

        private void UpdateEditState()
        {
            if (_grid.IsEditing)
            {
                _statusLabel.Text = "Editar";
                SetFormulaText(_grid.EditingText);
            }
            else
            {
                _statusLabel.Text = "Listo";
                if (!_formulaBox.Focused) SetFormulaText(GetActiveRaw());
            }
        }

        private string GetActiveRaw()
        {
            if (_workbook == null) return "";
            Cell c = _workbook.ActiveSheet.GetCell(_grid.ActiveRow, _grid.ActiveCol);
            return c == null || c.Raw == null ? "" : Builtins.DisplayFormula(c.Raw);
        }

        private void SetFormulaText(string text)
        {
            _syncing = true;
            try { _formulaBox.Text = text ?? ""; }
            finally { _syncing = false; }
        }

        private void UpdateStats()
        {
            if (_grid == null) return;
            SelectionStats st = _grid.GetSelectionStats();
            if (st.NumericCount > 0)
            {
                string sum = NumberFormatter.FormatNumber(st.Sum, 2, true);
                string avg = NumberFormatter.FormatNumber(st.Average, 2, true);
                _statsLabel.Text = "Suma: " + sum + "     Promedio: " + avg + "     Recuento: " + st.NumericCount.ToString(CultureInfo.InvariantCulture);
            }
            else if (st.NonEmptyCount > 0)
            {
                _statsLabel.Text = "Recuento: " + st.NonEmptyCount.ToString(CultureInfo.InvariantCulture);
            }
            else _statsLabel.Text = "";
        }

        private void UpdateRibbonState()
        {
            if (_ribbon == null || _grid == null) return;
            CellStyle st = _grid.GetActiveStyle();
            _ribbon.SetChecked("bold", st.Bold);
            _ribbon.SetChecked("italic", st.Italic);
            _ribbon.SetChecked("underline", st.Underline);
            _ribbon.SetChecked("alignleft", st.Align == 0);
            _ribbon.SetChecked("aligncenter", st.Align == 1);
            _ribbon.SetChecked("alignright", st.Align == 2);
            _ribbon.SetChecked("currency", st.Format == NumberFormatKind.Currency);
            _ribbon.SetChecked("percent", st.Format == NumberFormatKind.Percent);
            _ribbon.SetChecked("comma", st.Format == NumberFormatKind.Number);
            _ribbon.SetChecked("datefmt", st.Format == NumberFormatKind.Date);
            _ribbon.SetChecked("gridlines", _grid.ShowGridLines);
            _ribbon.SetChecked("showformulas", _grid.ShowFormulas);
            _ribbon.SetFontSizeValue(st.FontSize);
            _ribbon.SetFontNameValue(st.FontName);
            _ribbon.SetQuickState(_undo.CanUndo, _undo.CanRedo);
        }

        private void UpdateQuickState()
        {
            if (_ribbon != null) _ribbon.SetQuickState(_undo.CanUndo, _undo.CanRedo);
        }

        private void UpdateZoomLabel()
        {
            _zoomLabel.Text = ((int)Math.Round(_grid.Zoom * 100)).ToString(CultureInfo.InvariantCulture) + "%";
        }

        private void PushUndo()
        {
            if (_workbook == null) return;
            _undo.Push(_workbook);
            UpdateQuickState();
        }

        private void Undo()
        {
            if (_grid.IsEditing) _grid.CancelEdit();
            GridViewState view = CaptureView();
            Workbook prev = _undo.Undo(_workbook);
            if (prev == null) return;
            _workbook = prev;
            _workbook.Recalculate();
            _dirty = true;
            RestoreView(view);
            _statusLabel.Text = "Deshacer";
        }

        private void Redo()
        {
            if (_grid.IsEditing) _grid.CancelEdit();
            GridViewState view = CaptureView();
            Workbook next = _undo.Redo(_workbook);
            if (next == null) return;
            _workbook = next;
            _workbook.Recalculate();
            _dirty = true;
            RestoreView(view);
            _statusLabel.Text = "Rehacer";
        }

        private string ValidateFormula(string formula)
        {
            string error;
            if (!FormulaEngine.Validate(formula, out error)) return error;
            return null;
        }

        private void NameBoxKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                int r1, c1, r2, c2;
                if (AddressUtil.TryParseRange(_nameBox.Text, out r1, out c1, out r2, out c2))
                {
                    _grid.SelectRange(r1, c1, r2, c2);
                    _grid.Focus();
                }
                else
                {
                    _nameBox.Text = _grid.ActiveCellName;
                    _nameBox.SelectAll();
                }
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Escape)
            {
                _nameBox.Text = _grid.ActiveCellName;
                _grid.Focus();
                e.Handled = true;
            }
        }

        private void FormulaBoxKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                ApplyFormulaBar(false);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Escape)
            {
                SetFormulaText(GetActiveRaw());
                _grid.Focus();
                e.Handled = true;
            }
        }

        private void FormulaBoxLeave(object sender, EventArgs e)
        {
            if (_formulaBarBusy || _syncing) return;
            if (string.Equals(_formulaBox.Text, GetActiveRaw(), StringComparison.Ordinal)) return;
            ApplyFormulaBar(true);
        }

        private void ApplyFormulaBar(bool fromLeave)
        {
            if (_syncing || _formulaBarBusy) return;
            _formulaBarBusy = true;
            try
            {
                string text = _formulaBox.Text;
                string error;
                if (!_grid.ApplyCellText(text, out error))
                {
                    MessageBox.Show(this, "La fórmula no es válida:\n\n" + error, "Mini Excel",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    BeginInvoke(new MethodInvoker(delegate
                    {
                        _formulaBox.Focus();
                        _formulaBox.SelectAll();
                    }));
                    return;
                }
                _statusLabel.Text = "Listo";
                SetFormulaText(GetActiveRaw());
                if (!fromLeave)
                {
                    _grid.Focus();
                    _grid.MoveActiveCell(1, 0);
                }
            }
            finally
            {
                _formulaBarBusy = false;
            }
        }

        private void SelectSheet(int index)
        {
            if (_workbook == null || index < 0 || index >= _workbook.Sheets.Count) return;
            if (_grid.IsEditing) _grid.CommitEdit();
            _workbook.ActiveSheetIndex = index;
            _grid.SetSheet(_workbook.Sheets[index]);
            _tabs.ActiveIndex = index;
            _tabs.RefreshTabs();
            _syncing = true;
            try
            {
                _nameBox.Text = _grid.ActiveCellName;
                SetFormulaText(GetActiveRaw());
            }
            finally { _syncing = false; }
            UpdateStats();
            UpdateRibbonState();
        }

        private void AddSheet()
        {
            if (_grid.IsEditing) _grid.CommitEdit();
            PushUndo();
            Sheet s = _workbook.AddSheet(null);
            _workbook.ActiveSheetIndex = _workbook.Sheets.Count - 1;
            _grid.SetSheet(s);
            _tabs.Workbook = _workbook;
            _tabs.ActiveIndex = _workbook.ActiveSheetIndex;
            _tabs.RefreshTabs();
            _dirty = true;
            UpdateTitle();
            UpdateSelectionUi();
        }

        private void RenameSheet(int index)
        {
            if (_workbook == null || index < 0 || index >= _workbook.Sheets.Count) return;
            Sheet s = _workbook.Sheets[index];
            string name = InputDialog.Show(this, "Cambiar nombre de hoja", "Nuevo nombre:", s.Name);
            if (name == null) return;
            name = name.Trim();
            if (name.Length == 0) return;
            if (name.Length > 31)
            {
                MessageBox.Show(this, "El nombre de la hoja no puede superar los 31 caracteres.", "Mini Excel");
                return;
            }
            foreach (Sheet other in _workbook.Sheets)
            {
                if (other != s && string.Equals(other.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show(this, "Ya existe una hoja con ese nombre.", "Mini Excel");
                    return;
                }
            }
            PushUndo();
            string old = s.Name;
            s.Name = name;
            foreach (Sheet sheet in _workbook.Sheets)
                foreach (Cell cell in sheet.Cells.Values)
                {
                    if (!cell.IsFormula) continue;
                    string f = FormulaEngine.RenameSheetReferences(cell.Formula, old, name);
                    if (!string.Equals(f, cell.Formula, StringComparison.Ordinal))
                    {
                        cell.Formula = f;
                        cell.Raw = f;
                        cell.Ast = null;
                        cell.AstReady = false;
                    }
                }
            _workbook.Recalculate();
            _tabs.RefreshTabs();
            _dirty = true;
            UpdateTitle();
            GridChanged();
        }

        private void DeleteSheet(int index)
        {
            if (_workbook.Sheets.Count <= 1)
            {
                MessageBox.Show(this, "Un libro debe contener al menos una hoja.", "Mini Excel");
                return;
            }
            Sheet s = _workbook.Sheets[index];
            if (MessageBox.Show(this, "¿Eliminar la hoja «" + s.Name + "»?", "Mini Excel",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            PushUndo();
            _workbook.Sheets.RemoveAt(index);
            if (_workbook.ActiveSheetIndex >= _workbook.Sheets.Count)
                _workbook.ActiveSheetIndex = _workbook.Sheets.Count - 1;
            _grid.SetSheet(_workbook.Sheets[_workbook.ActiveSheetIndex]);
            _tabs.Workbook = _workbook;
            _tabs.ActiveIndex = _workbook.ActiveSheetIndex;
            _tabs.RefreshTabs();
            _dirty = true;
            UpdateTitle();
            UpdateSelectionUi();
        }

        private void InsertFunctionDialog()
        {
            using (FunctionDialog dlg = new FunctionDialog())
            {
                if (dlg.ShowDialog(this) == DialogResult.OK && !string.IsNullOrEmpty(dlg.SelectedFunction))
                    InsertFunctionText("=" + dlg.SelectedFunction + "(");
            }
        }

        private void InsertFunctionText(string prefix)
        {
            if (_grid.IsEditing) _grid.CommitEdit();
            SetFormulaText(prefix);
            _formulaBox.Focus();
            _formulaBox.SelectionStart = _formulaBox.TextLength;
        }

        private void ShowColorPalette(bool fill)
        {
            using (ColorPaletteForm cp = new ColorPaletteForm(fill ? "Color de relleno" : "Color de fuente", fill))
            {
                cp.StartPosition = FormStartPosition.Manual;
                Point p = Cursor.Position;
                Rectangle scr = Screen.FromPoint(p).WorkingArea;
                cp.Location = new Point(Math.Min(p.X, scr.Right - cp.Width - 4), Math.Min(p.Y, scr.Bottom - cp.Height - 4));
                if (cp.ShowDialog(this) == DialogResult.OK && cp.Accepted)
                {
                    Color color = cp.SelectedColor;
                    if (fill) _grid.ApplyStyle(delegate(CellStyle st) { st.FillColor = color; });
                    else _grid.ApplyStyle(delegate(CellStyle st) { st.FontColor = color.IsEmpty ? Color.Black : color; });
                }
            }
        }

        private void ShowBordersMenu()
        {
            ContextMenuStrip m = UiFactory.CreateMenu();
            m.Items.Add(UiFactory.Item("Todos los bordes", delegate { _grid.ApplyStyle(delegate(CellStyle s) { s.Borders = BorderFlags.All; }); }));
            m.Items.Add(UiFactory.Item("Borde inferior", delegate { _grid.ApplyStyle(delegate(CellStyle s) { s.Borders = BorderFlags.Bottom; }); }));
            m.Items.Add(UiFactory.Item("Borde superior", delegate { _grid.ApplyStyle(delegate(CellStyle s) { s.Borders = BorderFlags.Top; }); }));
            m.Items.Add(UiFactory.Separator());
            m.Items.Add(UiFactory.Item("Sin bordes", delegate { _grid.ApplyStyle(delegate(CellStyle s) { s.Borders = BorderFlags.None; }); }));
            m.Show(Cursor.Position);
        }

        private void ShowFormatCells()
        {
            using (FormatCellsDialog dlg = new FormatCellsDialog(_grid.GetActiveStyle()))
            {
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    CellStyle src = dlg.Result;
                    _grid.ApplyStyle(delegate(CellStyle dst) { CopyStyle(src, dst); });
                }
            }
        }

        private static void CopyStyle(CellStyle src, CellStyle dst)
        {
            dst.Bold = src.Bold;
            dst.Italic = src.Italic;
            dst.Underline = src.Underline;
            dst.Align = src.Align;
            dst.Format = src.Format;
            dst.Decimals = src.Decimals;
            dst.FontColor = src.FontColor;
            dst.FillColor = src.FillColor;
            dst.Borders = src.Borders;
            dst.FontSize = src.FontSize;
            dst.FontName = src.FontName;
        }

        private void ShowGridContextMenu()
        {
            ContextMenuStrip m = UiFactory.CreateMenu();
            if (_grid.Sheet != null && _grid.Sheet.IsProtected)
            {
                m.Items.Add(UiFactory.Item("Quitar protección de hoja...",
                    delegate { UnprotectSheet(); }));
                m.Items.Add(UiFactory.Separator());
            }
            if (_grid.HasSelectedImage)
            {
                m.Items.Add(UiFactory.Item("Recortar imagen...", delegate { CropSelectedImage(); }));
                m.Items.Add(UiFactory.Item("Eliminar imagen", delegate { _grid.DeleteSelectedImage(); }));
                m.Items.Add(UiFactory.Separator());
            }
            if (_grid.HasSelectedChart)
            {
                m.Items.Add(UiFactory.Item("Editar gráfico...", delegate { EditSelectedChart(); }));
                m.Items.Add(UiFactory.Item("Editar series...", delegate { EditChartSeries(); }));
                m.Items.Add(UiFactory.Item("Eliminar gráfico", delegate { _grid.DeleteSelectedChart(); }));
                m.Items.Add(UiFactory.Separator());
            }
            m.Items.Add(UiFactory.Item("Cortar", delegate { _grid.CopySelection(true); }));
            m.Items.Add(UiFactory.Item("Copiar", delegate { _grid.CopySelection(false); }));
            m.Items.Add(UiFactory.Item("Pegar", delegate { _grid.PasteClipboard(); }));
            m.Items.Add(UiFactory.Separator());
            m.Items.Add(UiFactory.Item("Insertar fila", delegate { _grid.InsertRow(); }));
            m.Items.Add(UiFactory.Item("Insertar columna", delegate { _grid.InsertColumn(); }));
            m.Items.Add(UiFactory.Item("Eliminar fila", delegate { _grid.DeleteRow(); }));
            m.Items.Add(UiFactory.Item("Eliminar columna", delegate { _grid.DeleteColumn(); }));
            m.Items.Add(UiFactory.Separator());
            m.Items.Add(UiFactory.Item("Borrar contenido", delegate { _grid.ClearSelectionContents(); }));
            m.Items.Add(UiFactory.Item("Combinar celdas", delegate { _grid.MergeSelection(); }));
            m.Items.Add(UiFactory.Item("Separar celdas", delegate { _grid.UnmergeSelection(); }));
            m.Items.Add(UiFactory.Item("Formato de celdas...", delegate { ShowFormatCells(); }));
            PivotRange pivot = PivotFormatting.Find(_grid.Sheet, _grid.ActiveRow, _grid.ActiveCol);
            if (pivot != null)
            {
                m.Items.Add(UiFactory.Item("Actualizar tabla dinámica",
                    delegate { RefreshPivot(pivot); }));
            }
            m.Items.Add(UiFactory.Separator());
            m.Items.Add(UiFactory.Item("Ordenar de A a Z", delegate { _grid.SortSelection(true); }));
            m.Items.Add(UiFactory.Item("Ordenar de Z a A", delegate { _grid.SortSelection(false); }));
            m.Show(Cursor.Position);
        }

        private void UnprotectSheet()
        {
            Sheet sheet = _grid.Sheet;
            if (sheet == null || !sheet.IsProtected) return;
            if (!string.IsNullOrEmpty(sheet.ProtectionPasswordHash))
            {
                string password = InputDialog.ShowPassword(this, "Quitar protección",
                    "Contraseña de la hoja:");
                if (password == null) return;
                if (!sheet.VerifyProtectionPassword(password))
                {
                    MessageBox.Show(this, "Contraseña incorrecta.", "Mini Excel",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
            sheet.IsProtected = false;
            _statusLabel.Text = "Protección de hoja quitada";
            _grid.Invalidate();
        }

        private void RefreshPivot(PivotRange pivot)
        {
            if (_grid.Sheet == null || pivot == null) return;
            if (_grid.Sheet.IsProtected)
            {
                MessageBox.Show(this, "La hoja está protegida. Quita la protección para actualizar la tabla.",
                    "Mini Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            _grid.BeginExternalEdit();
            if (!PivotEngine.Refresh(_grid.Sheet, pivot))
            {
                MessageBox.Show(this, "No hay datos de caché suficientes para actualizar esta tabla dinámica.",
                    "Mini Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            _grid.Invalidate();
            _statusLabel.Text = "Tabla dinámica actualizada";
        }

        private void ShowFind()
        {
            if (_findDialog == null || _findDialog.IsDisposed)
            {
                _findDialog = new FindDialog();
                _findDialog.SearchHandler = delegate(string text, bool matchCase, bool whole, bool formulas)
                {
                    int r, c;
                    if (_grid.FindNext(text, matchCase, whole, formulas, _grid.ActiveRow, _grid.ActiveCol, out r, out c))
                    {
                        _grid.SelectCell(r, c);
                        _grid.Focus();
                        return true;
                    }
                    return false;
                };
                _findDialog.ReplaceHandler = delegate(string find, string replace, bool matchCase,
                                                      bool whole, bool formulas)
                {
                    return _grid.ReplaceAll(find, replace, matchCase, whole, formulas);
                };
                _findDialog.Show(this);
            }
            else
            {
                _findDialog.Activate();
                _findDialog.FocusSearch();
            }
        }

        private void Execute(string cmd)
        {
            if (cmd == null) return;
            switch (cmd)
            {
                case "new": NewBook(true); break;
                case "print": ShowPrintDialog(); break;
                case "insertchart": InsertChart(); break;
                case "open":
                    using (OpenFileDialog dlg = new OpenFileDialog())
                    {
                        dlg.Title = "Abrir libro";
                        dlg.Filter = "Libros de Excel (*.xlsx;*.xls)|*.xlsx;*.xls|Libros de Excel modernos (*.xlsx)|*.xlsx|Libros de Excel 97-2003 (*.xls)|*.xls|Libros de Mini Excel (*.crb)|*.crb|Archivos CSV (*.csv)|*.csv|Todos los archivos (*.*)|*.*";
                        dlg.RestoreDirectory = true;
                        if (dlg.ShowDialog(this) == DialogResult.OK) OpenPath(dlg.FileName);
                    }
                    break;
                case "save": SaveBook(); break;
                case "saveas": SaveBookAs(); break;
                case "savestrict": SaveStrictAs(); break;
                case "exportcsv": ExportCsv(); break;
                case "exit": Close(); break;
                case "undo": Undo(); break;
                case "redo": Redo(); break;
                case "copy": _grid.CopySelection(false); break;
                case "cut": _grid.CopySelection(true); break;
                case "paste": _grid.PasteClipboard(); break;
                case "clear": _grid.ClearSelectionContents(); break;
                case "clearformat": _grid.ClearFormatting(); break;
                case "filter": ShowFilter(); break;
                case "validation": ShowValidation(); break;
                case "freeze": _grid.FreezePanesAtActive(); _dirty = true; UpdateTitle(); break;
                case "unfreeze": _grid.UnfreezePanes(); _dirty = true; UpdateTitle(); break;
                case "fontgrow":
                    _grid.ApplyStyle(delegate(CellStyle st)
                    { st.FontSize = Math.Min(72f, st.FontSize + 1f); });
                    break;
                case "fontshrink":
                    _grid.ApplyStyle(delegate(CellStyle st)
                    { st.FontSize = Math.Max(6f, st.FontSize - 1f); });
                    break;
                case "bold":
                    {
                        bool value = !_grid.GetActiveStyle().Bold;
                        _grid.ApplyStyle(delegate(CellStyle st) { st.Bold = value; });
                        break;
                    }
                case "italic":
                    {
                        bool value = !_grid.GetActiveStyle().Italic;
                        _grid.ApplyStyle(delegate(CellStyle st) { st.Italic = value; });
                        break;
                    }
                case "underline":
                    {
                        bool value = !_grid.GetActiveStyle().Underline;
                        _grid.ApplyStyle(delegate(CellStyle st) { st.Underline = value; });
                        break;
                    }
                case "alignleft": _grid.ApplyStyle(delegate(CellStyle st) { st.Align = 0; }); break;
                case "aligncenter": _grid.ApplyStyle(delegate(CellStyle st) { st.Align = 1; }); break;
                case "alignright": _grid.ApplyStyle(delegate(CellStyle st) { st.Align = 2; }); break;
                case "currency":
                    _grid.ApplyStyle(delegate(CellStyle st)
                    {
                        int decimals = st.Format == NumberFormatKind.General ? 0 : st.Decimals;
                        st.SetAccounting(decimals);
                    });
                    break;
                case "percent":
                    _grid.ApplyStyle(delegate(CellStyle st) { st.SetFormat(NumberFormatKind.Percent); st.Decimals = 0; });
                    break;
                case "comma":
                    _grid.ApplyStyle(delegate(CellStyle st) { st.SetFormat(NumberFormatKind.Number); st.Decimals = 2; });
                    break;
                case "datefmt":
                    _grid.ApplyStyle(delegate(CellStyle st) { st.SetFormat(NumberFormatKind.Date); });
                    break;
                case "incdec":
                    _grid.ApplyStyle(delegate(CellStyle st) { st.SetDecimals(st.Decimals + 1); });
                    break;
                case "decdec":
                    _grid.ApplyStyle(delegate(CellStyle st) { st.SetDecimals(st.Decimals - 1); });
                    break;
                case "fontcolor": ShowColorPalette(false); break;
                case "fillcolor": ShowColorPalette(true); break;
                case "borders": ShowBordersMenu(); break;
                case "formatcells": ShowFormatCells(); break;
                case "mergecells": _grid.MergeSelection(); break;
                case "unmergecells": _grid.UnmergeSelection(); break;
                case "autosum": _grid.AutoSum(); break;
                case "sortasc": _grid.SortSelection(true); break;
                case "sortdesc": _grid.SortSelection(false); break;
                case "insertrow": _grid.InsertRow(); break;
                case "insertcol": _grid.InsertColumn(); break;
                case "deleterow": _grid.DeleteRow(); break;
                case "deletecol": _grid.DeleteColumn(); break;
                case "find": ShowFind(); break;
                case "insertfunction": InsertFunctionDialog(); break;
                case "fnSUM": InsertFunctionText("=SUMA("); break;
                case "fnAVERAGE": InsertFunctionText("=PROMEDIO("); break;
                case "fnIF": InsertFunctionText("=SI("); break;
                case "fnMIN": InsertFunctionText("=MIN("); break;
                case "fnMAX": InsertFunctionText("=MAX("); break;
                case "fnCOUNT": InsertFunctionText("=CONTAR("); break;
                case "recalc":
                    _workbook.Recalculate();
                    _grid.Invalidate();
                    _statusLabel.Text = "Fórmulas recalculadas";
                    UpdateStats();
                    break;
                case "showformulas":
                    _grid.ShowFormulas = !_grid.ShowFormulas;
                    _grid.Invalidate();
                    UpdateRibbonState();
                    break;
                case "gridlines":
                    _grid.ShowGridLines = !_grid.ShowGridLines;
                    _grid.Invalidate();
                    UpdateRibbonState();
                    break;
                case "zoomin": _grid.SetZoom(_grid.Zoom + 0.1f); break;
                case "zoomout": _grid.SetZoom(_grid.Zoom - 0.1f); break;
                case "zoom100": _grid.SetZoom(1f); break;
                case "newsheet": AddSheet(); break;
                case "renamesheet": RenameSheet(_workbook.ActiveSheetIndex); break;
                case "deletesheet": DeleteSheet(_workbook.ActiveSheetIndex); break;
            }
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            Keys key = keyData & Keys.KeyCode;
            bool ctrl = (keyData & Keys.Control) != 0;
            bool shift = (keyData & Keys.Shift) != 0;
            bool alt = (keyData & Keys.Alt) != 0;
            bool textEditing = _formulaBox.Focused || _nameBox.Focused || _grid.IsEditing;

            if (alt && (key == Keys.Oemplus || key == Keys.Add))
            {
                Execute("autosum");
                return true;
            }
            if (key == Keys.F9)
            {
                Execute("recalc");
                return true;
            }
            if (key == Keys.Delete && _grid.HasSelectedImage)
            {
                _grid.DeleteSelectedImage();
                return true;
            }
            if (key == Keys.Delete && _grid.HasSelectedChart)
            {
                _grid.DeleteSelectedChart();
                return true;
            }
            if (textEditing)
            {
                if (ctrl && key == Keys.S)
                {
                    SaveBook();
                    return true;
                }
                return base.ProcessCmdKey(ref msg, keyData);
            }
            if (ctrl && !alt)
            {
                switch (key)
                {
                    case Keys.N: NewBook(true); return true;
                    case Keys.O: Execute("open"); return true;
                    case Keys.S:
                        if (shift) SaveBookAs(); else SaveBook();
                        return true;
                    case Keys.Z: Undo(); return true;
                    case Keys.Y: Redo(); return true;
                    case Keys.B: Execute("bold"); return true;
                    case Keys.I: Execute("italic"); return true;
                    case Keys.U: Execute("underline"); return true;
                    case Keys.P: ShowPrintDialog(); return true;
                    case Keys.C: Execute("copy"); return true;
                    case Keys.X: Execute("cut"); return true;
                    case Keys.V: Execute("paste"); return true;
                    case Keys.F: ShowFind(); return true;
                    case Keys.H: ShowFind(); return true;
                    case Keys.G:
                        _nameBox.Focus();
                        _nameBox.SelectAll();
                        return true;
                    case Keys.D1:
                        ShowFormatCells();
                        return true;
                }
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
    }
}
