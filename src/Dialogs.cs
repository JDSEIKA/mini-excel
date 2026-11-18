using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace CalcRapido
{
    public sealed class InputDialog : Form
    {
        private readonly TextBox _text;
        public string Value { get { return _text.Text; } }

        private InputDialog(string title, string prompt, string initial, bool password)
        {
            Text = title;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(390, 132);
            BackColor = Color.White;
            Font = Theme.Font(9f);

            Label lbl = new Label();
            lbl.Text = prompt;
            lbl.SetBounds(14, 14, 360, 20);
            Controls.Add(lbl);

            _text = new TextBox();
            _text.Text = initial ?? "";
            _text.SetBounds(14, 38, 360, 23);
            _text.Font = Theme.Font(10f);
            if (password) _text.UseSystemPasswordChar = true;
            Controls.Add(_text);

            Button ok = new Button();
            ok.Text = "Aceptar";
            ok.SetBounds(214, 82, 76, 28);
            ok.FlatStyle = FlatStyle.System;
            ok.Click += delegate { DialogResult = DialogResult.OK; Close(); };
            Controls.Add(ok);

            Button cancel = new Button();
            cancel.Text = "Cancelar";
            cancel.SetBounds(298, 82, 76, 28);
            cancel.FlatStyle = FlatStyle.System;
            cancel.DialogResult = DialogResult.Cancel;
            Controls.Add(cancel);

            AcceptButton = ok;
            CancelButton = cancel;
            Shown += delegate { _text.Focus(); _text.SelectAll(); };
        }

        public static string Show(IWin32Window owner, string title, string prompt, string initial)
        {
            return Show(owner, title, prompt, initial, false);
        }

        public static string ShowPassword(IWin32Window owner, string title, string prompt)
        {
            return Show(owner, title, prompt, "", true);
        }

        private static string Show(IWin32Window owner, string title, string prompt,
                                   string initial, bool password)
        {
            using (InputDialog d = new InputDialog(title, prompt, initial, password))
            {
                return d.ShowDialog(owner) == DialogResult.OK ? d.Value : null;
            }
        }
    }

    public sealed class FindDialog : Form
    {
        private readonly TextBox _text;
        private readonly TextBox _replace;
        private readonly CheckBox _matchCase;
        private readonly CheckBox _whole;
        private readonly CheckBox _formulas;
        private readonly Label _status;
        public Func<string, bool, bool, bool, bool> SearchHandler;
        public Func<string, string, bool, bool, bool, int> ReplaceHandler;

        public FindDialog()
        {
            Text = "Buscar";
            FormBorderStyle = FormBorderStyle.FixedToolWindow;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(390, 250);
            BackColor = Color.White;
            Font = Theme.Font(9f);
            KeyPreview = true;

            Label lbl = new Label();
            lbl.Text = "Buscar:";
            lbl.SetBounds(12, 14, 60, 20);
            Controls.Add(lbl);

            _text = new TextBox();
            _text.SetBounds(76, 12, 300, 23);
            _text.Font = Theme.Font(10f);
            Controls.Add(_text);

            Label replaceLabel = new Label();
            replaceLabel.Text = "Reemplazar:";
            replaceLabel.SetBounds(12, 43, 62, 20);
            Controls.Add(replaceLabel);

            _replace = new TextBox();
            _replace.SetBounds(76, 41, 300, 23);
            _replace.Font = Theme.Font(10f);
            Controls.Add(_replace);

            _matchCase = new CheckBox();
            _matchCase.Text = "Coincidir mayúsculas y minúsculas";
            _matchCase.SetBounds(76, 72, 260, 20);
            Controls.Add(_matchCase);

            _whole = new CheckBox();
            _whole.Text = "Solo celdas con el texto completo";
            _whole.SetBounds(76, 94, 260, 20);
            Controls.Add(_whole);

            _formulas = new CheckBox();
            _formulas.Text = "Buscar también en fórmulas";
            _formulas.Checked = true;
            _formulas.SetBounds(76, 116, 260, 20);
            Controls.Add(_formulas);

            _status = new Label();
            _status.Text = "";
            _status.ForeColor = Color.FromArgb(100, 100, 100);
            _status.SetBounds(12, 150, 364, 20);
            Controls.Add(_status);

            Button find = new Button();
            find.Text = "Buscar siguiente";
            find.SetBounds(76, 186, 118, 30);
            find.FlatStyle = FlatStyle.System;
            find.Click += delegate { DoSearch(); };
            Controls.Add(find);

            Button replace = new Button();
            replace.Text = "Reemplazar todo";
            replace.SetBounds(198, 186, 120, 30);
            replace.FlatStyle = FlatStyle.System;
            replace.Click += delegate { DoReplace(); };
            Controls.Add(replace);

            Button close = new Button();
            close.Text = "Cerrar";
            close.SetBounds(322, 186, 56, 30);
            close.FlatStyle = FlatStyle.System;
            close.Click += delegate { Close(); };
            Controls.Add(close);

            AcceptButton = find;
            CancelButton = close;
            KeyDown += delegate(object s, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Escape) Close();
            };
            Shown += delegate { _text.Focus(); _text.SelectAll(); };
        }

        public void FocusSearch()
        {
            _text.Focus();
            _text.SelectAll();
        }

        private void DoSearch()
        {
            if (string.IsNullOrEmpty(_text.Text))
            {
                _status.Text = "Escribe el texto que deseas buscar.";
                return;
            }
            bool found = false;
            if (SearchHandler != null)
                found = SearchHandler(_text.Text, _matchCase.Checked, _whole.Checked, _formulas.Checked);
            _status.Text = found ? "Coincidencia encontrada." : "No se encontró el texto.";
            _status.ForeColor = found ? Theme.Green : Color.FromArgb(160, 0, 0);
        }

        private void DoReplace()
        {
            if (string.IsNullOrEmpty(_text.Text))
            {
                _status.Text = "Escribe el texto que deseas reemplazar.";
                return;
            }
            if (ReplaceHandler == null)
            {
                _status.Text = "Reemplazar no está disponible.";
                return;
            }
            int count = ReplaceHandler(_text.Text, _replace.Text ?? "",
                _matchCase.Checked, _whole.Checked, _formulas.Checked);
            _status.Text = count == 1 ? "Se reemplazó 1 celda." :
                count > 0 ? "Se reemplazaron " + count + " celdas." : "No se encontró el texto.";
            _status.ForeColor = count > 0 ? Theme.Green : Color.FromArgb(160, 0, 0);
        }
    }

    public sealed class ValidationDialog : Form
    {
        private readonly TextBox _items;
        private readonly TextBox _range;
        private readonly TextBox _message;
        private readonly CheckBox _blank;
        public string[] Items
        {
            get
            {
                string[] raw = (_items.Text ?? "").Replace("\r", "").Split('\n');
                System.Collections.Generic.List<string> list = new System.Collections.Generic.List<string>();
                foreach (string s in raw) if (!string.IsNullOrEmpty(s.Trim())) list.Add(s.Trim());
                return list.ToArray();
            }
        }
        public string SourceRange { get { return (_range.Text ?? "").Trim(); } }
        public string Message { get { return _message.Text ?? ""; } }
        public bool AllowBlank { get { return _blank.Checked; } }

        public ValidationDialog()
        {
            Text = "Validación de datos";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false; MinimizeBox = false; ShowInTaskbar = false;
            ClientSize = new Size(400, 302);
            BackColor = Color.White;
            Font = new Font("Segoe UI", 9f);
            Label l = new Label(); l.Text = "Valores permitidos (uno por línea):";
            l.SetBounds(14, 12, 340, 20); Controls.Add(l);
            _items = new TextBox();
            _items.Multiline = true; _items.ScrollBars = ScrollBars.Vertical;
            _items.SetBounds(14, 34, 370, 100); Controls.Add(_items);
            Label lr = new Label(); lr.Text = "Rango de origen (opcional):";
            lr.SetBounds(14, 142, 160, 20); Controls.Add(lr);
            _range = new TextBox(); _range.SetBounds(178, 140, 206, 23); Controls.Add(_range);
            Label lm = new Label(); lm.Text = "Mensaje de error:";
            lm.SetBounds(14, 174, 120, 20); Controls.Add(lm);
            _message = new TextBox(); _message.Text = "El valor no está en la lista.";
            _message.SetBounds(134, 172, 250, 23); Controls.Add(_message);
            _blank = new CheckBox(); _blank.Text = "Permitir celda vacía"; _blank.Checked = true;
            _blank.SetBounds(14, 202, 180, 20); Controls.Add(_blank);
            Button ok = new Button(); ok.Text = "Aplicar"; ok.SetBounds(196, 250, 80, 28);
            ok.DialogResult = DialogResult.OK; Controls.Add(ok);
            Button cancel = new Button(); cancel.Text = "Cancelar"; cancel.SetBounds(282, 250, 80, 28);
            cancel.DialogResult = DialogResult.Cancel; Controls.Add(cancel);
            AcceptButton = ok; CancelButton = cancel;
        }
    }

    public sealed class CropImageDialog : Form
    {
        private readonly NumericUpDown _left, _top, _right, _bottom;
        public int LeftPercent { get { return (int)_left.Value; } }
        public int TopPercent { get { return (int)_top.Value; } }
        public int RightPercent { get { return (int)_right.Value; } }
        public int BottomPercent { get { return (int)_bottom.Value; } }

        public CropImageDialog()
        {
            Text = "Recortar imagen";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false; MinimizeBox = false; ShowInTaskbar = false;
            ClientSize = new Size(300, 210);
            BackColor = Color.White;
            Font = new Font("Segoe UI", 9f);
            Label title = new Label(); title.Text = "Porcentaje a recortar:";
            title.SetBounds(14, 12, 250, 20); Controls.Add(title);
            _left = AddPercent("Izquierda", 42);
            _top = AddPercent("Superior", 72);
            _right = AddPercent("Derecha", 102);
            _bottom = AddPercent("Inferior", 132);
            Button ok = new Button(); ok.Text = "Recortar"; ok.SetBounds(110, 170, 80, 28);
            ok.DialogResult = DialogResult.OK; Controls.Add(ok);
            Button cancel = new Button(); cancel.Text = "Cancelar"; cancel.SetBounds(196, 170, 80, 28);
            cancel.DialogResult = DialogResult.Cancel; Controls.Add(cancel);
            AcceptButton = ok; CancelButton = cancel;
        }

        private NumericUpDown AddPercent(string text, int y)
        {
            Label l = new Label(); l.Text = text; l.SetBounds(14, y + 3, 90, 20); Controls.Add(l);
            NumericUpDown n = new NumericUpDown();
            n.Minimum = 0; n.Maximum = 45; n.Value = 0;
            n.SetBounds(110, y, 60, 24); Controls.Add(n);
            Label p = new Label(); p.Text = "%"; p.SetBounds(174, y + 3, 20, 20); Controls.Add(p);
            return n;
        }
    }

    public sealed class FilterDialog : Form
    {
        private readonly ComboBox _operator;
        private readonly TextBox _text;
        public int ColumnIndex = -1;
        public int OperatorIndex { get { return _operator.SelectedIndex; } }
        public string FilterText { get { return _text.Text ?? ""; } }
        public bool ClearRequested;

        public FilterDialog(string columnName)
        {
            Text = "Autofiltro";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false; MinimizeBox = false; ShowInTaskbar = false;
            ClientSize = new Size(360, 170);
            BackColor = Color.White;
            Font = new Font("Segoe UI", 9f);
            Label lc = new Label(); lc.Text = "Columna: " + columnName;
            lc.SetBounds(14, 14, 330, 20); Controls.Add(lc);
            Label lo = new Label(); lo.Text = "Condición:";
            lo.SetBounds(14, 46, 78, 20); Controls.Add(lo);
            _operator = new ComboBox();
            _operator.DropDownStyle = ComboBoxStyle.DropDownList;
            _operator.Items.AddRange(new string[] { "Contiene", "Igual a", "No está vacío", "Está vacío" });
            _operator.SelectedIndex = 0;
            _operator.SetBounds(96, 43, 250, 24); Controls.Add(_operator);
            Label lt = new Label(); lt.Text = "Texto:";
            lt.SetBounds(14, 80, 78, 20); Controls.Add(lt);
            _text = new TextBox();
            _text.SetBounds(96, 78, 250, 23); Controls.Add(_text);
            Button apply = new Button(); apply.Text = "Aplicar"; apply.SetBounds(178, 126, 78, 28);
            apply.DialogResult = DialogResult.OK; Controls.Add(apply);
            Button clear = new Button(); clear.Text = "Quitar filtro"; clear.SetBounds(262, 126, 84, 28);
            clear.Click += delegate { ClearRequested = true; DialogResult = DialogResult.OK; };
            Controls.Add(clear);
            Button cancel = new Button(); cancel.Text = "Cancelar"; cancel.SetBounds(96, 126, 74, 28);
            cancel.DialogResult = DialogResult.Cancel; Controls.Add(cancel);
            AcceptButton = apply; CancelButton = cancel;
        }
    }

    public sealed class FormatCellsDialog : Form
    {
        private readonly CellStyle _style;
        private ComboBox _category;
        private NumericUpDown _decimals;
        private CheckBox _bold;
        private CheckBox _italic;
        private CheckBox _underline;
        private ComboBox _align;
        private NumericUpDown _fontSize;
        private ComboBox _borders;
        private Button _fontColor;
        private Button _fillColor;
        private Panel _preview;

        public CellStyle Result { get { return _style; } }

        public FormatCellsDialog(CellStyle initial)
        {
            _style = initial == null ? new CellStyle() : initial.Clone();
            Text = "Formato de celdas";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(394, 402);
            BackColor = Color.White;
            Font = Theme.Font(9f);

            Label title = new Label();
            title.Text = "Formato de celdas";
            title.Font = Theme.Font(13f, true);
            title.ForeColor = Theme.Green;
            title.SetBounds(14, 10, 340, 26);
            Controls.Add(title);

            Label catLbl = new Label();
            catLbl.Text = "Categoría:";
            catLbl.SetBounds(16, 52, 90, 20);
            Controls.Add(catLbl);

            _category = new ComboBox();
            _category.DropDownStyle = ComboBoxStyle.DropDownList;
            _category.SetBounds(112, 49, 240, 24);
            _category.Items.AddRange(new string[] { "General", "Número", "Moneda", "Porcentaje", "Fecha", "Hora", "Texto", "Personalizado" });
            _category.SelectedIndex = Math.Max(0, Math.Min(7, (int)_style.Format));
            _category.SelectedIndexChanged += delegate { UpdateControls(); };
            Controls.Add(_category);

            Label decLbl = new Label();
            decLbl.Text = "Decimales:";
            decLbl.SetBounds(16, 84, 90, 20);
            Controls.Add(decLbl);

            _decimals = new NumericUpDown();
            _decimals.SetBounds(112, 81, 60, 24);
            _decimals.Minimum = 0;
            _decimals.Maximum = 10;
            _decimals.Value = Math.Max(0, Math.Min(10, _style.Decimals));
            _decimals.ValueChanged += delegate { _style.SetDecimals((int)_decimals.Value); UpdatePreview(); };
            Controls.Add(_decimals);

            GroupBox fontBox = new GroupBox();
            fontBox.Text = "Fuente";
            fontBox.SetBounds(12, 118, 362, 82);
            Controls.Add(fontBox);

            _bold = new CheckBox();
            _bold.Text = "Negrita";
            _bold.SetBounds(14, 22, 80, 20);
            _bold.Checked = _style.Bold;
            _bold.CheckedChanged += delegate { _style.Bold = _bold.Checked; UpdatePreview(); };
            fontBox.Controls.Add(_bold);

            _italic = new CheckBox();
            _italic.Text = "Cursiva";
            _italic.SetBounds(100, 22, 80, 20);
            _italic.Checked = _style.Italic;
            _italic.CheckedChanged += delegate { _style.Italic = _italic.Checked; UpdatePreview(); };
            fontBox.Controls.Add(_italic);

            _underline = new CheckBox();
            _underline.Text = "Subrayado";
            _underline.SetBounds(184, 22, 94, 20);
            _underline.Checked = _style.Underline;
            _underline.CheckedChanged += delegate { _style.Underline = _underline.Checked; UpdatePreview(); };
            fontBox.Controls.Add(_underline);

            Label sizeLbl = new Label();
            sizeLbl.Text = "Tamaño:";
            sizeLbl.SetBounds(14, 50, 58, 20);
            fontBox.Controls.Add(sizeLbl);

            _fontSize = new NumericUpDown();
            _fontSize.SetBounds(76, 47, 60, 24);
            _fontSize.Minimum = 6;
            _fontSize.Maximum = 72;
            _fontSize.Value = Math.Max(6, Math.Min(72, (decimal)_style.FontSize));
            _fontSize.ValueChanged += delegate { _style.FontSize = (float)_fontSize.Value; UpdatePreview(); };
            fontBox.Controls.Add(_fontSize);

            _fontColor = new Button();
            _fontColor.Text = "Color de fuente";
            _fontColor.SetBounds(146, 46, 108, 26);
            _fontColor.FlatStyle = FlatStyle.System;
            _fontColor.Click += delegate { PickColor(false); };
            fontBox.Controls.Add(_fontColor);

            _fillColor = new Button();
            _fillColor.Text = "Relleno";
            _fillColor.SetBounds(260, 46, 88, 26);
            _fillColor.FlatStyle = FlatStyle.System;
            _fillColor.Click += delegate { PickColor(true); };
            fontBox.Controls.Add(_fillColor);

            Label alignLbl = new Label();
            alignLbl.Text = "Alineación:";
            alignLbl.SetBounds(16, 214, 90, 20);
            Controls.Add(alignLbl);

            _align = new ComboBox();
            _align.DropDownStyle = ComboBoxStyle.DropDownList;
            _align.SetBounds(112, 211, 120, 24);
            _align.Items.AddRange(new string[] { "General", "Izquierda", "Centro", "Derecha" });
            _align.SelectedIndex = Math.Max(0, Math.Min(3, _style.Align + 1));
            _align.SelectedIndexChanged += delegate { _style.Align = _align.SelectedIndex - 1; UpdatePreview(); };
            Controls.Add(_align);

            Label bordLbl = new Label();
            bordLbl.Text = "Bordes:";
            bordLbl.SetBounds(16, 246, 90, 20);
            Controls.Add(bordLbl);

            _borders = new ComboBox();
            _borders.DropDownStyle = ComboBoxStyle.DropDownList;
            _borders.SetBounds(112, 243, 180, 24);
            _borders.Items.AddRange(new string[] { "Ninguno", "Todos", "Inferior", "Superior" });
            _borders.SelectedIndex = 0;
            if (_style.Borders == BorderFlags.All) _borders.SelectedIndex = 1;
            else if (_style.Borders == BorderFlags.Bottom) _borders.SelectedIndex = 2;
            else if (_style.Borders == BorderFlags.Top) _borders.SelectedIndex = 3;
            _borders.SelectedIndexChanged += delegate
            {
                switch (_borders.SelectedIndex)
                {
                    case 1: _style.Borders = BorderFlags.All; break;
                    case 2: _style.Borders = BorderFlags.Bottom; break;
                    case 3: _style.Borders = BorderFlags.Top; break;
                    default: _style.Borders = BorderFlags.None; break;
                }
                UpdatePreview();
            };
            Controls.Add(_borders);

            Label previewLbl = new Label();
            previewLbl.Text = "Vista previa";
            previewLbl.SetBounds(16, 282, 100, 20);
            Controls.Add(previewLbl);

            _preview = new Panel();
            _preview.SetBounds(112, 278, 250, 44);
            _preview.BackColor = Color.White;
            _preview.BorderStyle = BorderStyle.FixedSingle;
            _preview.Paint += delegate(object s, PaintEventArgs e) { DrawPreview(e.Graphics); };
            Controls.Add(_preview);

            Button ok = new Button();
            ok.Text = "Aceptar";
            ok.SetBounds(206, 352, 82, 30);
            ok.FlatStyle = FlatStyle.System;
            ok.Click += delegate { DialogResult = DialogResult.OK; Close(); };
            Controls.Add(ok);

            Button cancel = new Button();
            cancel.Text = "Cancelar";
            cancel.SetBounds(294, 352, 82, 30);
            cancel.FlatStyle = FlatStyle.System;
            cancel.DialogResult = DialogResult.Cancel;
            Controls.Add(cancel);

            AcceptButton = ok;
            CancelButton = cancel;
            UpdateControls();
            UpdateColorButtons();
        }

        private void UpdateControls()
        {
            if (_category == null) return;
            NumberFormatKind oldFormat = _style.Format;
            NumberFormatKind format = (NumberFormatKind)_category.SelectedIndex;
            if (oldFormat != format && format != NumberFormatKind.Custom)
                _style.CustomFormat = null;
            if (format == NumberFormatKind.Custom && string.IsNullOrEmpty(_style.CustomFormat))
                _style.CustomFormat = _style.Decimals > 0
                    ? "#,##0." + new string('0', _style.Decimals)
                    : "#,##0";
            _style.Format = format;
            bool dec = _style.Format != NumberFormatKind.Date && _style.Format != NumberFormatKind.Time && _style.Format != NumberFormatKind.Text;
            _decimals.Enabled = dec;
            UpdatePreview();
        }

        private void UpdatePreview()
        {
            if (_preview != null) _preview.Invalidate();
        }

        private void PickColor(bool fill)
        {
            using (ColorPaletteForm cp = new ColorPaletteForm(fill ? "Color de relleno" : "Color de fuente", fill))
            {
                cp.StartPosition = FormStartPosition.Manual;
                Point p = Cursor.Position;
                Rectangle scr = Screen.FromPoint(p).WorkingArea;
                cp.Location = new Point(Math.Min(p.X, scr.Right - cp.Width - 4), Math.Min(p.Y, scr.Bottom - cp.Height - 4));
                if (cp.ShowDialog(this) == DialogResult.OK && cp.Accepted)
                {
                    if (fill) _style.FillColor = cp.SelectedColor;
                    else _style.FontColor = cp.SelectedColor.IsEmpty ? Color.Black : cp.SelectedColor;
                    UpdateColorButtons();
                    UpdatePreview();
                }
            }
        }

        private void UpdateColorButtons()
        {
            if (_fontColor == null) return;
            _fontColor.BackColor = _style.FontColor;
            _fontColor.ForeColor = _style.FontColor.GetBrightness() < 0.5f ? Color.White : Color.Black;
            _fillColor.BackColor = _style.FillColor.IsEmpty ? Color.White : _style.FillColor;
            _fillColor.ForeColor = _fillColor.BackColor.GetBrightness() < 0.5f ? Color.White : Color.Black;
        }

        private void DrawPreview(Graphics g)
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            if (!_style.FillColor.IsEmpty)
                using (SolidBrush b = new SolidBrush(_style.FillColor)) g.FillRectangle(b, _preview.ClientRectangle);
            FontStyle fs = FontStyle.Regular;
            if (_style.Bold) fs |= FontStyle.Bold;
            if (_style.Italic) fs |= FontStyle.Italic;
            if (_style.Underline) fs |= FontStyle.Underline;
            using (Font f = new Font(_style.FontName, Math.Max(7f, _style.FontSize), fs))
            {
                CellValue v = _style.Format == NumberFormatKind.Date || _style.Format == NumberFormatKind.Time
                    ? CellValue.FromNumber(DateTime.Today.ToOADate()) : CellValue.FromNumber(1234.5);
                string text = NumberFormatter.Format(v, _style);
                if (_style.Borders != BorderFlags.None)
                {
                    using (Pen p = new Pen(Color.FromArgb(110, 110, 110)))
                        g.DrawRectangle(p, 0, 0, _preview.ClientSize.Width - 1, _preview.ClientSize.Height - 1);
                }
                TextRenderer.DrawText(g, text, f, _preview.ClientRectangle, _style.FontColor,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
            }
        }
    }

    public sealed class FunctionDialog : Form
    {
        private readonly ListBox _list;
        private readonly TextBox _desc;
        private readonly Label _signature;
        public string SelectedFunction { get; private set; }

        public FunctionDialog()
        {
            Text = "Insertar función";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(620, 430);
            BackColor = Color.White;
            Font = Theme.Font(9f);

            Label title = new Label();
            title.Text = "Insertar función";
            title.Font = Theme.Font(13f, true);
            title.ForeColor = Theme.Green;
            title.SetBounds(14, 10, 300, 26);
            Controls.Add(title);

            _list = new ListBox();
            _list.SetBounds(14, 44, 290, 330);
            _list.Font = Theme.Font(9.5f);
            _list.IntegralHeight = false;
            foreach (FunctionInfo fi in FunctionCatalog.All) _list.Items.Add(fi);
            _list.SelectedIndexChanged += delegate { UpdateInfo(); };
            _list.DoubleClick += delegate { Accept(); };
            Controls.Add(_list);

            Label sigLbl = new Label();
            sigLbl.Text = "Sintaxis";
            sigLbl.SetBounds(318, 44, 100, 18);
            sigLbl.ForeColor = Color.FromArgb(100, 100, 100);
            Controls.Add(sigLbl);

            _signature = new Label();
            _signature.SetBounds(318, 62, 286, 42);
            _signature.Font = Theme.Font(9.5f, true);
            _signature.ForeColor = Theme.Green;
            Controls.Add(_signature);

            Label descLbl = new Label();
            descLbl.Text = "Descripción";
            descLbl.SetBounds(318, 112, 100, 18);
            descLbl.ForeColor = Color.FromArgb(100, 100, 100);
            Controls.Add(descLbl);

            _desc = new TextBox();
            _desc.SetBounds(318, 130, 286, 180);
            _desc.Multiline = true;
            _desc.ReadOnly = true;
            _desc.BorderStyle = BorderStyle.FixedSingle;
            _desc.BackColor = Color.FromArgb(250, 250, 250);
            _desc.ScrollBars = ScrollBars.Vertical;
            Controls.Add(_desc);

            Button insert = new Button();
            insert.Text = "Insertar";
            insert.SetBounds(430, 382, 82, 30);
            insert.FlatStyle = FlatStyle.System;
            insert.Click += delegate { Accept(); };
            Controls.Add(insert);

            Button cancel = new Button();
            cancel.Text = "Cancelar";
            cancel.SetBounds(518, 382, 82, 30);
            cancel.FlatStyle = FlatStyle.System;
            cancel.DialogResult = DialogResult.Cancel;
            Controls.Add(cancel);

            AcceptButton = insert;
            CancelButton = cancel;
            if (_list.Items.Count > 0) _list.SelectedIndex = 0;
        }

        private void UpdateInfo()
        {
            FunctionInfo fi = _list.SelectedItem as FunctionInfo;
            if (fi == null) return;
            _signature.Text = fi.Signature;
            _desc.Text = fi.Description + "\r\n\r\nCategoría: " + fi.Category;
        }

        private void Accept()
        {
            FunctionInfo fi = _list.SelectedItem as FunctionInfo;
            if (fi == null) return;
            SelectedFunction = fi.Name;
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
