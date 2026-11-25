using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace CalcRapido
{
    public static class UiFactory
    {
        public static ContextMenuStrip CreateMenu()
        {
            ContextMenuStrip m = new ContextMenuStrip();
            m.Renderer = new ToolStripProfessionalRenderer(new OfficeColorTable());
            m.ShowImageMargin = false;
            m.ShowCheckMargin = false;
            m.Font = Theme.Font(9f);
            m.BackColor = Color.White;
            m.Padding = new Padding(2);
            return m;
        }

        public static ToolStripMenuItem Item(string text, Keys shortcut, EventHandler click)
        {
            ToolStripMenuItem item = new ToolStripMenuItem(text);
            item.ShortcutKeys = shortcut;
            item.ShowShortcutKeys = shortcut != Keys.None;
            item.Font = Theme.Font(9f);
            item.ForeColor = Theme.FontDefault;
            if (click != null) item.Click += click;
            return item;
        }

        public static ToolStripMenuItem Item(string text, EventHandler click)
        {
            return Item(text, Keys.None, click);
        }

        public static ToolStripSeparator Separator()
        {
            return new ToolStripSeparator();
        }
    }

    public sealed class RibbonButton : Control
    {
        public string Command;
        public string Label;
        public IconKind Icon;
        public bool Large;
        public bool Checkable;
        public bool Checked;
        public bool DropDown;
        public bool LabelBeside;
        public string TextIcon;

        private bool _hover;
        private bool _pressed;

        public RibbonButton(string command, string label, IconKind icon, bool large)
        {
            Command = command;
            Label = label;
            Icon = icon;
            Large = large;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw |
                     ControlStyles.SupportsTransparentBackColor, true);
            BackColor = Color.White;
            Cursor = Cursors.Hand;
            TabStop = false;
            AccessibleName = label;
            if (large)
            {
                Size = new Size(54, 67);
                Font = Theme.Font(7.2f);
            }
            else
            {
                Size = new Size(26, 21);
                Font = Theme.Font(8f);
            }
        }

        protected override void OnMouseEnter(EventArgs e)
        {
            base.OnMouseEnter(e);
            _hover = true;
            Invalidate();
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            _hover = false;
            _pressed = false;
            Invalidate();
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            _pressed = true;
            Invalidate();
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            _pressed = false;
            Invalidate();
        }

        protected override void OnEnabledChanged(EventArgs e)
        {
            base.OnEnabledChanged(e);
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            Color back = Color.White;
            if (!Enabled) back = Color.FromArgb(250, 250, 250);
            else if (_pressed) back = Theme.Pressed;
            else if (Checked) back = Color.FromArgb(220, 239, 228);
            else if (_hover) back = Theme.Hover;
            using (SolidBrush b = new SolidBrush(back)) g.FillRectangle(b, ClientRectangle);
            if (Checked)
            {
                using (Pen p = new Pen(Color.FromArgb(170, 205, 185))) g.DrawRectangle(p, 0, 0, Width - 1, Height - 1);
            }
            else if (_hover)
            {
                using (Pen p = new Pen(Color.FromArgb(200, 224, 210))) g.DrawRectangle(p, 0, 0, Width - 1, Height - 1);
            }

            Color color = Enabled ? Color.FromArgb(55, 55, 55) : Color.FromArgb(165, 165, 165);
            if (Large)
            {
                Rectangle ir = new Rectangle((Width - 24) / 2, 4, 24, 24);
                IconPainter.Draw(g, Icon, ir, color, true);
                Rectangle tr = new Rectangle(1, 30, Width - 2, Height - 31);
                TextRenderer.DrawText(g, Label, Font, tr, color,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.Top | TextFormatFlags.WordBreak |
                    TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
            }
            else if (LabelBeside)
            {
                Rectangle ir = new Rectangle(3, (Height - 17) / 2, 17, 17);
                if (!string.IsNullOrEmpty(TextIcon))
                {
                    using (Font f = new Font("Segoe UI", 8.5f))
                    using (StringFormat sf = new StringFormat())
                    {
                        sf.Alignment = StringAlignment.Center;
                        sf.LineAlignment = StringAlignment.Center;
                        using (SolidBrush br = new SolidBrush(color))
                            g.DrawString(TextIcon, f, br, ir, sf);
                    }
                }
                else IconPainter.Draw(g, Icon, ir, color, false);
                Rectangle tr = new Rectangle(24, 0, Width - 27, Height);
                TextRenderer.DrawText(g, Label, Font, tr, color,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter |
                    TextFormatFlags.EndEllipsis | TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix);
            }
            else
            {
                Rectangle ir = new Rectangle((Width - 17) / 2, (Height - 17) / 2, 17, 17);
                if (!string.IsNullOrEmpty(TextIcon))
                {
                    using (Font f = new Font("Segoe UI", 8.5f, TextIcon == "B" ? FontStyle.Bold : FontStyle.Regular))
                    using (StringFormat sf = new StringFormat())
                    {
                        sf.Alignment = StringAlignment.Center;
                        sf.LineAlignment = StringAlignment.Center;
                        using (SolidBrush br = new SolidBrush(color))
                            g.DrawString(TextIcon, f, br, new RectangleF(0, 0, Width, Height), sf);
                    }
                }
                else IconPainter.Draw(g, Icon, ir, color, false);
            }
            if (DropDown)
            {
                PointF[] pts = new PointF[]
                {
                    new PointF(Width - 7, Height - 6),
                    new PointF(Width - 2, Height - 6),
                    new PointF(Width - 4.5f, Height - 3)
                };
                using (SolidBrush br = new SolidBrush(color)) g.FillPolygon(br, pts);
            }
        }
    }

    public sealed class RibbonGroup : Panel
    {
        public string Caption;
        private int _maxCols;
        private readonly List<Control> _items = new List<Control>();

        public RibbonGroup(string caption)
        {
            Caption = caption;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            BackColor = Color.White;
            Height = 89;
            Width = 60;
        }

        public void AddItem(Control c, int col, int row, int colSpan, int rowSpan)
        {
            c.Tag = new int[] { col, row, colSpan, rowSpan };
            _items.Add(c);
            Controls.Add(c);
            if (col + colSpan > _maxCols) _maxCols = col + colSpan;
            Width = Math.Max(40, _maxCols * 28 + 7);
            PerformLayout();
        }

        public void Finish()
        {
            Width = Math.Max(40, _maxCols * 28 + 7);
            PerformLayout();
        }

        protected override void OnLayout(LayoutEventArgs levent)
        {
            base.OnLayout(levent);
            foreach (Control c in _items)
            {
                int[] p = c.Tag as int[];
                if (p == null) continue;
                c.Left = 3 + p[0] * 28;
                c.Top = 3 + p[1] * 23;
                c.Width = p[2] * 28 - 2;
                c.Height = p[3] * 23 - 2;
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            using (SolidBrush b = new SolidBrush(Color.White)) e.Graphics.FillRectangle(b, ClientRectangle);
            using (Pen p = new Pen(Color.FromArgb(226, 226, 226))) e.Graphics.DrawLine(p, Width - 1, 6, Width - 1, Height - 6);
            using (Font f = Theme.Font(7.4f))
            {
                TextRenderer.DrawText(e.Graphics, Caption, f,
                    new Rectangle(6, Height - 15, Width - 10, 14), Color.FromArgb(100, 100, 100),
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
            }
        }
    }

    public sealed class RibbonTabStrip : Control
    {
        public string[] TabNames = new string[] { "Inicio", "Insertar", "Fórmulas", "Datos", "Vista" };
        public int ActiveIndex;
        public bool SaveEnabled = true;
        public bool UndoEnabled;
        public bool RedoEnabled;
        public event EventHandler TabChanged;
        public event EventHandler FileClicked;
        public event Action<string> QatCommand;

        private Rectangle _fileRect;
        private readonly List<Rectangle> _tabRects = new List<Rectangle>();
        private readonly List<Rectangle> _qatRects = new List<Rectangle>();
        private int _hoverTab = -1;
        private int _hoverQat = -1;
        private bool _hoverFile;

        public RibbonTabStrip()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            Height = 30;
            BackColor = Color.FromArgb(243, 243, 243);
            Cursor = Cursors.Default;
        }

        private void LayoutRects()
        {
            _tabRects.Clear();
            _qatRects.Clear();
            _fileRect = new Rectangle(0, 0, 64, Height);
            int x = 72;
            using (Graphics g = CreateGraphics())
            using (Font f = Theme.Font(9.2f))
            {
                for (int i = 0; i < TabNames.Length; i++)
                {
                    Size sz = TextRenderer.MeasureText(g, TabNames[i], f);
                    int w = Math.Max(58, sz.Width + 24);
                    _tabRects.Add(new Rectangle(x, 1, w, Height - 2));
                    x += w;
                }
            }
            int qx = Width - 3 * 27 - 8;
            for (int i = 0; i < 3; i++) _qatRects.Add(new Rectangle(qx + i * 27, 4, 25, 23));
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            LayoutRects();
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            int tab = -1;
            for (int i = 0; i < _tabRects.Count; i++) if (_tabRects[i].Contains(e.Location)) { tab = i; break; }
            int qat = -1;
            for (int i = 0; i < _qatRects.Count; i++) if (_qatRects[i].Contains(e.Location)) { qat = i; break; }
            bool file = _fileRect.Contains(e.Location);
            if (tab != _hoverTab || qat != _hoverQat || file != _hoverFile)
            {
                _hoverTab = tab; _hoverQat = qat; _hoverFile = file;
                Cursor = (tab >= 0 || qat >= 0 || file) ? Cursors.Hand : Cursors.Default;
                Invalidate();
            }
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            _hoverTab = -1; _hoverQat = -1; _hoverFile = false;
            Cursor = Cursors.Default;
            Invalidate();
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            LayoutRects();
            if (_fileRect.Contains(e.Location))
            {
                if (FileClicked != null) FileClicked(this, EventArgs.Empty);
                return;
            }
            for (int i = 0; i < _tabRects.Count; i++)
            {
                if (_tabRects[i].Contains(e.Location))
                {
                    if (ActiveIndex != i)
                    {
                        ActiveIndex = i;
                        Invalidate();
                        if (TabChanged != null) TabChanged(this, EventArgs.Empty);
                    }
                    return;
                }
            }
            for (int i = 0; i < _qatRects.Count; i++)
            {
                if (_qatRects[i].Contains(e.Location))
                {
                    string[] cmds = new string[] { "save", "undo", "redo" };
                    if (QatCommand != null) QatCommand(cmds[i]);
                    return;
                }
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            if (_tabRects.Count == 0) LayoutRects();
            using (SolidBrush b = new SolidBrush(Color.FromArgb(243, 243, 243))) g.FillRectangle(b, ClientRectangle);
            using (Pen p = new Pen(Color.FromArgb(214, 214, 214))) g.DrawLine(p, 0, Height - 1, Width, Height - 1);

            using (SolidBrush b = new SolidBrush(Theme.Green)) g.FillRectangle(b, _fileRect);
            using (SolidBrush b = new SolidBrush(_hoverFile ? Color.FromArgb(46, 139, 87) : Theme.Green)) g.FillRectangle(b, _fileRect);
            using (Font fileFont = Theme.Font(9f))
            {
                TextRenderer.DrawText(g, "Archivo", fileFont, _fileRect, Color.White,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
            }

            for (int i = 0; i < _tabRects.Count; i++)
            {
                Rectangle r = _tabRects[i];
                if (i == ActiveIndex)
                {
                    using (SolidBrush b = new SolidBrush(Color.White)) g.FillRectangle(b, r);
                    using (Pen p = new Pen(Theme.Green, 2f)) g.DrawLine(p, r.Left + 4, r.Top, r.Right - 4, r.Top);
                }
                else if (i == _hoverTab)
                {
                    using (SolidBrush b = new SolidBrush(Theme.Hover)) g.FillRectangle(b, r);
                }
                Color text = i == ActiveIndex ? Theme.Green : Color.FromArgb(70, 70, 70);
                using (Font tf = Theme.Font(9f, i == ActiveIndex))
                {
                    TextRenderer.DrawText(g, TabNames[i], tf, r, text,
                        TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
                }
            }

            using (Pen sep = new Pen(Color.FromArgb(220, 220, 220))) g.DrawLine(sep, _qatRects[0].Left - 8, 7, _qatRects[0].Left - 8, Height - 8);
            string[] qcmds = new string[] { "save", "undo", "redo" };
            bool[] qenabled = new bool[] { SaveEnabled, UndoEnabled, RedoEnabled };
            IconKind[] qicons = new IconKind[] { IconKind.Save, IconKind.Undo, IconKind.Redo };
            for (int i = 0; i < _qatRects.Count; i++)
            {
                Rectangle r = _qatRects[i];
                if (_hoverQat == i && qenabled[i])
                {
                    using (SolidBrush b = new SolidBrush(Theme.Hover)) g.FillRectangle(b, r);
                    using (Pen p = new Pen(Color.FromArgb(200, 224, 210))) g.DrawRectangle(p, r);
                }
                Color c = qenabled[i] ? Color.FromArgb(80, 80, 80) : Color.FromArgb(185, 185, 185);
                IconPainter.Draw(g, qicons[i], new Rectangle(r.X + 4, r.Y + 3, 17, 17), c, false);
            }
        }
    }

    public sealed class RibbonBar : UserControl
    {
        private readonly RibbonTabStrip _tabs;
        private readonly Panel _content;
        private readonly List<RibbonGroup>[] _tabGroups;
        private readonly Dictionary<string, List<RibbonButton>> _buttons = new Dictionary<string, List<RibbonButton>>();
        private readonly ToolTip _tooltip = new ToolTip();
        private ComboBox _fontSize;
        private ComboBox _fontName;
        private bool _updatingFont;
        private bool _updatingFontName;
        private int _activeTab;

        public event Action<string> Command;
        public event Action<float> FontSizeSelected;
        public event Action<string> FontNameSelected;
        public event Action<string> RecentFileSelected;
        private readonly List<string> _recentFiles = new List<string>();

        public void SetRecentFiles(IEnumerable<string> files)
        {
            _recentFiles.Clear();
            if (files != null)
                foreach (string path in files)
                    if (!string.IsNullOrEmpty(path)) _recentFiles.Add(path);
        }

        public RibbonBar()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
            BackColor = Color.White;
            Height = 120;
            _tabs = new RibbonTabStrip();
            _tabs.Dock = DockStyle.Top;
            _tabs.TabChanged += delegate { SetTab(_tabs.ActiveIndex); };
            _tabs.FileClicked += delegate { ShowFileMenu(); };
            _tabs.QatCommand += delegate(string cmd) { Raise(cmd); };
            _content = new Panel();
            _content.Dock = DockStyle.Fill;
            _content.BackColor = Color.White;
            Controls.Add(_content);
            Controls.Add(_tabs);
            _tabGroups = new List<RibbonGroup>[_tabs.TabNames.Length];
            BuildTabs();
            SetTab(0);
        }

        public RibbonTabStrip Tabs { get { return _tabs; } }

        public void SetQuickState(bool canUndo, bool canRedo)
        {
            _tabs.UndoEnabled = canUndo;
            _tabs.RedoEnabled = canRedo;
            _tabs.Invalidate();
        }

        private RibbonGroup NewGroup(List<RibbonGroup> list, string caption)
        {
            RibbonGroup g = new RibbonGroup(caption);
            list.Add(g);
            _content.Controls.Add(g);
            return g;
        }

        private RibbonButton AddButton(RibbonGroup g, string cmd, string label, IconKind icon, bool large,
                                       int col, int row, int colSpan, int rowSpan, bool checkable = false,
                                       bool dropDown = false, string textIcon = null, bool labelBeside = false)
        {
            RibbonButton b = new RibbonButton(cmd, label, icon, large);
            b.Checkable = checkable;
            b.DropDown = dropDown;
            b.TextIcon = textIcon;
            b.LabelBeside = labelBeside;
            b.Click += delegate { Raise(cmd); };
            g.AddItem(b, col, row, colSpan, rowSpan);
            List<RibbonButton> list;
            if (!_buttons.TryGetValue(cmd, out list)) { list = new List<RibbonButton>(); _buttons[cmd] = list; }
            list.Add(b);
            _tooltip.SetToolTip(b, TipFor(cmd, label));
            return b;
        }

        private static string TipFor(string cmd, string label)
        {
            switch (cmd)
            {
                case "paste": return "Pegar (Ctrl+V)";
                case "cut": return "Cortar (Ctrl+X)";
                case "copy": return "Copiar (Ctrl+C)";
                case "clear": return "Borrar contenido (Supr)";
                case "fontgrow": return "Aumentar tamaño de fuente";
                case "fontshrink": return "Disminuir tamaño de fuente";
                case "bold": return "Negrita (Ctrl+B)";
                case "italic": return "Cursiva (Ctrl+I)";
                case "underline": return "Subrayado (Ctrl+U)";
                case "formatcells": return "Formato de celdas (Ctrl+1)";
                case "find": return "Buscar (Ctrl+F)";
                case "autosum": return "Autosuma (Alt+=)";
                case "recalc": return "Recalcular (F9)";
                case "save": return "Guardar (Ctrl+S)";
                case "undo": return "Deshacer (Ctrl+Z)";
                case "redo": return "Rehacer (Ctrl+Y)";
                default: return label;
            }
        }

        private void AddFontNameCombo(RibbonGroup g, int col, int row, int colSpan, int rowSpan)
        {
            _fontName = new ComboBox();
            _fontName.DropDownStyle = ComboBoxStyle.DropDownList;
            _fontName.FlatStyle = FlatStyle.Flat;
            _fontName.Font = Theme.Font(8.5f);
            _fontName.BackColor = Color.White;
            List<string> families = new List<string>();
            foreach (FontFamily family in FontFamily.Families)
            {
                try
                {
                    if (!string.IsNullOrEmpty(family.Name)) families.Add(family.Name);
                }
                catch { }
            }
            families.Sort(StringComparer.CurrentCultureIgnoreCase);
            _fontName.Items.AddRange(families.ToArray());
            int idx = _fontName.Items.IndexOf("Segoe UI");
            if (idx >= 0) _fontName.SelectedIndex = idx;
            _fontName.SelectedIndexChanged += delegate
            {
                if (_updatingFontName || _fontName.SelectedItem == null) return;
                if (FontNameSelected != null) FontNameSelected(_fontName.SelectedItem.ToString());
            };
            g.AddItem(_fontName, col, row, colSpan, rowSpan);
            _tooltip.SetToolTip(_fontName, "Tipo de fuente");
        }

        private void AddFontCombo(RibbonGroup g, int col, int row, int colSpan, int rowSpan)
        {
            _fontSize = new ComboBox();
            _fontSize.DropDownStyle = ComboBoxStyle.DropDownList;
            _fontSize.FlatStyle = FlatStyle.Flat;
            _fontSize.Font = Theme.Font(9f);
            _fontSize.BackColor = Color.White;
            string[] sizes = new string[] { "8", "9", "10", "11", "12", "14", "16", "18", "20", "24", "28", "36" };
            _fontSize.Items.AddRange(sizes);
            _fontSize.SelectedItem = "10";
            _fontSize.SelectedIndexChanged += delegate
            {
                if (_updatingFont || _fontSize.SelectedItem == null) return;
                float f;
                if (float.TryParse(_fontSize.SelectedItem.ToString(), NumberStyles.Float, CultureInfo.GetCultureInfo("es-CO"), out f) ||
                    float.TryParse(_fontSize.SelectedItem.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out f))
                {
                    if (FontSizeSelected != null) FontSizeSelected(f);
                }
            };
            g.AddItem(_fontSize, col, row, colSpan, rowSpan);
            _tooltip.SetToolTip(_fontSize, "Tamaño de fuente");
        }

        private void BuildTabs()
        {
            List<RibbonGroup> home = new List<RibbonGroup>();
            _tabGroups[0] = home;
            RibbonGroup clip = NewGroup(home, "Portapapeles");
            AddButton(clip, "paste", "Pegar", IconKind.Paste, true, 0, 0, 2, 3);
            AddButton(clip, "cut", "Cortar", IconKind.Cut, false, 2, 0, 3, 1,
                false, false, null, true);
            AddButton(clip, "copy", "Copiar", IconKind.Copy, false, 2, 1, 3, 1,
                false, false, null, true);
            AddButton(clip, "clear", "Borrar", IconKind.Close, false, 2, 2, 3, 1,
                false, false, null, true);
            clip.Finish();

            RibbonGroup font = NewGroup(home, "Fuente");
            AddFontNameCombo(font, 0, 0, 3, 1);
            AddFontCombo(font, 3, 0, 2, 1);
            AddButton(font, "fontgrow", "Aumentar fuente", IconKind.None, false, 0, 1, 2, 1,
                false, false, "A+");
            AddButton(font, "fontshrink", "Disminuir fuente", IconKind.None, false, 2, 1, 2, 1,
                false, false, "A-");
            AddButton(font, "undo", "Deshacer", IconKind.Undo, false, 0, 2, 3, 1,
                false, false, null, true);
            AddButton(font, "redo", "Rehacer", IconKind.Redo, false, 3, 2, 3, 1,
                false, false, null, true);
            font.Finish();
            RibbonGroup style = NewGroup(home, "Estilo");
            AddButton(style, "bold", "Negrita (Ctrl+B)", IconKind.Bold, true, 0, 0, 2, 3, true);
            AddButton(style, "italic", "Cursiva (Ctrl+I)", IconKind.Italic, true, 2, 0, 2, 3, true);
            AddButton(style, "underline", "Subrayado (Ctrl+U)", IconKind.Underline, true, 4, 0, 2, 3, true);
            AddButton(style, "fontcolor", "Color de fuente", IconKind.FontColor, true, 6, 0, 2, 3, false, true);
            AddButton(style, "fillcolor", "Color de relleno", IconKind.FillColor, true, 8, 0, 2, 3, false, true);
            AddButton(style, "borders", "Bordes", IconKind.Borders, true, 10, 0, 2, 3, false, true);
            AddButton(style, "clearformat", "Borrar formato", IconKind.Close, true, 12, 0, 2, 3);
            style.Finish();

            RibbonGroup align = NewGroup(home, "Alineación");
            AddButton(align, "alignleft", "Alinear a la izquierda", IconKind.AlignLeft, false, 0, 0, 1, 1, true);
            AddButton(align, "aligncenter", "Centrar", IconKind.AlignCenter, false, 1, 0, 1, 1, true);
            AddButton(align, "alignright", "Alinear a la derecha", IconKind.AlignRight, false, 2, 0, 1, 1, true);
            AddButton(align, "formatcells", "Formato de celdas (Ctrl+1)", IconKind.Table, false, 0, 1, 3, 1);
            AddButton(align, "mergecells", "Combinar celdas", IconKind.Table, false, 0, 2, 1, 1);
            AddButton(align, "unmergecells", "Separar celdas", IconKind.Grid, false, 1, 2, 1, 1);
            align.Finish();

            RibbonGroup number = NewGroup(home, "Número");
            AddButton(number, "currency", "Formato de moneda", IconKind.Currency, false, 0, 0, 1, 1);
            AddButton(number, "percent", "Formato de porcentaje", IconKind.Percent, false, 1, 0, 1, 1);
            AddButton(number, "comma", "Separador de miles", IconKind.Comma, false, 2, 0, 1, 1);
            AddButton(number, "incdec", "Aumentar decimales", IconKind.IncDecimals, false, 0, 1, 1, 1);
            AddButton(number, "decdec", "Disminuir decimales", IconKind.DecDecimals, false, 1, 1, 1, 1);
            AddButton(number, "datefmt", "Formato de fecha", IconKind.Date, false, 2, 1, 1, 1);
            number.Finish();

            RibbonGroup cells = NewGroup(home, "Celdas");
            AddButton(cells, "insertrow", "Insertar fila", IconKind.InsertRow, false, 0, 0, 1, 1);
            AddButton(cells, "insertcol", "Insertar columna", IconKind.InsertCol, false, 1, 0, 1, 1);
            AddButton(cells, "deleterow", "Eliminar fila", IconKind.DeleteRow, false, 0, 1, 1, 1);
            AddButton(cells, "deletecol", "Eliminar columna", IconKind.DeleteCol, false, 1, 1, 1, 1);
            cells.Finish();

            RibbonGroup styles = NewGroup(home, "Estilos");
            AddButton(styles, "formatcells", "Formato", IconKind.Table, true, 0, 0, 2, 3);
            AddButton(styles, "clearformat", "Borrar formato", IconKind.Close, false, 2, 0, 1, 1);
            AddButton(styles, "datefmt", "Fecha", IconKind.Date, false, 2, 1, 1, 1);
            styles.Finish();

            RibbonGroup edit = NewGroup(home, "Edición");
            AddButton(edit, "autosum", "Autosuma", IconKind.Sum, true, 0, 0, 2, 3);
            AddButton(edit, "find", "Buscar (Ctrl+F)", IconKind.Find, false, 2, 0, 1, 1);
            edit.Finish();

            List<RibbonGroup> insert = new List<RibbonGroup>();
            _tabGroups[1] = insert;
            RibbonGroup insCells = NewGroup(insert, "Celdas");
            AddButton(insCells, "insertrow", "Insertar fila", IconKind.InsertRow, true, 0, 0, 2, 3);
            AddButton(insCells, "insertcol", "Insertar columna", IconKind.InsertCol, true, 2, 0, 2, 3);
            AddButton(insCells, "deleterow", "Eliminar fila", IconKind.DeleteRow, true, 4, 0, 2, 3);
            AddButton(insCells, "deletecol", "Eliminar columna", IconKind.DeleteCol, true, 6, 0, 2, 3);
            insCells.Finish();
            RibbonGroup sheets = NewGroup(insert, "Hojas");
            AddButton(sheets, "newsheet", "Nueva hoja", IconKind.NewSheet, true, 0, 0, 2, 3);
            AddButton(sheets, "renamesheet", "Cambiar nombre", IconKind.Rename, true, 2, 0, 2, 3);
            AddButton(sheets, "deletesheet", "Eliminar hoja", IconKind.Close, true, 4, 0, 2, 3);
            sheets.Finish();
            RibbonGroup insertData = NewGroup(insert, "Datos");
            AddButton(insertData, "find", "Buscar", IconKind.Find, true, 0, 0, 2, 3);
            insertData.Finish();
            RibbonGroup insertCharts = NewGroup(insert, "Gráficos");
            AddButton(insertCharts, "insertchart", "Insertar gráfico", IconKind.Chart, true, 0, 0, 2, 3);
            insertCharts.Finish();

            List<RibbonGroup> formulas = new List<RibbonGroup>();
            _tabGroups[2] = formulas;
            RibbonGroup library = NewGroup(formulas, "Biblioteca de funciones");
            AddButton(library, "insertfunction", "Insertar función", IconKind.Function, true, 0, 0, 2, 3);
            AddButton(library, "fnSUM", "SUMA", IconKind.Sum, false, 2, 0, 3, 1,
                false, false, "\u03A3", true);
            AddButton(library, "fnAVERAGE", "PROMEDIO", IconKind.None, false, 5, 0, 4, 1,
                false, false, "x\u0304", true);
            AddButton(library, "fnIF", "SI", IconKind.None, false, 2, 1, 3, 1,
                false, false, "\u0192", true);
            AddButton(library, "fnMIN", "MIN", IconKind.None, false, 5, 1, 4, 1,
                false, false, "m", true);
            AddButton(library, "fnMAX", "MAX", IconKind.None, false, 2, 2, 3, 1,
                false, false, "M", true);
            AddButton(library, "fnCOUNT", "CONTAR", IconKind.None, false, 5, 2, 4, 1,
                false, false, "n", true);
            library.Finish();
            RibbonGroup calc = NewGroup(formulas, "Cálculo");
            AddButton(calc, "recalc", "Recalcular", IconKind.Recalculate, true, 0, 0, 2, 3);
            AddButton(calc, "showformulas", "Mostrar fórmulas", IconKind.ShowFormulas, false, 2, 0, 1, 1, true);
            AddButton(calc, "autosum", "Autosuma", IconKind.Sum, false, 2, 1, 1, 1);
            calc.Finish();

            List<RibbonGroup> data = new List<RibbonGroup>();
            _tabGroups[3] = data;
            RibbonGroup sort = NewGroup(data, "Ordenar y filtrar");
            AddButton(sort, "sortasc", "Ordenar A-Z", IconKind.SortAsc, true, 0, 0, 2, 3);
            AddButton(sort, "sortdesc", "Ordenar Z-A", IconKind.SortDesc, true, 2, 0, 2, 3);
            AddButton(sort, "find", "Buscar", IconKind.Find, true, 4, 0, 2, 3);
            sort.Finish();
            RibbonGroup filter = NewGroup(data, "Filtro");
            AddButton(filter, "filter", "Autofiltro", IconKind.Filter, true, 0, 0, 2, 3);
            filter.Finish();
            RibbonGroup validation = NewGroup(data, "Validación");
            AddButton(validation, "validation", "Validación de datos", IconKind.Table, true, 0, 0, 2, 3);
            validation.Finish();
            RibbonGroup dataTools = NewGroup(data, "Herramientas");
            AddButton(dataTools, "formatcells", "Formato", IconKind.Table, true, 0, 0, 2, 3);
            AddButton(dataTools, "clearformat", "Borrar formato", IconKind.Close, false, 2, 0, 1, 1);
            dataTools.Finish();

            List<RibbonGroup> view = new List<RibbonGroup>();
            _tabGroups[4] = view;
            RibbonGroup show = NewGroup(view, "Mostrar");
            AddButton(show, "gridlines", "Líneas de cuadrícula", IconKind.Grid, true, 0, 0, 2, 3, true);
            AddButton(show, "showformulas", "Mostrar fórmulas", IconKind.ShowFormulas, true, 2, 0, 2, 3, true);
            show.Finish();
            RibbonGroup panes = NewGroup(view, "Paneles");
            AddButton(panes, "freeze", "Inmovilizar paneles", IconKind.Grid, true, 0, 0, 2, 3);
            AddButton(panes, "unfreeze", "Movilizar paneles", IconKind.Close, false, 2, 0, 2, 3);
            panes.Finish();
            RibbonGroup zoom = NewGroup(view, "Zoom");
            AddButton(zoom, "zoomin", "Acercar", IconKind.ZoomIn, true, 0, 0, 2, 3);
            AddButton(zoom, "zoomout", "Alejar", IconKind.ZoomOut, false, 2, 0, 1, 1);
            AddButton(zoom, "zoom100", "Zoom 100%", IconKind.Zoom100, false, 2, 1, 1, 1);
            zoom.Finish();
        }

        public void SetTab(int index)
        {
            if (index < 0 || index >= _tabGroups.Length) return;
            _activeTab = index;
            _tabs.ActiveIndex = index;
            for (int t = 0; t < _tabGroups.Length; t++)
            {
                if (_tabGroups[t] == null) continue;
                bool visible = t == index;
                foreach (RibbonGroup g in _tabGroups[t]) g.Visible = visible;
            }
            LayoutGroups();
            _tabs.Invalidate();
        }

        private void LayoutGroups()
        {
            if (_tabGroups == null || _activeTab >= _tabGroups.Length || _tabGroups[_activeTab] == null) return;
            int x = 2;
            foreach (RibbonGroup g in _tabGroups[_activeTab])
            {
                g.Left = x;
                g.Top = 0;
                g.Height = _content.ClientSize.Height;
                x += g.Width;
            }
            _content.Invalidate();
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            LayoutGroups();
        }

        private void Raise(string cmd)
        {
            if (Command != null) Command(cmd);
        }

        public void SetChecked(string cmd, bool value)
        {
            List<RibbonButton> list;
            if (_buttons.TryGetValue(cmd, out list))
                foreach (RibbonButton b in list) { b.Checked = value; b.Invalidate(); }
        }

        public void SetFontSizeValue(float size)
        {
            if (_fontSize == null) return;
            string s = size.ToString("0.#", CultureInfo.GetCultureInfo("es-CO"));
            _updatingFont = true;
            if (!_fontSize.Items.Contains(s)) _fontSize.Items.Add(s);
            _fontSize.SelectedItem = s;
            _updatingFont = false;
        }

        public void SetFontNameValue(string name)
        {
            if (_fontName == null || string.IsNullOrEmpty(name)) return;
            _updatingFontName = true;
            int idx = _fontName.Items.IndexOf(name);
            if (idx < 0) { _fontName.Items.Add(name); idx = _fontName.Items.IndexOf(name); }
            _fontName.SelectedIndex = idx;
            _updatingFontName = false;
        }

        public void SetSaveEnabled(bool enabled)
        {
            _tabs.SaveEnabled = enabled;
            _tabs.Invalidate();
        }

        private void ShowFileMenu()
        {
            ContextMenuStrip m = UiFactory.CreateMenu();
            m.Items.Add(UiFactory.Item("Nuevo", Keys.Control | Keys.N, delegate { Raise("new"); }));
            m.Items.Add(UiFactory.Item("Abrir...", Keys.Control | Keys.O, delegate { Raise("open"); }));
            ToolStripMenuItem recent = new ToolStripMenuItem("Archivos recientes");
            recent.Font = Theme.Font(9f);
            if (_recentFiles.Count == 0)
            {
                ToolStripMenuItem empty = new ToolStripMenuItem("(Sin archivos recientes)");
                empty.Enabled = false;
                recent.DropDownItems.Add(empty);
            }
            else
            {
                foreach (string path in _recentFiles)
                {
                    string full = path;
                    ToolStripMenuItem item = new ToolStripMenuItem(Path.GetFileName(full));
                    item.ToolTipText = full;
                    item.Font = Theme.Font(9f);
                    item.Click += delegate { if (RecentFileSelected != null) RecentFileSelected(full); };
                    recent.DropDownItems.Add(item);
                }
            }
            m.Items.Add(recent);
            m.Items.Add(UiFactory.Separator());
            m.Items.Add(UiFactory.Item("Guardar", Keys.Control | Keys.S, delegate { Raise("save"); }));
            m.Items.Add(UiFactory.Item("Guardar como...", Keys.Control | Keys.Shift | Keys.S, delegate { Raise("saveas"); }));
            m.Items.Add(UiFactory.Item("Guardar copia OOXML strict...", delegate { Raise("savestrict"); }));
            m.Items.Add(UiFactory.Item("Exportar a CSV...", delegate { Raise("exportcsv"); }));
            m.Items.Add(UiFactory.Item("Imprimir...", Keys.Control | Keys.P, delegate { Raise("print"); }));
            m.Items.Add(UiFactory.Separator());
            m.Items.Add(UiFactory.Item("Acerca de Mini Excel...", delegate { ShowAbout(); }));
            m.Items.Add(UiFactory.Item("Salir", delegate { Raise("exit"); }));
            Rectangle r = _tabs.RectangleToScreen(_tabs.ClientRectangle);
            m.Show(new Point(r.Left + 2, r.Bottom));
        }

        private void ShowAbout()
        {
            string text =
                "Mini Excel · Beta 0.0.6\r\n" +
                "Hoja de cálculo nativa para Windows\r\n\r\n" +
                "Arranque típico: 0,3 s · 457 KB · sin dependencias externas\r\n" +
                "Formatos: .crb (nativo), .xlsx (Excel) y .csv\r\n\r\n" +
                "Atajos principales\r\n" +
                "Ctrl+S  Guardar          Ctrl+O  Abrir\r\n" +
                "Ctrl+Z  Deshacer         Ctrl+Y  Rehacer\r\n" +
                "Ctrl+B / I / U  Negrita, cursiva, subrayado\r\n" +
                "Ctrl+1  Formato de celdas\r\n" +
                "Ctrl+C / X / V  Copiar, cortar, pegar\r\n" +
                "Alt+=  Autosuma          F9  Recalcular\r\n" +
                "Ctrl+F  Buscar           Ctrl+G  Ir a una celda\r\n" +
                "Ctrl+P  Imprimir           F9  Recalcular\r\n" +
                "F2  Editar celda         Supr  Borrar contenido";
            MessageBox.Show(FindForm(), text, "Acerca de Mini Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    public sealed class ColorPaletteForm : Form
    {
        private readonly bool _allowEmpty;
        private readonly string _title;
        private Rectangle _emptyRect;
        private Rectangle _moreRect;
        private bool _inColorDialog;
        private Color _selectedColor = Color.Empty;
        public Color SelectedColor { get { return _selectedColor; } set { _selectedColor = value; } }
        public bool Accepted;

        private static readonly Color[] Palette = new Color[]
        {
            Color.Black, Color.FromArgb(64,64,64), Color.FromArgb(89,89,89), Color.FromArgb(128,128,128), Color.FromArgb(166,166,166), Color.FromArgb(200,200,200), Color.FromArgb(230,230,230), Color.White,
            Color.FromArgb(192,0,0), Color.FromArgb(255,0,0), Color.FromArgb(255,102,0), Color.FromArgb(255,153,0), Color.FromArgb(255,204,0), Color.FromArgb(255,255,0), Color.FromArgb(153,204,0), Color.FromArgb(0,153,0),
            Color.FromArgb(0,102,0), Color.FromArgb(0,153,153), Color.FromArgb(0,204,255), Color.FromArgb(0,102,204), Color.FromArgb(0,0,204), Color.FromArgb(51,51,153), Color.FromArgb(102,0,153), Color.FromArgb(153,0,204),
            Color.FromArgb(204,0,102), Color.FromArgb(255,0,102), Color.FromArgb(255,153,204), Color.FromArgb(255,204,153), Color.FromArgb(153,102,51), Color.FromArgb(102,51,0), Color.FromArgb(51,51,51), Color.FromArgb(128,128,0),
            Color.FromArgb(31,78,121), Color.FromArgb(68,114,196), Color.FromArgb(112,173,71), Color.FromArgb(255,192,0), Color.FromArgb(237,125,49), Color.FromArgb(255,0,0), Color.FromArgb(112,48,160), Color.FromArgb(0,176,240)
        };

        public ColorPaletteForm(string title, bool allowEmpty)
        {
            _title = title;
            _allowEmpty = allowEmpty;
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            BackColor = Color.White;
            ClientSize = new Size(250, 286);
            KeyPreview = true;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            if (e.KeyCode == Keys.Escape) { DialogResult = DialogResult.Cancel; Close(); }
        }

        protected override void OnDeactivate(EventArgs e)
        {
            base.OnDeactivate(e);
            if (!_inColorDialog) { DialogResult = DialogResult.Cancel; Close(); }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(Color.White);
            using (Pen p = new Pen(Color.FromArgb(180, 180, 180))) g.DrawRectangle(p, 0, 0, Width - 1, Height - 1);
            using (Font titleFont = Theme.Font(9f, true))
            {
                TextRenderer.DrawText(g, _title, titleFont, new Rectangle(10, 6, Width - 20, 18), Theme.FontDefault,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter);
            }
            int sw = 20;
            int gap = 4;
            int x0 = 9;
            int y0 = 30;
            for (int i = 0; i < Palette.Length; i++)
            {
                int col = i % 8;
                int row = i / 8;
                Rectangle r = new Rectangle(x0 + col * (sw + gap), y0 + row * (sw + gap), sw, sw);
                using (SolidBrush b = new SolidBrush(Palette[i])) g.FillRectangle(b, r);
                using (Pen p = new Pen(Color.FromArgb(190, 190, 190))) g.DrawRectangle(p, r);
            }
            int by = y0 + 8 * (sw + gap) + 8;
            _emptyRect = new Rectangle(9, by, 105, 26);
            _moreRect = new Rectangle(120, by + 2, 120, 22);
            using (SolidBrush b = new SolidBrush(Color.FromArgb(248, 248, 248))) g.FillRectangle(b, _emptyRect);
            using (Pen p = new Pen(Color.FromArgb(200, 200, 200))) g.DrawRectangle(p, _emptyRect);
            string emptyText = _allowEmpty ? "Sin relleno" : "Automático";
            TextRenderer.DrawText(g, emptyText, Theme.Font(8.5f), _emptyRect, Theme.FontDefault,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            TextRenderer.DrawText(g, "Más colores...", Theme.Font(8.5f), _moreRect, Theme.Green,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }

        protected override void OnMouseClick(MouseEventArgs e)
        {
            base.OnMouseClick(e);
            int sw = 20, gap = 4, x0 = 9, y0 = 30;
            for (int i = 0; i < Palette.Length; i++)
            {
                int col = i % 8;
                int row = i / 8;
                Rectangle r = new Rectangle(x0 + col * (sw + gap), y0 + row * (sw + gap), sw, sw);
                if (r.Contains(e.Location))
                {
                    SelectedColor = Palette[i];
                    Accepted = true;
                    DialogResult = DialogResult.OK;
                    Close();
                    return;
                }
            }
            if (_emptyRect.Contains(e.Location))
            {
                SelectedColor = _allowEmpty ? Color.Empty : Color.Black;
                Accepted = true;
                DialogResult = DialogResult.OK;
                Close();
                return;
            }
            if (_moreRect.Contains(e.Location))
            {
                _inColorDialog = true;
                using (ColorDialog cd = new ColorDialog())
                {
                    cd.FullOpen = true;
                    cd.Color = SelectedColor.IsEmpty ? Color.Black : SelectedColor;
                    if (cd.ShowDialog(this) == DialogResult.OK)
                    {
                        SelectedColor = cd.Color;
                        Accepted = true;
                        _inColorDialog = false;
                        DialogResult = DialogResult.OK;
                        Close();
                        return;
                    }
                }
                _inColorDialog = false;
                Activate();
            }
        }
    }
}
