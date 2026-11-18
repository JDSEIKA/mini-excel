using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace CalcRapido
{
    public sealed class SheetTabs : Control
    {
        public Workbook Workbook;
        public int ActiveIndex;
        public event Action<int> SheetSelected;
        public event EventHandler AddRequested;
        public event Action<int> RenameRequested;
        public event Action<int> DeleteRequested;

        private readonly List<Rectangle> _rects = new List<Rectangle>();
        private Rectangle _addRect;
        private int _hover = -1;
        private bool _hoverAdd;

        public SheetTabs()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            Height = 27;
            BackColor = Color.FromArgb(243, 243, 243);
        }

        public void RefreshTabs()
        {
            _rects.Clear();
            if (Workbook == null) { Invalidate(); return; }
            int x = 8;
            using (Graphics g = CreateGraphics())
            using (Font f = Theme.Font(8.8f))
            {
                foreach (Sheet s in Workbook.Sheets)
                {
                    int w = Math.Max(88, TextRenderer.MeasureText(g, s.Name, f).Width + 26);
                    _rects.Add(new Rectangle(x, 1, w, Height - 2));
                    x += w;
                }
            }
            _addRect = new Rectangle(x + 4, 3, 21, Height - 6);
            Invalidate();
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            RefreshTabs();
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            int h = -1;
            for (int i = 0; i < _rects.Count; i++) if (_rects[i].Contains(e.Location)) { h = i; break; }
            bool ha = _addRect.Contains(e.Location);
            if (h != _hover || ha != _hoverAdd)
            {
                _hover = h; _hoverAdd = ha;
                Cursor = (h >= 0 || ha) ? Cursors.Hand : Cursors.Default;
                Invalidate();
            }
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            _hover = -1; _hoverAdd = false;
            Cursor = Cursors.Default;
            Invalidate();
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (e.Button == MouseButtons.Left)
            {
                if (_addRect.Contains(e.Location))
                {
                    if (AddRequested != null) AddRequested(this, EventArgs.Empty);
                    return;
                }
                for (int i = 0; i < _rects.Count; i++)
                {
                    if (_rects[i].Contains(e.Location))
                    {
                        if (ActiveIndex != i)
                        {
                            ActiveIndex = i;
                            Invalidate();
                            if (SheetSelected != null) SheetSelected(i);
                        }
                        return;
                    }
                }
            }
            else if (e.Button == MouseButtons.Right)
            {
                for (int i = 0; i < _rects.Count; i++)
                {
                    if (_rects[i].Contains(e.Location))
                    {
                        if (ActiveIndex != i)
                        {
                            ActiveIndex = i;
                            Invalidate();
                            if (SheetSelected != null) SheetSelected(i);
                        }
                        ContextMenuStrip m = UiFactory.CreateMenu();
                        int idx = i;
                        m.Items.Add(UiFactory.Item("Cambiar nombre...", delegate { if (RenameRequested != null) RenameRequested(idx); }));
                        m.Items.Add(UiFactory.Item("Eliminar hoja", delegate { if (DeleteRequested != null) DeleteRequested(idx); }));
                        m.Items.Add(UiFactory.Separator());
                        m.Items.Add(UiFactory.Item("Nueva hoja", delegate { if (AddRequested != null) AddRequested(this, EventArgs.Empty); }));
                        m.Show(this, e.Location);
                        return;
                    }
                }
            }
        }

        protected override void OnMouseDoubleClick(MouseEventArgs e)
        {
            base.OnMouseDoubleClick(e);
            for (int i = 0; i < _rects.Count; i++)
            {
                if (_rects[i].Contains(e.Location))
                {
                    if (RenameRequested != null) RenameRequested(i);
                    return;
                }
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            if (_rects.Count == 0) RefreshTabs();
            using (SolidBrush b = new SolidBrush(Color.FromArgb(243, 243, 243))) g.FillRectangle(b, ClientRectangle);
            using (Pen p = new Pen(Color.FromArgb(210, 210, 210))) g.DrawLine(p, 0, 0, Width, 0);
            for (int i = 0; i < _rects.Count; i++)
            {
                Rectangle r = _rects[i];
                bool active = i == ActiveIndex;
                if (active)
                {
                    using (SolidBrush b = new SolidBrush(Color.White)) g.FillRectangle(b, r);
                    using (Pen p = new Pen(Theme.Green, 2f)) g.DrawLine(p, r.Left, r.Top, r.Right, r.Top);
                }
                else if (i == _hover)
                {
                    using (SolidBrush b = new SolidBrush(Theme.Hover)) g.FillRectangle(b, r);
                }
                if (i > 0)
                {
                    using (Pen p = new Pen(Color.FromArgb(214, 214, 214))) g.DrawLine(p, r.Left, 5, r.Left, Height - 6);
                }
                Sheet sheet = Workbook != null && i < Workbook.Sheets.Count ? Workbook.Sheets[i] : null;
                string name = sheet == null ? "" : sheet.Name;
                using (Font f = Theme.Font(8.8f, active))
                {
                    TextRenderer.DrawText(g, name, f, r,
                        active ? Theme.Green : Color.FromArgb(75, 75, 75),
                        TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
                }
            }
            if (_hoverAdd)
            {
                using (SolidBrush b = new SolidBrush(Theme.GreenHover)) g.FillEllipse(b, _addRect);
            }
            using (Pen pen = new Pen(Theme.Green, 1.5f))
            {
                g.DrawEllipse(pen, _addRect);
                g.DrawLine(pen, _addRect.Left + 6, _addRect.Top + _addRect.Height / 2f, _addRect.Right - 6, _addRect.Top + _addRect.Height / 2f);
                g.DrawLine(pen, _addRect.Left + _addRect.Width / 2f, _addRect.Top + 5, _addRect.Left + _addRect.Width / 2f, _addRect.Bottom - 5);
            }
        }
    }
}
