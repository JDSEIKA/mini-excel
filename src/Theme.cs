using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Windows.Forms;

namespace CalcRapido
{
    public static class Theme
    {
        public static readonly Color Green = Color.FromArgb(33, 115, 70);
        public static readonly Color GreenDark = Color.FromArgb(24, 84, 51);
        public static readonly Color GreenLight = Color.FromArgb(232, 244, 236);
        public static readonly Color GreenHover = Color.FromArgb(214, 233, 221);
        public static readonly Color RibbonBack = Color.FromArgb(248, 248, 248);
        public static readonly Color RibbonBorder = Color.FromArgb(214, 214, 214);
        public static readonly Color GridLine = Color.FromArgb(217, 217, 217);
        // Encabezados oscuros, como la vista de Excel de la referencia.
        public static readonly Color GridHeaderBack = Color.FromArgb(89, 89, 89);
        public static readonly Color GridHeaderText = Color.White;
        public static readonly Color GridHeaderBorder = Color.FromArgb(64, 64, 64);
        public static readonly Color GridHeaderSelected = Color.FromArgb(64, 64, 64);
        public static readonly Color GridHeaderActive = Color.FromArgb(45, 45, 45);
        public static readonly Color CellText = Color.FromArgb(31, 31, 31);
        public static readonly Color CellError = Color.FromArgb(192, 0, 0);
        public static readonly Color SelectionFill = Color.FromArgb(86, 33, 115, 70);
        public static readonly Color StatusBack = Color.FromArgb(243, 243, 243);
        public static readonly Color Hover = Color.FromArgb(232, 242, 236);
        public static readonly Color Pressed = Color.FromArgb(210, 231, 218);
        public static readonly Color Border = Color.FromArgb(214, 214, 214);
        public static readonly Color FontDefault = Color.FromArgb(31, 31, 31);

        public static Font Font(float size, bool bold = false, bool italic = false)
        {
            return new Font("Segoe UI", size, (bold ? FontStyle.Bold : FontStyle.Regular) | (italic ? FontStyle.Italic : FontStyle.Regular), GraphicsUnit.Point);
        }
    }

    public enum IconKind
    {
        None, Save, Undo, Redo, New, Open, Print, Paste, Cut, Copy, FormatPainter,
        Bold, Italic, Underline, FontColor, FillColor, Borders,
        AlignLeft, AlignCenter, AlignRight, Currency, Percent, Comma,
        IncDecimals, DecDecimals, Date, InsertRow, InsertCol, DeleteRow, DeleteCol,
        Sum, SortAsc, SortDesc, Find, Function, Grid, ShowFormulas, NewSheet,
        ZoomIn, ZoomOut, Zoom100, Help, Close, Rename, Table, Recalculate, Chart, Filter
    }

    public static class IconPainter
    {
        public static void Draw(Graphics g, IconKind kind, Rectangle r, Color color, bool large)
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            float s = Math.Min(r.Width, r.Height);
            float x = r.X;
            float y = r.Y;
            using (Pen pen = new Pen(color, Math.Max(1.3f, s / 13f)))
            {
                pen.StartCap = LineCap.Round;
                pen.EndCap = LineCap.Round;
                pen.LineJoin = LineJoin.Round;
                switch (kind)
                {
                    case IconKind.Save:
                        {
                            RectangleF body = new RectangleF(x + s * 0.16f, y + s * 0.12f, s * 0.68f, s * 0.76f);
                            g.DrawRectangle(pen, body.X, body.Y, body.Width, body.Height);
                            g.DrawLine(pen, x + s * 0.3f, y + s * 0.12f, x + s * 0.3f, y + s * 0.4f);
                            g.DrawLine(pen, x + s * 0.7f, y + s * 0.12f, x + s * 0.7f, y + s * 0.4f);
                            g.DrawRectangle(pen, x + s * 0.32f, y + s * 0.58f, s * 0.36f, s * 0.3f);
                            break;
                        }
                    case IconKind.Undo:
                    case IconKind.Redo:
                        {
                            bool redo = kind == IconKind.Redo;
                            Color accent = redo ? Color.FromArgb(22, 163, 74) :
                                                  Color.FromArgb(37, 99, 235);
                            RectangleF oval = new RectangleF(x + s * 0.06f, y + s * 0.06f,
                                s * 0.88f, s * 0.88f);
                            using (SolidBrush badge = new SolidBrush(Color.FromArgb(30, accent)))
                                g.FillEllipse(badge, oval);
                            using (Pen arrow = new Pen(accent, Math.Max(2.2f, s / 7f)))
                            {
                                arrow.StartCap = LineCap.Round;
                                arrow.EndCap = LineCap.Round;
                                g.DrawArc(arrow, x + s * 0.16f, y + s * 0.16f,
                                    s * 0.68f, s * 0.68f, redo ? 35 : 145, 250);
                                float angle = (redo ? 285 : 35) * (float)Math.PI / 180f;
                                float ex = x + s * 0.5f + (float)Math.Cos(angle) * s * 0.34f;
                                float ey = y + s * 0.5f + (float)Math.Sin(angle) * s * 0.34f;
                                float dir = redo ? 1f : -1f;
                                using (SolidBrush head = new SolidBrush(accent))
                                    g.FillPolygon(head, new PointF[]
                                    {
                                        new PointF(ex + dir * s * 0.13f, ey),
                                        new PointF(ex - dir * s * 0.07f, ey - s * 0.11f),
                                        new PointF(ex - dir * s * 0.07f, ey + s * 0.11f)
                                    });
                            }
                            using (Font glyph = new Font("Segoe UI Symbol", s * 0.28f,
                                FontStyle.Bold, GraphicsUnit.Pixel))
                            using (StringFormat sf = new StringFormat())
                            using (SolidBrush text = new SolidBrush(accent))
                            {
                                sf.Alignment = StringAlignment.Center;
                                sf.LineAlignment = StringAlignment.Center;
                                g.DrawString(redo ? "↷" : "↶", glyph, text,
                                    new RectangleF(x, y + s * 0.02f, s, s), sf);
                            }
                            break;
                        }
                    case IconKind.New:
                        {
                            g.DrawRectangle(pen, x + s * 0.22f, y + s * 0.12f, s * 0.56f, s * 0.76f);
                            g.DrawLine(pen, x + s * 0.62f, y + s * 0.12f, x + s * 0.78f, y + s * 0.28f);
                            g.DrawLine(pen, x + s * 0.62f, y + s * 0.12f, x + s * 0.62f, y + s * 0.28f);
                            g.DrawLine(pen, x + s * 0.62f, y + s * 0.28f, x + s * 0.78f, y + s * 0.28f);
                            break;
                        }
                    case IconKind.Open:
                        {
                            g.DrawRectangle(pen, x + s * 0.14f, y + s * 0.32f, s * 0.5f, s * 0.45f);
                            using (GraphicsPath p = new GraphicsPath())
                            {
                                p.AddLine(x + s * 0.14f, y + s * 0.32f, x + s * 0.28f, y + s * 0.18f);
                                p.AddLine(x + s * 0.28f, y + s * 0.18f, x + s * 0.55f, y + s * 0.18f);
                                p.AddLine(x + s * 0.55f, y + s * 0.18f, x + s * 0.6f, y + s * 0.32f);
                                g.DrawPath(pen, p);
                            }
                            g.DrawLine(pen, x + s * 0.58f, y + s * 0.45f, x + s * 0.9f, y + s * 0.45f);
                            g.DrawLine(pen, x + s * 0.9f, y + s * 0.45f, x + s * 0.78f, y + s * 0.77f);
                            break;
                        }
                    case IconKind.Print:
                        {
                            g.DrawRectangle(pen, x + s * 0.2f, y + s * 0.42f, s * 0.6f, s * 0.35f);
                            g.DrawRectangle(pen, x + s * 0.3f, y + s * 0.15f, s * 0.4f, s * 0.27f);
                            g.DrawRectangle(pen, x + s * 0.3f, y + s * 0.62f, s * 0.4f, s * 0.24f);
                            break;
                        }
                    case IconKind.Paste:
                        {
                            g.DrawRectangle(pen, x + s * 0.22f, y + s * 0.2f, s * 0.56f, s * 0.68f);
                            g.DrawRectangle(pen, x + s * 0.36f, y + s * 0.12f, s * 0.28f, s * 0.16f);
                            g.DrawLine(pen, x + s * 0.34f, y + s * 0.45f, x + s * 0.66f, y + s * 0.45f);
                            g.DrawLine(pen, x + s * 0.34f, y + s * 0.6f, x + s * 0.66f, y + s * 0.6f);
                            break;
                        }
                    case IconKind.Cut:
                        {
                            g.DrawEllipse(pen, x + s * 0.15f, y + s * 0.62f, s * 0.2f, s * 0.2f);
                            g.DrawEllipse(pen, x + s * 0.65f, y + s * 0.62f, s * 0.2f, s * 0.2f);
                            g.DrawLine(pen, x + s * 0.32f, y + s * 0.62f, x + s * 0.62f, y + s * 0.16f);
                            g.DrawLine(pen, x + s * 0.68f, y + s * 0.62f, x + s * 0.38f, y + s * 0.16f);
                            break;
                        }
                    case IconKind.Copy:
                        {
                            g.DrawRectangle(pen, x + s * 0.16f, y + s * 0.28f, s * 0.48f, s * 0.56f);
                            g.DrawRectangle(pen, x + s * 0.36f, y + s * 0.14f, s * 0.48f, s * 0.56f);
                            break;
                        }
                    case IconKind.FormatPainter:
                        {
                            g.DrawRectangle(pen, x + s * 0.22f, y + s * 0.6f, s * 0.56f, s * 0.22f);
                            g.DrawLine(pen, x + s * 0.5f, y + s * 0.6f, x + s * 0.5f, y + s * 0.35f);
                            g.DrawEllipse(pen, x + s * 0.4f, y + s * 0.1f, s * 0.2f, s * 0.25f);
                            break;
                        }
                    case IconKind.Bold:
                    case IconKind.Italic:
                    case IconKind.Underline:
                        {
                            string t = kind == IconKind.Bold ? "B" : kind == IconKind.Italic ? "I" : "U";
                            FontStyle st = kind == IconKind.Bold ? FontStyle.Bold : kind == IconKind.Italic ? FontStyle.Italic : FontStyle.Underline;
                            using (Font f = new Font("Segoe UI", s * 0.62f, st, GraphicsUnit.Pixel))
                            using (StringFormat sf = new StringFormat())
                            {
                                sf.Alignment = StringAlignment.Center;
                                sf.LineAlignment = StringAlignment.Center;
                                using (SolidBrush br = new SolidBrush(color))
                                    g.DrawString(t, f, br, new RectangleF(x, y - 1, s, s), sf);
                            }
                            break;
                        }
                    case IconKind.FontColor:
                        {
                            using (Font f = new Font("Segoe UI", s * 0.55f, FontStyle.Bold, GraphicsUnit.Pixel))
                            using (SolidBrush br = new SolidBrush(color))
                                g.DrawString("A", f, br, new PointF(x + s * 0.18f, y + s * 0.02f));
                            using (SolidBrush br = new SolidBrush(Color.FromArgb(192, 0, 0)))
                                g.FillRectangle(br, x + s * 0.18f, y + s * 0.78f, s * 0.64f, s * 0.14f);
                            break;
                        }
                    case IconKind.FillColor:
                        {
                            using (GraphicsPath p = new GraphicsPath())
                            {
                                p.AddPolygon(new PointF[]
                                {
                                    new PointF(x + s*0.16f, y + s*0.5f),
                                    new PointF(x + s*0.5f, y + s*0.14f),
                                    new PointF(x + s*0.82f, y + s*0.46f),
                                    new PointF(x + s*0.48f, y + s*0.8f)
                                });
                                g.DrawPath(pen, p);
                            }
                            g.DrawLine(pen, x + s * 0.28f, y + s * 0.72f, x + s * 0.72f, y + s * 0.72f);
                            using (SolidBrush br = new SolidBrush(Color.FromArgb(255, 196, 0)))
                                g.FillRectangle(br, x + s * 0.2f, y + s * 0.82f, s * 0.6f, s * 0.12f);
                            break;
                        }
                    case IconKind.Borders:
                        {
                            g.DrawRectangle(pen, x + s * 0.18f, y + s * 0.18f, s * 0.64f, s * 0.64f);
                            g.DrawLine(pen, x + s * 0.18f, y + s * 0.5f, x + s * 0.82f, y + s * 0.5f);
                            g.DrawLine(pen, x + s * 0.5f, y + s * 0.18f, x + s * 0.5f, y + s * 0.82f);
                            break;
                        }
                    case IconKind.AlignLeft:
                    case IconKind.AlignCenter:
                    case IconKind.AlignRight:
                        {
                            float left = x + s * 0.15f;
                            float right = x + s * 0.85f;
                            for (int i = 0; i < 4; i++)
                            {
                                float yy = y + s * (0.22f + i * 0.18f);
                                float x1 = left, x2 = right;
                                if (kind == IconKind.AlignCenter) { x1 = x + s * 0.28f; x2 = x + s * 0.72f; }
                                if (kind == IconKind.AlignRight) { x1 = x + s * 0.4f; }
                                if (i == 1) { if (kind == IconKind.AlignLeft) x2 = x + s * 0.65f; else if (kind == IconKind.AlignCenter) { x1 = x + s * 0.2f; x2 = x + s * 0.8f; } else x1 = x + s * 0.2f; }
                                if (i == 3) { if (kind == IconKind.AlignLeft) x2 = x + s * 0.55f; else if (kind == IconKind.AlignCenter) { x1 = x + s * 0.36f; x2 = x + s * 0.64f; } else x1 = x + s * 0.3f; }
                                g.DrawLine(pen, x1, yy, x2, yy);
                            }
                            break;
                        }
                    case IconKind.Currency:
                    case IconKind.Percent:
                    case IconKind.Comma:
                        {
                            string t = kind == IconKind.Currency ? "$" : kind == IconKind.Percent ? "%" : ",0";
                            using (Font f = new Font("Segoe UI", s * 0.55f, FontStyle.Bold, GraphicsUnit.Pixel))
                            using (StringFormat sf = new StringFormat())
                            {
                                sf.Alignment = StringAlignment.Center;
                                sf.LineAlignment = StringAlignment.Center;
                                using (SolidBrush br = new SolidBrush(color))
                                    g.DrawString(t, f, br, new RectangleF(x, y, s, s), sf);
                            }
                            break;
                        }
                    case IconKind.IncDecimals:
                    case IconKind.DecDecimals:
                        {
                            using (Font f = new Font("Segoe UI", s * 0.44f, FontStyle.Bold, GraphicsUnit.Pixel))
                            using (SolidBrush br = new SolidBrush(color))
                                g.DrawString(".00", f, br, new PointF(x + s * 0.02f, y + s * 0.15f));
                            bool plus = kind == IconKind.IncDecimals;
                            using (Pen p = new Pen(color, Math.Max(1.4f, s / 14f)))
                            {
                                if (plus)
                                {
                                    g.DrawLine(p, x + s * 0.65f, y + s * 0.72f, x + s * 0.88f, y + s * 0.72f);
                                    g.DrawLine(p, x + s * 0.765f, y + s * 0.61f, x + s * 0.765f, y + s * 0.83f);
                                }
                                else g.DrawLine(p, x + s * 0.65f, y + s * 0.72f, x + s * 0.88f, y + s * 0.72f);
                            }
                            break;
                        }
                    case IconKind.Date:
                        {
                            g.DrawRectangle(pen, x + s * 0.16f, y + s * 0.2f, s * 0.68f, s * 0.66f);
                            g.DrawLine(pen, x + s * 0.16f, y + s * 0.38f, x + s * 0.84f, y + s * 0.38f);
                            g.DrawLine(pen, x + s * 0.32f, y + s * 0.12f, x + s * 0.32f, y + s * 0.28f);
                            g.DrawLine(pen, x + s * 0.68f, y + s * 0.12f, x + s * 0.68f, y + s * 0.28f);
                            break;
                        }
                    case IconKind.InsertRow:
                    case IconKind.InsertCol:
                    case IconKind.DeleteRow:
                    case IconKind.DeleteCol:
                        {
                            bool delete = kind == IconKind.DeleteRow || kind == IconKind.DeleteCol;
                            bool row = kind == IconKind.InsertRow || kind == IconKind.DeleteRow;
                            using (Pen p = new Pen(color, Math.Max(1.2f, s / 15f)))
                            {
                                g.DrawRectangle(p, x + s * 0.14f, y + s * 0.16f, s * 0.72f, s * 0.68f);
                                if (row) g.DrawLine(p, x + s * 0.14f, y + s * 0.42f, x + s * 0.86f, y + s * 0.42f);
                                else g.DrawLine(p, x + s * 0.42f, y + s * 0.16f, x + s * 0.42f, y + s * 0.84f);
                                Color accent = delete ? Color.FromArgb(192, 0, 0) : Theme.Green;
                                using (Pen ap = new Pen(accent, Math.Max(1.8f, s / 10f)))
                                {
                                    float cx = x + s * 0.63f, cy = y + s * 0.63f;
                                    g.DrawLine(ap, cx - s * 0.13f, cy, cx + s * 0.13f, cy);
                                    if (!delete) g.DrawLine(ap, cx, cy - s * 0.13f, cx, cy + s * 0.13f);
                                }
                            }
                            break;
                        }
                    case IconKind.Sum:
                        {
                            using (Font f = new Font("Segoe UI", s * 0.72f, FontStyle.Bold, GraphicsUnit.Pixel))
                            using (StringFormat sf = new StringFormat())
                            {
                                sf.Alignment = StringAlignment.Center;
                                sf.LineAlignment = StringAlignment.Center;
                                using (SolidBrush br = new SolidBrush(color))
                                    g.DrawString("\u03A3", f, br, new RectangleF(x, y - s * 0.04f, s, s), sf);
                            }
                            break;
                        }
                    case IconKind.SortAsc:
                    case IconKind.SortDesc:
                        {
                            bool asc = kind == IconKind.SortAsc;
                            using (Font f = new Font("Segoe UI", s * 0.42f, FontStyle.Bold, GraphicsUnit.Pixel))
                            using (SolidBrush br = new SolidBrush(color))
                            {
                                g.DrawString(asc ? "A" : "Z", f, br, new PointF(x + s * 0.12f, y + s * 0.08f));
                                g.DrawString(asc ? "Z" : "A", f, br, new PointF(x + s * 0.12f, y + s * 0.5f));
                            }
                            using (Pen p = new Pen(color, Math.Max(1.4f, s / 14f)))
                            {
                                float ax = x + s * 0.72f;
                                float ay1 = asc ? y + s * 0.22f : y + s * 0.78f;
                                float ay2 = asc ? y + s * 0.78f : y + s * 0.22f;
                                g.DrawLine(p, ax, ay1, ax, ay2);
                                g.DrawLine(p, ax - s * 0.1f, ay2 + (asc ? -s * 0.12f : s * 0.12f), ax, ay2);
                                g.DrawLine(p, ax + s * 0.1f, ay2 + (asc ? -s * 0.12f : s * 0.12f), ax, ay2);
                            }
                            break;
                        }
                    case IconKind.Filter:
                        {
                            PointF[] funnel = new PointF[]
                            {
                                new PointF(x + s * 0.16f, y + s * 0.18f),
                                new PointF(x + s * 0.84f, y + s * 0.18f),
                                new PointF(x + s * 0.60f, y + s * 0.50f),
                                new PointF(x + s * 0.60f, y + s * 0.82f),
                                new PointF(x + s * 0.40f, y + s * 0.70f),
                                new PointF(x + s * 0.40f, y + s * 0.50f)
                            };
                            g.DrawPolygon(pen, funnel);
                            break;
                        }
                    case IconKind.Find:
                        {
                            g.DrawEllipse(pen, x + s * 0.14f, y + s * 0.14f, s * 0.5f, s * 0.5f);
                            g.DrawLine(pen, x + s * 0.58f, y + s * 0.58f, x + s * 0.86f, y + s * 0.86f);
                            break;
                        }
                    case IconKind.Function:
                        {
                            using (Font f = new Font("Georgia", s * 0.58f, FontStyle.Italic, GraphicsUnit.Pixel))
                            using (StringFormat sf = new StringFormat())
                            {
                                sf.Alignment = StringAlignment.Center;
                                sf.LineAlignment = StringAlignment.Center;
                                using (SolidBrush br = new SolidBrush(color))
                                    g.DrawString("fx", f, br, new RectangleF(x, y, s, s), sf);
                            }
                            break;
                        }
                    case IconKind.Grid:
                        {
                            g.DrawRectangle(pen, x + s * 0.14f, y + s * 0.14f, s * 0.72f, s * 0.72f);
                            g.DrawLine(pen, x + s * 0.38f, y + s * 0.14f, x + s * 0.38f, y + s * 0.86f);
                            g.DrawLine(pen, x + s * 0.62f, y + s * 0.14f, x + s * 0.62f, y + s * 0.86f);
                            g.DrawLine(pen, x + s * 0.14f, y + s * 0.38f, x + s * 0.86f, y + s * 0.38f);
                            g.DrawLine(pen, x + s * 0.14f, y + s * 0.62f, x + s * 0.86f, y + s * 0.62f);
                            break;
                        }
                    case IconKind.ShowFormulas:
                        {
                            using (Font f = new Font("Consolas", s * 0.56f, FontStyle.Bold, GraphicsUnit.Pixel))
                            using (StringFormat sf = new StringFormat())
                            {
                                sf.Alignment = StringAlignment.Center;
                                sf.LineAlignment = StringAlignment.Center;
                                using (SolidBrush br = new SolidBrush(color))
                                    g.DrawString("=", f, br, new RectangleF(x, y, s, s), sf);
                            }
                            break;
                        }
                    case IconKind.Recalculate:
                        {
                            g.DrawArc(pen, x + s * 0.18f, y + s * 0.18f, s * 0.64f, s * 0.64f, 35, 285);
                            PointF tip = new PointF(x + s * 0.82f, y + s * 0.3f);
                            g.DrawLine(pen, tip, new PointF(tip.X - s * 0.16f, tip.Y - s * 0.02f));
                            g.DrawLine(pen, tip, new PointF(tip.X + s * 0.02f, tip.Y - s * 0.18f));
                            break;
                        }
                    case IconKind.NewSheet:
                        {
                            g.DrawRectangle(pen, x + s * 0.16f, y + s * 0.14f, s * 0.54f, s * 0.72f);
                            using (Pen p = new Pen(Theme.Green, Math.Max(1.6f, s / 11f)))
                            {
                                g.DrawLine(p, x + s * 0.68f, y + s * 0.68f, x + s * 0.9f, y + s * 0.68f);
                                g.DrawLine(p, x + s * 0.79f, y + s * 0.57f, x + s * 0.79f, y + s * 0.79f);
                            }
                            break;
                        }
                    case IconKind.ZoomIn:
                    case IconKind.ZoomOut:
                        {
                            g.DrawEllipse(pen, x + s * 0.14f, y + s * 0.14f, s * 0.52f, s * 0.52f);
                            g.DrawLine(pen, x + s * 0.58f, y + s * 0.58f, x + s * 0.86f, y + s * 0.86f);
                            using (Pen p = new Pen(color, Math.Max(1.4f, s / 14f)))
                            {
                                g.DrawLine(p, x + s * 0.28f, y + s * 0.4f, x + s * 0.52f, y + s * 0.4f);
                                if (kind == IconKind.ZoomIn) g.DrawLine(p, x + s * 0.4f, y + s * 0.28f, x + s * 0.4f, y + s * 0.52f);
                            }
                            break;
                        }
                    case IconKind.Zoom100:
                        {
                            using (Font f = new Font("Segoe UI", s * 0.32f, FontStyle.Bold, GraphicsUnit.Pixel))
                            using (StringFormat sf = new StringFormat())
                            {
                                sf.Alignment = StringAlignment.Center;
                                sf.LineAlignment = StringAlignment.Center;
                                using (SolidBrush br = new SolidBrush(color))
                                    g.DrawString("100%", f, br, new RectangleF(x, y, s, s), sf);
                            }
                            break;
                        }
                    case IconKind.Help:
                        {
                            using (Font f = new Font("Segoe UI", s * 0.55f, FontStyle.Bold, GraphicsUnit.Pixel))
                            using (StringFormat sf = new StringFormat())
                            {
                                sf.Alignment = StringAlignment.Center;
                                sf.LineAlignment = StringAlignment.Center;
                                using (SolidBrush br = new SolidBrush(color))
                                    g.DrawString("?", f, br, new RectangleF(x, y, s, s), sf);
                            }
                            break;
                        }
                    case IconKind.Close:
                        {
                            g.DrawLine(pen, x + s * 0.24f, y + s * 0.24f, x + s * 0.76f, y + s * 0.76f);
                            g.DrawLine(pen, x + s * 0.76f, y + s * 0.24f, x + s * 0.24f, y + s * 0.76f);
                            break;
                        }
                    case IconKind.Rename:
                        {
                            g.DrawRectangle(pen, x + s * 0.14f, y + s * 0.25f, s * 0.55f, s * 0.55f);
                            g.DrawLine(pen, x + s * 0.6f, y + s * 0.4f, x + s * 0.86f, y + s * 0.14f);
                            break;
                        }
                    case IconKind.Chart:
                        {
                            g.DrawLine(pen, x + s * 0.16f, y + s * 0.82f, x + s * 0.86f, y + s * 0.82f);
                            g.DrawLine(pen, x + s * 0.16f, y + s * 0.82f, x + s * 0.16f, y + s * 0.16f);
                            g.DrawRectangle(pen, x + s * 0.27f, y + s * 0.52f, s * 0.14f, s * 0.3f);
                            g.DrawRectangle(pen, x + s * 0.47f, y + s * 0.38f, s * 0.14f, s * 0.44f);
                            g.DrawRectangle(pen, x + s * 0.67f, y + s * 0.22f, s * 0.14f, s * 0.6f);
                            break;
                        }
                    case IconKind.Table:
                        {
                            g.DrawRectangle(pen, x + s * 0.14f, y + s * 0.2f, s * 0.72f, s * 0.6f);
                            g.DrawLine(pen, x + s * 0.14f, y + s * 0.38f, x + s * 0.86f, y + s * 0.38f);
                            g.DrawLine(pen, x + s * 0.4f, y + s * 0.38f, x + s * 0.4f, y + s * 0.8f);
                            break;
                        }
                }
            }
        }
    }

    public sealed class OfficeColorTable : ProfessionalColorTable
    {
        public override Color MenuItemSelected { get { return Theme.GreenHover; } }
        public override Color MenuItemBorder { get { return Theme.GreenHover; } }
        public override Color MenuBorder { get { return Theme.Border; } }
        public override Color ToolStripDropDownBackground { get { return Color.White; } }
        public override Color ImageMarginGradientBegin { get { return Color.White; } }
        public override Color ImageMarginGradientMiddle { get { return Color.White; } }
        public override Color ImageMarginGradientEnd { get { return Color.White; } }
        public override Color MenuItemSelectedGradientBegin { get { return Theme.GreenHover; } }
        public override Color MenuItemSelectedGradientEnd { get { return Theme.GreenHover; } }
        public override Color MenuItemPressedGradientBegin { get { return Theme.GreenHover; } }
        public override Color MenuItemPressedGradientEnd { get { return Theme.GreenHover; } }
        public override Color SeparatorDark { get { return Theme.Border; } }
        public override Color SeparatorLight { get { return Color.White; } }
    }
}
