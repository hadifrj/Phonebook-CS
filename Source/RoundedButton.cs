using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace Phonebook
{
    public class RoundedButton : Button
    {
        public int CornerRadius { get; set; } = 8;
        public Color HoverColor { get; set; } = Color.FromArgb(15, 91, 84);
        private Color _normalColor = Color.FromArgb(20, 115, 105);
        private bool _changingHover;

        public RoundedButton()
        {
            this.DoubleBuffered = true;
            this.FlatStyle = FlatStyle.Flat;
            this.FlatAppearance.BorderSize = 0;
            this.BackColor = Color.FromArgb(20, 115, 105);
            _normalColor = this.BackColor;
            this.ForeColor = Color.White;
            this.Font = FontHelper.GetFont(10, FontStyle.Bold);
            this.Cursor = Cursors.Hand;
            this.Size = new Size(100, 35);
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            // بریدن فیزیکی گوشه‌ها برای اینکه دالبر نداشته باشد
            using (var path = GetRoundedPath(ClientRectangle, CornerRadius))
            {
                this.Region = new Region(path);
            }
        }

        protected override void OnPaint(PaintEventArgs pevent)
        {
            Graphics g = pevent.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            Rectangle rect = new Rectangle(0, 0, this.Width, this.Height);
            using (GraphicsPath path = GetRoundedPath(rect, CornerRadius))
            using (SolidBrush brush = new SolidBrush(this.BackColor))
            {
                g.FillPath(brush, path);
            }
            TextRenderer.DrawText(g, this.Text, this.Font, rect, this.ForeColor,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }

        protected override void OnMouseEnter(EventArgs e)
        {
            base.OnMouseEnter(e);
            if (HoverColor != Color.Empty)
            {
                _normalColor = this.BackColor;
                _changingHover = true;
                this.BackColor = HoverColor;
                _changingHover = false;
            }
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            this.BackColor = _normalColor;
        }

        private GraphicsPath GetRoundedPath(Rectangle rect, int radius)
        {
            GraphicsPath path = new GraphicsPath();
            float r2 = radius * 2f;
            path.AddArc(rect.X, rect.Y, r2, r2, 180, 90);
            path.AddArc(rect.Width - r2, rect.Y, r2, r2, 270, 90);
            path.AddArc(rect.Width - r2, rect.Height - r2, r2, r2, 0, 90);
            path.AddArc(rect.X, rect.Height - r2, r2, r2, 90, 90);
            path.CloseFigure();
            return path;
        }
    }
}
