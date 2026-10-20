using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace Phonebook
{
    public class RoundedButton : Button
    {
        public int CornerRadius { get; set; } = 10;
        public Color HoverColor { get; set; } = Color.Empty;

        public RoundedButton()
        {
            this.FlatStyle = FlatStyle.Flat;
            this.FlatAppearance.BorderSize = 0;
            this.BackColor = Color.FromArgb(13, 110, 253);
            this.ForeColor = Color.White;
            this.Font = new Font("Tahoma", 9, FontStyle.Bold);
            this.Cursor = Cursors.Hand;
        }

        protected override void OnPaint(PaintEventArgs pevent)
        {
            Graphics g = pevent.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            Rectangle rect = new Rectangle(0, 0, this.Width, this.Height);
            GraphicsPath path = GetRoundedPath(rect, CornerRadius);
            using (SolidBrush brush = new SolidBrush(this.BackColor))
                g.FillPath(brush, path);
            TextRenderer.DrawText(g, this.Text, this.Font, rect, this.ForeColor,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }

        protected override void OnMouseEnter(EventArgs e)
        {
            base.OnMouseEnter(e);
            if (HoverColor != Color.Empty) this.BackColor = HoverColor;
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            this.BackColor = Color.FromArgb(13, 110, 253);
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
