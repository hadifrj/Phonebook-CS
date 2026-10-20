using System;
using System.Drawing;
using System.Windows.Forms;

namespace Phonebook
{
    public class FormDialog : Form
    {
        public string[] Result { get; private set; }

        public FormDialog(string title, string[] labels, string[] initialValues = null)
        {
            this.Text = title;
            this.Size = new Size(400, 100 + (labels.Length * 50));
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.White;
            this.RightToLeft = RightToLeft.Yes;
            this.Font = FontHelper.GetFont(11);

            var tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = labels.Length + 1, Padding = new Padding(20) };
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));

            var txts = new TextBox[labels.Length];

            for (int i = 0; i < labels.Length; i++)
            {
                tlp.Controls.Add(new Label { Text = labels[i], Anchor = AnchorStyles.Right, AutoSize = true, Font = FontHelper.GetFont(11, FontStyle.Bold) }, 1, i);
                txts[i] = new TextBox { Dock = DockStyle.Fill, Font = FontHelper.GetFont(11), Margin = new Padding(0, 8, 0, 8) };
                if (initialValues != null && initialValues.Length > i) txts[i].Text = initialValues[i];
                tlp.Controls.Add(txts[i], 0, i);
            }

            var btnPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 10, 0, 0) };
            var btnOk = new RoundedButton { Text = "تایید", Size = new Size(100, 35), BackColor = Color.FromArgb(25, 135, 84) };
            var btnCancel = new RoundedButton { Text = "انصراف", Size = new Size(100, 35), BackColor = Color.FromArgb(108, 117, 125) };
            
            btnOk.Click += (s, e) => {
                Result = new string[labels.Length];
                for (int i = 0; i < labels.Length; i++) Result[i] = txts[i].Text;
                this.DialogResult = DialogResult.OK;
                this.Close();
            };
            btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };
            
            btnPanel.Controls.Add(btnOk);
            btnPanel.Controls.Add(btnCancel);
            tlp.Controls.Add(btnPanel, 0, labels.Length);
            tlp.SetColumnSpan(btnPanel, 2);

            this.Controls.Add(tlp);
        }
    }
}
