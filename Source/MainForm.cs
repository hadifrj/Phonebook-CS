using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml;

namespace Phonebook
{
    public class MainForm : Form
    {
        private List<Customer> _customers = new List<Customer>();
        private Customer _selectedCustomer = null;
        
        private DataGridView gridCustomers, gridPhones, gridPersons;
        private TextBox txtSearch, txtName, txtAddress, txtPostal;
        private TabControl tabDetails;

        public MainForm()
        {
            this.RightToLeft = RightToLeft.Yes;
            InitializeUI();
            LoadExcelData();
            BindData();
        }

        private void InitializeUI()
        {
            this.Text = "دفترچه تلفن پرتابل";
            this.Size = new Size(1150, 720);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(238, 241, 247);
            this.Font = FontHelper.GetFont(10);
            
            // آپدیت گوشه‌های پنجره با تغییر سایز
            this.Resize += (s, e) => {
                if (this.WindowState == FormWindowState.Maximized) 
                    this.Region = null;
                else {
                    var path = GetRoundedPath(ClientRectangle, 16);
                    this.Region = new Region(path);
                }
            };

            // ---------------- هدر بالا ----------------
            var panelTop = new Panel { Dock = DockStyle.Top, Height = 65, BackColor = Color.White };
            panelTop.Paint += (s, e) => e.Graphics.DrawLine(new Pen(Color.FromArgb(215, 220, 230)), 0, panelTop.Height - 1, panelTop.Width, panelTop.Height - 1);
            
            var btnAdd = new RoundedButton { Text = "+ مشتری جدید", Location = new Point(20, 15), Size = new Size(130, 36) };
            btnAdd.Click += (s, e) => AddCustomer();
            
            var lblSearch = new Label { Text = "جستجو:", Location = new Point(170, 23), AutoSize = true, Font = FontHelper.GetFont(11, FontStyle.Bold) };
            txtSearch = new TextBox { Location = new Point(230, 20), Width = 250, Font = FontHelper.GetFont(11) };
            txtSearch.TextChanged += (s, e) => BindData();
            
            panelTop.Controls.AddRange(new Control[] { btnAdd, lblSearch, txtSearch });
            this.Controls.Add(panelTop);

            // ---------------- بدنه اصلی (Split) ----------------
            var split = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Vertical, FixedPanel = FixedPanel.Panel2, SplitterDistance = 650 };
            split.Panel1.Padding = new Padding(10, 10, 10, 10); // لیست (چپ)
            split.Panel2.Padding = new Padding(10, 10, 10, 10); // جزئیات (راست)
            this.Controls.Add(split);
            split.BringToFront();

            // ================== سمت راست: جزئیات (تب‌بندی) ==================
            tabDetails = new TabControl { Dock = DockStyle.Fill, Font = FontHelper.GetFont(11) };
            
            // --- تب 1: مشخصات ---
            var tabInfo = new TabPage("مشخصات");
            tabInfo.BackColor = Color.White;
            var tlpInfo = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 6, Padding = new Padding(20) };
            tlpInfo.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            tlpInfo.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));

            tlpInfo.Controls.Add(new Label { Text = "نام مشتری:", Anchor = AnchorStyles.Right, AutoSize = true, Font = FontHelper.GetFont(11, FontStyle.Bold) }, 1, 0);
            txtName = new TextBox { Dock = DockStyle.Fill, Font = FontHelper.GetFont(11), Margin = new Padding(0, 8, 0, 8) };
            tlpInfo.Controls.Add(txtName, 0, 0);

            tlpInfo.Controls.Add(new Label { Text = "آدرس:", Anchor = AnchorStyles.Right, AutoSize = true, Font = FontHelper.GetFont(11, FontStyle.Bold) }, 1, 1);
            txtAddress = new TextBox { Dock = DockStyle.Fill, Font = FontHelper.GetFont(11), Margin = new Padding(0, 8, 0, 8) };
            tlpInfo.Controls.Add(txtAddress, 0, 1);

            tlpInfo.Controls.Add(new Label { Text = "کد پستی:", Anchor = AnchorStyles.Right, AutoSize = true, Font = FontHelper.GetFont(11, FontStyle.Bold) }, 1, 2);
            txtPostal = new TextBox { Dock = DockStyle.Fill, Font = FontHelper.GetFont(11), Margin = new Padding(0, 8, 0, 8) };
            tlpInfo.Controls.Add(txtPostal, 0, 2);

            var flowBtns = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, Margin = new Padding(0, 15, 0, 0) };
            var btnSave = new RoundedButton { Text = "ذخیره تغییرات", Size = new Size(120, 36), BackColor = Color.FromArgb(25, 135, 84) };
            btnSave.Click += (s, e) => SaveCustomer();
            var btnDelete = new RoundedButton { Text = "حذف", Size = new Size(100, 36), BackColor = Color.FromArgb(220, 53, 69) };
            btnDelete.Click += (s, e) => DeleteCustomer();
            var btnPrint = new RoundedButton { Text = "چاپ آدرس پستی", Size = new Size(130, 36), BackColor = Color.FromArgb(255, 193, 7), ForeColor = Color.Black };
            btnPrint.Click += (s, e) => PrintAddress();
            
            flowBtns.Controls.Add(btnSave);
            flowBtns.Controls.Add(btnPrint);
            flowBtns.Controls.Add(btnDelete);
            tlpInfo.Controls.Add(flowBtns, 0, 3);
            tlpInfo.SetColumnSpan(flowBtns, 2);
            
            tabInfo.Controls.Add(tlpInfo);
            tabDetails.TabPages.Add(tabInfo);

            // --- تب 2: شماره‌ها ---
            var tabPhones = new TabPage("شماره‌ها");
            tabPhones.BackColor = Color.White;
            var panelPhones = new Panel { Dock = DockStyle.Fill, Padding = new Padding(15) };
            
            var flowPhoneBtns = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 50, FlowDirection = FlowDirection.RightToLeft };
            var btnAddPhone = new RoundedButton { Text = "+ افزودن شماره", Size = new Size(130, 35) };
            btnAddPhone.Click += (s, e) => AddPhone();
            var btnDelPhone = new RoundedButton { Text = "حذف", Size = new Size(80, 35), BackColor = Color.FromArgb(220, 53, 69) };
            btnDelPhone.Click += (s, e) => DeletePhone();
            flowPhoneBtns.Controls.Add(btnAddPhone);
            flowPhoneBtns.Controls.Add(btnDelPhone);

            gridPhones = CreateModernGrid();
            gridPhones.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Label", HeaderText = "نوع", FillWeight = 30 });
            gridPhones.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Number", HeaderText = "شماره", FillWeight = 70 });
            
            panelPhones.Controls.Add(gridPhones);
            panelPhones.Controls.Add(flowPhoneBtns);
            tabPhones.Controls.Add(panelPhones);
            tabDetails.TabPages.Add(tabPhones);

            // --- تب 3: افراد ---
            var tabPersons = new TabPage("افراد");
            tabPersons.BackColor = Color.White;
            var panelPersons = new Panel { Dock = DockStyle.Fill, Padding = new Padding(15) };

            var flowPersonBtns = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 50, FlowDirection = FlowDirection.RightToLeft };
            var btnAddPerson = new RoundedButton { Text = "+ افزودن فرد", Size = new Size(130, 35) };
            btnAddPerson.Click += (s, e) => AddPerson();
            var btnDelPerson = new RoundedButton { Text = "حذف", Size = new Size(80, 35), BackColor = Color.FromArgb(220, 53, 69) };
            btnDelPerson.Click += (s, e) => DeletePerson();
            flowPersonBtns.Controls.Add(btnAddPerson);
            flowPersonBtns.Controls.Add(btnDelPerson);

            gridPersons = CreateModernGrid();
            gridPersons.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Name", HeaderText = "نام فرد", FillWeight = 40 });
            gridPersons.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Role", HeaderText = "سمت", FillWeight = 30 });
            gridPersons.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Phone", HeaderText = "تلفن", FillWeight = 30 });

            panelPersons.Controls.Add(gridPersons);
            panelPersons.Controls.Add(flowPersonBtns);
            tabPersons.Controls.Add(panelPersons);
            tabDetails.TabPages.Add(tabPersons);

            split.Panel2.Controls.Add(tabDetails);

            // ================== سمت چپ: لیست مشتریان ==================
            var gridPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.White, Padding = new Padding(1) };
            gridCustomers = CreateModernGrid();
            gridCustomers.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            gridCustomers.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Name", HeaderText = "نام مشتری", FillWeight = 40 });
            gridCustomers.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Address", HeaderText = "آدرس", FillWeight = 45 });
            gridCustomers.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "PostalCode", HeaderText = "کد پستی", FillWeight = 15 });
            gridCustomers.SelectionChanged += (s, e) => SelectCustomer();
            
            gridPanel.Controls.Add(gridCustomers);
            split.Panel1.Controls.Add(gridPanel);
        }

        // متد کمکی برای ساخت جدول‌های زیبا و مدرن
        private DataGridView CreateModernGrid()
        {
            var grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false, ReadOnly = true,
                AutoGenerateColumns = false, Font = FontHelper.GetFont(10),
                BackgroundColor = Color.White, BorderStyle = BorderStyle.None,
                RowHeadersVisible = false, AllowUserToResizeRows = false,
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 40, RowTemplate = { Height = 35 }
            };
            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(31, 42, 68);
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            grid.ColumnHeadersDefaultCellStyle.Font = FontHelper.GetFont(10, FontStyle.Bold);
            grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(246, 248, 252);
            grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(13, 110, 253);
            grid.DefaultCellStyle.SelectionForeColor = Color.White;
            return grid;
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

        // ---------------- منطق دیتابیس ----------------
        private void LoadExcelData()
        {
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                using (Stream stream = assembly.GetManifestResourceStream("Phonebook.customers.xlsx"))
                {
                    if (stream == null) return;
                    using (ExcelPackage package = new ExcelPackage(stream))
                    {
                        var wsCustomers = package.Workbook.Worksheets["Customers"];
                        var wsPhones = package.Workbook.Worksheets["Phones"];
                        var wsPersons = package.Workbook.Worksheets["Persons"];
                        
                        if (wsCustomers == null || wsCustomers.Dimension == null) return;

                        for (int row = 2; row <= wsCustomers.Dimension.End.Row; row++)
                        {
                            var name = wsCustomers.Cells[row, 1].Text;
                            if (string.IsNullOrWhiteSpace(name)) continue;
                            _customers.Add(new Customer { Id = row, Name = name, Address = wsCustomers.Cells[row, 2].Text, PostalCode = wsCustomers.Cells[row, 3].Text });
                        }
                        if (wsPhones != null && wsPhones.Dimension != null)
                        {
                            for (int row = 2; row <= wsPhones.Dimension.End.Row; row++)
                            {
                                var customer = _customers.FirstOrDefault(c => c.Name == wsPhones.Cells[row, 1].Text);
                                if (customer != null) customer.Phones.Add(new Phone { Label = wsPhones.Cells[row, 2].Text, Number = wsPhones.Cells[row, 3].Text });
                            }
                        }
                        if (wsPersons != null && wsPersons.Dimension != null)
                        {
                            for (int row = 2; row <= wsPersons.Dimension.End.Row; row++)
                            {
                                var customer = _customers.FirstOrDefault(c => c.Name == wsPersons.Cells[row, 1].Text);
                                if (customer != null) customer.Persons.Add(new Person { Name = wsPersons.Cells[row, 2].Text, Role = wsPersons.Cells[row, 3].Text, Phone = wsPersons.Cells[row, 4].Text });
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("خطا در خواندن اکسل: " + ex.Message); }
        }

        private void BindData()
        {
            var search = txtSearch.Text.Trim();
            var filtered = _customers.Where(c => c.Name.Contains(search) || c.Address.Contains(search) || c.PostalCode.Contains(search) || c.Phones.Any(p => p.Number.Contains(search))).ToList();
            gridCustomers.DataSource = null;
            gridCustomers.DataSource = filtered;
            if (gridCustomers.Rows.Count > 0) gridCustomers.Rows[0].Selected = true;
        }

        private void SelectCustomer()
        {
            if (gridCustomers.SelectedRows.Count == 0) 
            {
                _selectedCustomer = null;
                txtName.Text = ""; txtAddress.Text = ""; txtPostal.Text = "";
                gridPhones.DataSource = null; gridPersons.DataSource = null;
                return;
            }
            _selectedCustomer = gridCustomers.SelectedRows[0].DataBoundItem as Customer;
            if (_selectedCustomer == null) return;

            txtName.Text = _selectedCustomer.Name;
            txtAddress.Text = _selectedCustomer.Address;
            txtPostal.Text = _selectedCustomer.PostalCode;

            gridPhones.DataSource = null; gridPhones.DataSource = _selectedCustomer.Phones;
            gridPersons.DataSource = null; gridPersons.DataSource = _selectedCustomer.Persons;
        }

        // ---------------- عملیات مشتری ----------------
        private void AddCustomer()
        {
            var newCustomer = new Customer { Id = _customers.Count + 1, Name = "مشتری جدید" };
            _customers.Add(newCustomer);
            BindData();
            txtName.Focus();
        }

        private void SaveCustomer()
        {
            if (_selectedCustomer == null) return;
            _selectedCustomer.Name = txtName.Text;
            _selectedCustomer.Address = txtAddress.Text;
            _selectedCustomer.PostalCode = txtPostal.Text;
            BindData();
            MessageBox.Show("اطلاعات در حافظه ذخیره شد.", "موفق", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void DeleteCustomer()
        {
            if (_selectedCustomer == null) return;
            if (MessageBox.Show("آیا مطمئن هستید؟", "حذف", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                _customers.Remove(_selectedCustomer);
                BindData();
            }
        }

        // ---------------- عملیات شماره‌ها ----------------
        private void AddPhone()
        {
            if (_selectedCustomer == null) return;
            using (var dlg = new FormDialog("افزودن شماره", new[] { "شماره", "نوع (همراه/ثابت)" }))
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    _selectedCustomer.Phones.Add(new Phone { Label = dlg.Result[1], Number = dlg.Result[0] });
                    SelectCustomer(); // رفرش
                }
            }
        }

        private void DeletePhone()
        {
            if (gridPhones.SelectedRows.Count == 0 || _selectedCustomer == null) return;
            var p = gridPhones.SelectedRows[0].DataBoundItem as Phone;
            if (p != null) _selectedCustomer.Phones.Remove(p);
            SelectCustomer();
        }

        // ---------------- عملیات افراد ----------------
        private void AddPerson()
        {
            if (_selectedCustomer == null) return;
            using (var dlg = new FormDialog("افزودن فرد", new[] { "نام فرد", "سمت", "تلفن" }))
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    _selectedCustomer.Persons.Add(new Person { Name = dlg.Result[0], Role = dlg.Result[1], Phone = dlg.Result[2] });
                    SelectCustomer();
                }
            }
        }

        private void DeletePerson()
        {
            if (gridPersons.SelectedRows.Count == 0 || _selectedCustomer == null) return;
            var p = gridPersons.SelectedRows[0].DataBoundItem as Person;
            if (p != null) _selectedCustomer.Persons.Remove(p);
            SelectCustomer();
        }

        // ---------------- چاپ آدرس ----------------
        private void PrintAddress()
        {
            if (_selectedCustomer == null) return;

            PrintDocument pd = new PrintDocument();
            pd.PrintPage += (s, ev) =>
            {
                Font titleFont = new Font("Tahoma", 22, FontStyle.Bold);
                Font textFont = new Font("Tahoma", 18, FontStyle.Regular);
                
                // تنظیم فاصله‌ها برای چاپ روی پاکت نامه
                float y = 150;
                ev.Graphics.DrawString(_selectedCustomer.Name, titleFont, Brushes.Black, 100, y);
                y += 50;
                ev.Graphics.DrawString(_selectedCustomer.Address, textFont, Brushes.Black, 100, y);
                y += 40;
                ev.Graphics.DrawString("کد پستی: " + _selectedCustomer.PostalCode, textFont, Brushes.Black, 100, y);
            };

            using (PrintPreviewDialog preview = new PrintPreviewDialog { Document = pd, Width = 800, Height = 600 })
            {
                preview.ShowDialog(this);
            }
        }
    }
}
