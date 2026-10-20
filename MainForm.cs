using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
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

        public MainForm()
        {
            InitializeUI();
            LoadExcelData();
            BindData();
        }

        private void InitializeUI()
        {
            this.Text = "دفترچه تلفن";
            this.Size = new Size(1000, 650);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.White;
            this.Font = new Font("Tahoma", 9);
            this.RightToLeft = RightToLeft.Yes;
            this.RightToLeftLayout = true;

            // گرد کردن پنجره اصلی
            this.Resize += (s, e) => {
                GraphicsPath path = new GraphicsPath();
                int r = 20;
                path.AddArc(0, 0, r, r, 180, 90);
                path.AddArc(this.Width - r, 0, r, r, 270, 90);
                path.AddArc(this.Width - r, this.Height - r, r, r, 0, 90);
                path.AddArc(0, this.Height - r, r, r, 90, 90);
                this.Region = new Region(path);
            };

            // دکمه مشتری جدید
            var btnAdd = new RoundedButton { Text = "+ مشتری جدید", Location = new Point(20, 20), Size = new Size(120, 35) };
            btnAdd.Click += (s, e) => AddCustomer();
            this.Controls.Add(btnAdd);

            // جستجو
            var lblSearch = new Label { Text = "جستجو:", Location = new Point(160, 28), AutoSize = true };
            txtSearch = new TextBox { Location = new Point(210, 25), Width = 200 };
            txtSearch.TextChanged += (s, e) => BindData();
            this.Controls.Add(lblSearch);
            this.Controls.Add(txtSearch);

            // جدول مشتریان (ستون‌ها دستی تعریف می‌شوند)
            gridCustomers = new DataGridView 
            { 
                Location = new Point(20, 70), Size = new Size(400, 520), 
                AllowUserToAddRows = false, ReadOnly = true, 
                SelectionMode = DataGridViewSelectionMode.FullRowSelect, 
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, 
                BackgroundColor = Color.White, BorderStyle = BorderStyle.FixedSingle,
                AutoGenerateColumns = false
            };
            gridCustomers.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Name", HeaderText = "نام مشتری", FillWeight = 50 });
            gridCustomers.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Address", HeaderText = "آدرس", FillWeight = 40 });
            gridCustomers.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "PostalCode", HeaderText = "کد پستی", FillWeight = 15 });
            gridCustomers.SelectionChanged += (s, e) => SelectCustomer();
            this.Controls.Add(gridCustomers);

            // پنل جزئیات
            var panelDetails = new Panel { Location = new Point(440, 70), Size = new Size(530, 520), BorderStyle = BorderStyle.FixedSingle };
            this.Controls.Add(panelDetails);

            var lblName = new Label { Text = "نام مشتری:", Location = new Point(400, 20), AutoSize = true };
            txtName = new TextBox { Location = new Point(200, 20), Width = 300 };
            var lblAddress = new Label { Text = "آدرس:", Location = new Point(400, 60), AutoSize = true };
            txtAddress = new TextBox { Location = new Point(200, 60), Width = 300 };
            var lblPostal = new Label { Text = "کد پستی:", Location = new Point(400, 100), AutoSize = true };
            txtPostal = new TextBox { Location = new Point(200, 100), Width = 300 };
            
            var btnSave = new RoundedButton { Text = "ذخیره", Location = new Point(200, 150), Size = new Size(100, 35) };
            btnSave.Click += (s, e) => SaveCustomer();
            var btnDelete = new RoundedButton { Text = "حذف", Location = new Point(310, 150), Size = new Size(100, 35), BackColor = Color.FromArgb(220, 53, 69) };
            btnDelete.Click += (s, e) => DeleteCustomer();

            panelDetails.Controls.AddRange(new Control[] { lblName, txtName, lblAddress, txtAddress, lblPostal, txtPostal, btnSave, btnDelete });

            // جدول شماره‌ها
            gridPhones = new DataGridView 
            { 
                Location = new Point(20, 230), Size = new Size(480, 130), 
                AllowUserToAddRows = false, ReadOnly = true, 
                BackgroundColor = Color.White, BorderStyle = BorderStyle.FixedSingle,
                AutoGenerateColumns = false 
            };
            gridPhones.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Label", HeaderText = "نوع", Width = 150 });
            gridPhones.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Number", HeaderText = "شماره", Width = 250 });
            panelDetails.Controls.Add(gridPhones);

            // جدول افراد
            gridPersons = new DataGridView 
            { 
                Location = new Point(20, 370), Size = new Size(480, 130), 
                AllowUserToAddRows = false, ReadOnly = true, 
                BackgroundColor = Color.White, BorderStyle = BorderStyle.FixedSingle,
                AutoGenerateColumns = false 
            };
            gridPersons.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Name", HeaderText = "نام فرد", Width = 150 });
            gridPersons.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Role", HeaderText = "سمت", Width = 150 });
            gridPersons.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = "Phone", HeaderText = "تلفن", Width = 150 });
            panelDetails.Controls.Add(gridPersons);
        }

        private void LoadExcelData()
        {
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                using (Stream stream = assembly.GetManifestResourceStream("Phonebook.customers.xlsx"))
                {
                    if (stream == null) { MessageBox.Show("فایل اکسل پیدا نشد!"); return; }
                    using (ExcelPackage package = new ExcelPackage(stream))
                    {
                        var wsCustomers = package.Workbook.Worksheets["Customers"];
                        var wsPhones = package.Workbook.Worksheets["Phones"];
                        var wsPersons = package.Workbook.Worksheets["Persons"];
                        
                        if (wsCustomers == null || wsCustomers.Dimension == null) 
                        { 
                            MessageBox.Show("شیت Customers در اکسل خالی است یا پیدا نشد!"); 
                            return; 
                        }

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
        }

        private void SelectCustomer()
        {
            if (gridCustomers.SelectedRows.Count == 0) return;
            _selectedCustomer = gridCustomers.SelectedRows[0].DataBoundItem as Customer;
            if (_selectedCustomer == null) return;

            txtName.Text = _selectedCustomer.Name;
            txtAddress.Text = _selectedCustomer.Address;
            txtPostal.Text = _selectedCustomer.PostalCode;

            gridPhones.DataSource = null;
            gridPhones.DataSource = _selectedCustomer.Phones;

            gridPersons.DataSource = null;
            gridPersons.DataSource = _selectedCustomer.Persons;
        }

        private void AddCustomer()
        {
            var newCustomer = new Customer { Id = _customers.Count + 1, Name = "مشتری جدید" };
            _customers.Add(newCustomer);
            BindData();
        }

        private void SaveCustomer()
        {
            if (_selectedCustomer == null) return;
            _selectedCustomer.Name = txtName.Text;
            _selectedCustomer.Address = txtAddress.Text;
            _selectedCustomer.PostalCode = txtPostal.Text;
            BindData();
            MessageBox.Show("اطلاعات ذخیره شد.");
        }

        private void DeleteCustomer()
        {
            if (_selectedCustomer == null) return;
            if (MessageBox.Show("آیا مطمئن هستید؟", "حذف", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                _customers.Remove(_selectedCustomer);
                BindData();
            }
        }
    }
}
