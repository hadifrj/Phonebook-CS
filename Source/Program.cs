using System;
using System.IO;
using System.Windows.Forms;

namespace Phonebook
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            try
            {
                FontHelper.LoadFonts();
                
                // اگر فونت لود نشده باشد، یک هشدار می‌دهد تا مطمئن شویم فایل‌ها آپلود شده‌اند
                if (!FontHelper.FontLoaded)
                {
                    MessageBox.Show("هشدار: فونت وزیرمتن پیدا نشد! برنامه با فونت پیش‌فرض (Tahoma) اجرا می‌شود. لطفاً مطمئن شوید فایل‌های فونت در پوشه Assets آپلود شده‌اند.", "هشدار فونت", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new MainForm());
            }
            catch (Exception ex)
            {
                try
                {
                    string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                    string logPath = Path.Combine(appData, "Phonebook", "error.log");
                    Directory.CreateDirectory(Path.GetDirectoryName(logPath));
                    File.WriteAllText(logPath, DateTime.Now.ToString() + "\n\n" + ex.ToString());
                }
                catch { }

                MessageBox.Show("برنامه با خطا مواجه شد. فایل error.log در AppData بررسی کنید.\n\n" + ex.Message, 
                                "خطای مهلک", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
