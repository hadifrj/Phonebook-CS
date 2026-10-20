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
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new MainForm());
            }
            catch (Exception ex)
            {
                // ذخیره خطا در یک فایل متنی کنار برنامه
                try
                {
                    File.WriteAllText("error.log", DateTime.Now.ToString() + "\n\n" + ex.ToString());
                }
                catch { }

                MessageBox.Show("برنامه با خطا مواجه شد. فایل error.log در کنار برنامه بررسی کنید.\n\n" + ex.Message, 
                                "خطای مهلک", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
