using System;
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
                // اگر هر خطایی رخ داد، این پنجره باز می‌شود و متن خطا را نشان می‌دهد
                MessageBox.Show("خطای مهلک در اجرای برنامه:\n\n" + ex.ToString(), 
                                "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
