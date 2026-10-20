
using System;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace Phonebook
{
    public static class FontHelper
    {
        private static PrivateFontCollection _fonts = new PrivateFontCollection();
        public static string FamilyName = "Tahoma"; // فونت پیش‌فرض در صورت خطا

        public static void LoadFonts()
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                // نام منابع دقیقاً طبق LogicalName در فایل csproj تنظیم شده است
                LoadFont(assembly, "Phonebook.Vazirmatn-Regular.ttf");
                LoadFont(assembly, "Phonebook.Vazirmatn-Bold.ttf");
                
                if (_fonts.Families.Length > 0)
                {
                    FamilyName = _fonts.Families[0].Name;
                }
            }
            catch 
            {
                // اگر فونت لود نشد، برنامه کرش نمی‌کند و از Tahoma استفاده می‌کند
            }
        }

        private static void LoadFont(Assembly assembly, string resourceName)
        {
            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream != null)
                {
                    byte[] fontData = new byte[stream.Length];
                    stream.Read(fontData, 0, fontData.Length);
                    IntPtr fontPtr = Marshal.AllocCoTaskMem(fontData.Length);
                    Marshal.Copy(fontData, 0, fontPtr, fontData.Length);
                    _fonts.AddMemoryFont(fontPtr, fontData.Length);
                    Marshal.FreeCoTaskMem(fontPtr);
                }
            }
        }

        public static Font GetFont(float size, FontStyle style = FontStyle.Regular)
        {
            try 
            { 
                // از فونت لود شده استفاده می‌کند
                return new Font(_fonts.Families[0], size, style); 
            }
            catch 
            { 
                // اگر فونت نبود، از Tahoma استفاده می‌کند
                return new Font("Tahoma", size, style); 
            }
        }
    }
}
