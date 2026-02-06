using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media.Imaging;

namespace lnk_to_EXE
{
    public static class IconExtractor
    {
        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr ExtractIcon(IntPtr hInst, string lpszExeFileName, int nIconIndex);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool DestroyIcon(IntPtr hIcon);

        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);

        public static BitmapSource? ExtractIcon(string filePath, int iconIndex)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                return null;

            try
            {
                IntPtr hIcon = ExtractIcon(IntPtr.Zero, filePath, iconIndex);

                if (hIcon == IntPtr.Zero || hIcon.ToInt32() == 1)
                    return null;

                try
                {
                    var icon = Icon.FromHandle(hIcon);
                    var bitmap = icon.ToBitmap();
                    
                    IntPtr hBitmap = bitmap.GetHbitmap();
                    try
                    {
                        return Imaging.CreateBitmapSourceFromHBitmap(
                            hBitmap,
                            IntPtr.Zero,
                            Int32Rect.Empty,
                            BitmapSizeOptions.FromEmptyOptions());
                    }
                    finally
                    {
                        DeleteObject(hBitmap);
                    }
                }
                finally
                {
                    DestroyIcon(hIcon);
                }
            }
            catch
            {
                return null;
            }
        }
    }
}
