using System;
using System.IO;
using System.IO.Packaging;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    /// <summary>
    /// Demonstrates presence of theme and windows parts.
    /// </summary>
    public static class ThemeAndWindows {
        public static void Example_ThemeAndWindows(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Theme and Windows");
            string filePath = Path.Combine(folderPath, "Theme and Windows.vsdx");

            VisioDocument document = new();
            document.Theme = new VisioTheme { Name = "Office Theme" };
            document.AddPage("Page-1");
            document.Save(filePath);

            using Package package = Package.Open(filePath, FileMode.Open, FileAccess.Read);
            Console.WriteLine(package.PartExists(new Uri("/visio/theme/theme1.xml", UriKind.Relative)));
            Console.WriteLine(package.PartExists(new Uri("/visio/windows.xml", UriKind.Relative)));

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
