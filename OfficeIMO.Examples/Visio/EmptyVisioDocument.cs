using System;
using System.IO;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    /// <summary>
    /// Demonstrates creating an empty Visio document without theme.
    /// </summary>
    public static class EmptyVisioDocument {
        public static void Example_EmptyVisio(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Creating empty document");
            string filePath = Path.Combine(folderPath, "Empty Visio.vsdx");

            VisioDocument document = new();
            document.AddPage("Page-1");
            document.Save(filePath);

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
