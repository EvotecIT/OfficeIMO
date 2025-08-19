using System;
using System.IO;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    /// <summary>
    /// Demonstrates basic <see cref="VisioDocument"/> usage.
    /// </summary>
    public static class BasicVisioDocument {
        public static void Example_BasicVisio(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Creating basic document");
            string filePath = Path.Combine(folderPath, "Basic Visio.vsdx");

            VisioDocument document = new();
            VisioPage page = document.AddPage("Page-1");
            page.Shapes.Add(new VisioShape("1", 1, 1, 2, 1, string.Empty) { NameU = "Rectangle" });
            document.Save(filePath);

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}

