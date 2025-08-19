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
            page.PageWidth = 11.69291338582677;
            page.PageHeight = 8.26771653543307;
            page.ViewCenterX = 5.8424184863857;
            page.ViewCenterY = 4.133858091015;
            page.Shapes.Add(new VisioShape("1") {
                NameU = "Rectangle",
                PinX = 2.047244040636296,
                PinY = 6.73228320203895,
                Width = 1.574803149606299,
                Height = 1.181102362204724
            });
            document.Save(filePath);

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}

