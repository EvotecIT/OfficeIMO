using System;
using System.IO;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    /// <summary>
    /// Demonstrates loading an existing <see cref="VisioDocument"/>.
    /// </summary>
    public static class ReadVisioDocument {
        public static void Example_ReadVisio(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Reading basic document");
            string filePath = Path.Combine(folderPath, "Basic Visio.vsdx");

            VisioDocument document = VisioDocument.Load(filePath);
            foreach (VisioPage page in document.Pages) {
                Console.WriteLine($"Page: {page.Name}");
                foreach (VisioShape shape in page.Shapes) {
                    string master = shape.Master?.NameU ?? "None";
                    Console.WriteLine($"  Shape {shape.Id} {shape.NameU} {shape.Text} Master:{master}");
                }
            }

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
