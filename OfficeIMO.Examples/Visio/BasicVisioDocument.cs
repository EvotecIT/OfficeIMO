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
            VisioPage page = document.AddPage("Page1");
            VisioShape shape1 = new("Shape1");
            VisioShape shape2 = new("Shape2");
            page.Shapes.Add(shape1);
            page.Shapes.Add(shape2);
            page.Connectors.Add(new VisioConnector(shape1, shape2));

            // Saving not yet implemented; placeholder for future development.
        }
    }
}

