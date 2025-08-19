using System;
using System.IO;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    public static class MasterShapes {
        public static void Run() {
            string filePath = Path.Combine(Path.GetTempPath(), "MasterShapes.vsdx");
            VisioDocument document = new();
            VisioPage page = document.AddPage("Page-1");
            page.Shapes.Add(new VisioShape("1", 1, 1, 2, 1, "First") { NameU = "Rectangle" });
            page.Shapes.Add(new VisioShape("2", 4, 1, 2, 1, "Second") { NameU = "Rectangle" });
            document.Save(filePath);
            Console.WriteLine(filePath);
        }
    }
}