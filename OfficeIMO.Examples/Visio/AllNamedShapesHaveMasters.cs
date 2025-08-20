using System;
using System.IO;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    public static class AllNamedShapesHaveMasters {
        public static void Run() {
            string filePath = Path.Combine(Path.GetTempPath(), "AllNamedShapesHaveMasters.vsdx");
            VisioDocument document = new();
            VisioPage page = document.AddPage("Page-1");

            VisioShape master1Shape = new("0", 0, 0, 2, 1, string.Empty);
            VisioMaster master1 = new("2", "Rectangle1", master1Shape);

            VisioShape master2Shape = new("0", 0, 0, 2, 1, string.Empty);
            VisioMaster master2 = new("3", "Rectangle2", master2Shape);

            page.Shapes.Add(new VisioShape("1", 1, 1, 2, 1, "First") { Master = master1 });
            page.Shapes.Add(new VisioShape("2", 4, 1, 2, 1, "Second") { Master = master2 });

            document.Save(filePath);
            Console.WriteLine(filePath);
        }
    }
}
