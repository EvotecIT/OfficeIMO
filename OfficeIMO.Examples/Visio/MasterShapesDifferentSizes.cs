using System;
using System.IO;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    public static class MasterShapesDifferentSizes {
        public static void Run() {
            string filePath = Path.Combine(Path.GetTempPath(), "MasterShapesDifferentSizes.vsdx");
            VisioDocument document = new();
            VisioPage page = document.AddPage("Page-1");

            VisioShape masterShape = new("0", 0, 0, 2, 1, string.Empty);
            VisioMaster rectangle = new("2", "Rectangle", masterShape);

            VisioShape shape1 = new("1", 1, 1, 2, 1, "Same") { Master = rectangle };
            VisioShape shape2 = new("2", 4, 1, 3, 2, "Different") { Master = rectangle };

            page.Shapes.Add(shape1);
            page.Shapes.Add(shape2);

            document.Save(filePath);
            Console.WriteLine(filePath);
        }
    }
}
