using System;
using System.IO;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Fluent;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates fluent API for building presentations.
    /// </summary>
    public static class FluentPowerPoint {
        public static void Example_FluentPowerPoint(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Creating presentation with fluent API");
            string filePath = Path.Combine(folderPath, "FluentPowerPoint.pptx");
            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                const double marginCm = 1.5;
                const double gutterCm = 1.0;
                PowerPointLayoutBox content = presentation.SlideSize.GetContentBoxCm(marginCm);
                PowerPointLayoutBox[] columns = presentation.SlideSize.GetColumnsCm(2, marginCm, gutterCm);
                PowerPointLayoutBox[] rows = presentation.SlideSize.GetRowsCm(2, marginCm, gutterCm);
                PowerPointLayoutBox listRow = rows[0];
                PowerPointLayoutBox calloutRow = rows[1];
                PowerPointLayoutBox leftList = new(columns[0].Left, listRow.Top, columns[0].Width, listRow.Height);
                PowerPointLayoutBox rightList = new(columns[1].Left, listRow.Top, columns[1].Width, listRow.Height);
                presentation.AsFluent()
                    .Slide(0, 0, s => {
                        s.Title("Fluent Presentation", tb => {
                            tb.FontSize = 32;
                            tb.Color = "1F4E79";
                        });
                        s.TextBox("Hello from fluent API", leftList.Left, leftList.Top, leftList.Width, leftList.Height,
                            tb => {
                            tb.AddBullet("Built with builders");
                            tb.AddBullet("Configurable content");
                        });
                        s.Numbered(tb => {
                            tb.Left = rightList.Left;
                            tb.Top = rightList.Top;
                            tb.Width = rightList.Width;
                            tb.Height = rightList.Height;
                        }, "Step one", "Step two");
                        s.Shape(A.ShapeTypeValues.Rectangle, content.Left, calloutRow.Top, content.Width,
                            calloutRow.Height,
                            shape => shape.Fill("E7F7FF").Stroke("007ACC", 2));
                        s.Notes("Example notes");
                    })
                    .Slide(s => s.Title("Second Slide"))
                    .End()
                    .Save();
            }

            Helpers.Open(filePath, openPowerPoint);
        }
    }
}
