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
                long slideWidth = presentation.SlideSize.WidthEmus;
                long margin = PowerPointUnits.Cm(1.5);
                long gutter = PowerPointUnits.Cm(1);
                long contentWidth = slideWidth - (2 * margin);
                long columnWidth = (contentWidth - gutter) / 2;
                long listTop = PowerPointUnits.Cm(5.5);
                long listHeight = PowerPointUnits.Cm(4);
                long calloutTop = listTop + listHeight + PowerPointUnits.Cm(1.5);
                long calloutHeight = PowerPointUnits.Cm(2.5);
                presentation.AsFluent()
                    .Slide(0, 0, s => {
                        s.Title("Fluent Presentation", tb => {
                            tb.FontSize = 32;
                            tb.Color = "1F4E79";
                        });
                        s.TextBox("Hello from fluent API", tb => {
                            tb.Left = margin;
                            tb.Top = listTop;
                            tb.Width = columnWidth;
                            tb.Height = listHeight;
                            tb.AddBullet("Built with builders");
                            tb.AddBullet("Configurable content");
                        });
                        s.Numbered(tb => {
                            tb.Left = margin + columnWidth + gutter;
                            tb.Top = listTop;
                            tb.Width = columnWidth;
                            tb.Height = listHeight;
                        }, "Step one", "Step two");
                        s.Shape(A.ShapeTypeValues.Rectangle, margin, calloutTop, contentWidth, calloutHeight,
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
