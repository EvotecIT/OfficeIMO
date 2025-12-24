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
                const double titleHeightCm = 1.6;
                const double calloutHeightCm = 2.6;
                const double gapCm = 0.6;
                PowerPointLayoutBox content = presentation.SlideSize.GetContentBoxCm(marginCm);
                PowerPointLayoutBox[] columns = presentation.SlideSize.GetColumnsCm(2, marginCm, gutterCm);
                double bodyHeightCm = content.HeightCm - titleHeightCm - calloutHeightCm - (gapCm * 2);
                double bodyTopCm = marginCm + titleHeightCm + gapCm;
                double calloutTopCm = bodyTopCm + bodyHeightCm + gapCm;
                long bodyTop = PowerPointUnits.FromCentimeters(bodyTopCm);
                long bodyHeight = PowerPointUnits.FromCentimeters(bodyHeightCm);
                long calloutTop = PowerPointUnits.FromCentimeters(calloutTopCm);
                long calloutHeight = PowerPointUnits.FromCentimeters(calloutHeightCm);
                PowerPointLayoutBox leftList = new(columns[0].Left, bodyTop, columns[0].Width, bodyHeight);
                PowerPointLayoutBox rightList = new(columns[1].Left, bodyTop, columns[1].Width, bodyHeight);
                presentation.AsFluent()
                    .Slide(0, 0, s => {
                        s.TitleCm("Fluent Presentation", marginCm, marginCm, content.WidthCm, titleHeightCm, tb => {
                            tb.FontSize = 32;
                            tb.Color = "1F4E79";
                        });
                        s.TextBox("Hello from fluent API", leftList.Left, leftList.Top, leftList.Width, leftList.Height,
                            tb => {
                            tb.AddBullet("Built with builders");
                            tb.AddBullet("Configurable content");
                            tb.ApplyAutoSpacing(lineSpacingMultiplier: 1.15);
                        });
                        s.Numbered(tb => {
                            tb.Left = rightList.Left;
                            tb.Top = rightList.Top;
                            tb.Width = rightList.Width;
                            tb.Height = rightList.Height;
                            tb.ApplyAutoSpacing(lineSpacingMultiplier: 1.15);
                        }, "Step one", "Step two");
                        s.Shape(A.ShapeTypeValues.Rectangle, content.Left, calloutTop, content.Width, calloutHeight,
                            shape => shape.Fill("E7F7FF").Stroke("007ACC", 2));
                        s.TextBox("Tip: fluent builders compose slides quickly while keeping layouts consistent.",
                            content.Left + PowerPointUnits.Cm(0.4), calloutTop + PowerPointUnits.Cm(0.3),
                            content.Width - PowerPointUnits.Cm(0.8), calloutHeight - PowerPointUnits.Cm(0.6),
                            tb => {
                                tb.FontSize = 16;
                                tb.Color = "1F4E79";
                            });
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
