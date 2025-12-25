using System;
using System.IO;
using OfficeIMO.PowerPoint;
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
                const double calloutHeightCm = 2.4;
                const double gapCm = 0.6;
                PowerPointLayoutBox content = presentation.SlideSize.GetContentBoxCm(marginCm);
                double bodyHeightCm = content.HeightCm - titleHeightCm - calloutHeightCm - (gapCm * 2);
                double bodyTopCm = content.TopCm + titleHeightCm + gapCm;
                long bodyTop = PowerPointUnits.FromCentimeters(bodyTopCm);
                long bodyHeight = PowerPointUnits.FromCentimeters(bodyHeightCm);
                long calloutTop = bodyTop + bodyHeight + PowerPointUnits.FromCentimeters(gapCm);
                long calloutHeight = PowerPointUnits.FromCentimeters(calloutHeightCm);

                PowerPointLayoutBox body = PowerPointLayoutBox.FromCentimeters(content.LeftCm, bodyTopCm, content.WidthCm, bodyHeightCm);
                PowerPointLayoutBox[] columns = body.SplitColumnsCm(2, gutterCm);
                PowerPointLayoutBox[] leftRows = columns[0].SplitRowsCm(2, 0.4);

                presentation.AsFluent()
                    .Slide(0, 0, s => {
                        s.TitleCm("Fluent Presentation", marginCm, marginCm, content.WidthCm, titleHeightCm, tb => {
                            tb.FontSize = 32;
                            tb.Color = "1F4E79";
                        });
                        s.TextBox("Hello from fluent API", leftRows[0].Left, leftRows[0].Top, leftRows[0].Width, leftRows[0].Height,
                            tb => {
                                tb.FontSize = 18;
                                tb.Color = "1F4E79";
                            });
                        s.TextBox(string.Empty, leftRows[1].Left, leftRows[1].Top, leftRows[1].Width, leftRows[1].Height,
                            tb => {
                                tb.AddBullet("Built with builders");
                                tb.AddBullet("Configurable content");
                                tb.AddBullet("Readable layout helpers");
                                tb.ApplyAutoSpacing(lineSpacingMultiplier: 1.15, spaceAfterPoints: 2);
                            });
                        s.Numbered(tb => {
                            tb.Left = columns[1].Left;
                            tb.Top = columns[1].Top;
                            tb.Width = columns[1].Width;
                            tb.Height = columns[1].Height;
                            tb.ApplyAutoSpacing(lineSpacingMultiplier: 1.15, spaceAfterPoints: 2);
                        }, "Step one", "Step two", "Step three");
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
