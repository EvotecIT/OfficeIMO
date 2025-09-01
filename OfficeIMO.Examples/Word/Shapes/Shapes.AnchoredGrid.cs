using System;
using OfficeIMO.Word;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Word {
    internal static partial class Shapes {
        internal static void Example_AnchoredShapesGrid(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with anchored DrawingML shapes grid");
            string filePath = System.IO.Path.Combine(folderPath, "DocumentWithAnchoredShapesGrid.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Anchored DrawingML Shapes (tight grid)");

                // Grid settings (points)
                double startX = 30;   // left margin
                double startY = 80;   // top margin
                double colStep = 110; // horizontal spacing
                double rowStep = 100; // vertical spacing

                // Sizes by shape category
                (double w, double h) Rect(double w, double h) => (w, h);

                var shapes = new (ShapeType type, (double w, double h) size, Color fill, Color stroke)[] {
                    (ShapeType.Rectangle, Rect(90, 50), Color.LightSkyBlue, Color.DarkBlue),
                    (ShapeType.Ellipse, Rect(80, 50), Color.LightGreen, Color.DarkGreen),
                    (ShapeType.RoundedRectangle, Rect(90, 50), Color.Khaki, Color.Olive),
                    (ShapeType.Triangle, Rect(70, 60), Color.Coral, Color.DarkRed),
                    (ShapeType.Diamond, Rect(70, 70), Color.Plum, Color.Purple),
                    (ShapeType.Hexagon, Rect(90, 60), Color.SandyBrown, Color.SaddleBrown),
                    (ShapeType.RightArrow, Rect(100, 40), Color.CornflowerBlue, Color.SteelBlue),
                    (ShapeType.LeftArrow, Rect(100, 40), Color.Gold, Color.DarkGoldenrod),
                    (ShapeType.UpArrow, Rect(60, 90), Color.LightPink, Color.HotPink),
                    (ShapeType.DownArrow, Rect(60, 90), Color.LightGray, Color.DimGray),
                    (ShapeType.Heart, Rect(80, 70), Color.Pink, Color.HotPink),
                    (ShapeType.Cloud, Rect(110, 70), Color.WhiteSmoke, Color.Gray),
                    (ShapeType.Donut, Rect(90, 90), Color.Goldenrod, Color.Maroon),
                    (ShapeType.Can, Rect(80, 100), Color.LightSteelBlue, Color.SteelBlue),
                    (ShapeType.Cube, Rect(90, 90), Color.MediumPurple, Color.Indigo)
                };

                int cols = 5; // tighter grid
                var placed = new List<(double left, double top, double w, double h, string name)>();
                for (int i = 0; i < shapes.Length; i++) {
                    int row = i / cols;
                    int col = i % cols;
                    double left = startX + col * colStep;
                    double top = startY + row * rowStep;

                    var (type, size, fill, stroke) = shapes[i];
                    var p = document.AddParagraph("");
                    var shp = p.AddShapeDrawing(type, size.w, size.h, left, top);
                    shp.FillColor = fill;
                    shp.StrokeColor = stroke;
                    shp.StrokeWeight = 1.5;
                    placed.Add((left, top, size.w, size.h, type.ToString()));
                }

                // Add labels centered above each shape using text boxes (absolute positioning)
                foreach (var item in placed) {
                    double labelWpt = Math.Max(70, item.w); // ensure label width reasonable
                    double labelHpt = 14;
                    double labelLeftPt = item.left + (item.w - labelWpt) / 2.0;
                    double labelTopPt = item.top - (labelHpt + 8); // a bit above

                    var lp = document.AddParagraph("");
                    var tb = lp.AddTextBox("", WrapTextImage.InFrontOfText);
                    tb.WrapText = WrapTextImage.InFrontOfText;
                    tb.HorizontalPositionRelativeFrom = DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalRelativePositionValues.Page;
                    tb.VerticalPositionRelativeFrom = DocumentFormat.OpenXml.Drawing.Wordprocessing.VerticalRelativePositionValues.Page;
                    int PtToEmus(double pt) => (int)System.Math.Round(pt * 12700.0);
                    double PtToCm(double pt) => pt * 2.54 / 72.0;
                    tb.HorizontalPositionOffset = PtToEmus(labelLeftPt);
                    tb.VerticalPositionOffset = PtToEmus(labelTopPt);
                    tb.WidthCentimeters = PtToCm(labelWpt);
                    tb.HeightCentimeters = PtToCm(labelHpt);
                    var paragraph = tb.Paragraphs.FirstOrDefault() ?? lp;
                    paragraph.Text = item.name;
                }

                // Add simple horizontal connectors (RightArrow) between adjacent shapes per row
                for (int i = 0; i < placed.Count; i++) {
                    int row = i / cols;
                    int col = i % cols;
                    if (col == cols - 1) continue; // last col, no connector to the right
                    var from = placed[i];
                    var to = placed[i + 1];
                    double gapLeft = from.left + from.w;
                    double available = to.left - gapLeft;
                    double arrowH = 12;
                    double arrowW = Math.Max(available - 8, 8);
                    double arrowTop = from.top + (from.h - arrowH) / 2.0;

                    var cp = document.AddParagraph("");
                    var conn = cp.AddShapeDrawing(ShapeType.RightArrow, arrowW, arrowH, gapLeft + 4, arrowTop);
                    conn.FillColor = SixLabors.ImageSharp.Color.Gray;
                    conn.StrokeColor = SixLabors.ImageSharp.Color.DimGray;
                    conn.StrokeWeight = 1;
                }

                document.Save(openWord);
                OfficeIMO.Examples.Utils.Validation.ValidateDoc(filePath);
            }
        }
    }
}
