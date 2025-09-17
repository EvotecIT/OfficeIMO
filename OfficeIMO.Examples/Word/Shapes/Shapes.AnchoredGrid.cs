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

                // Define shapes succinctly and build the anchored grid via a helper
                var shapes = new List<(ShapeType type, double wPt, double hPt, Color fill, Color stroke, string label)> {
                    (ShapeType.Rectangle, 90, 50, Color.LightSkyBlue, Color.DarkBlue, "Rectangle"),
                    (ShapeType.Ellipse, 80, 50, Color.LightGreen, Color.DarkGreen, "Ellipse"),
                    (ShapeType.RoundedRectangle, 90, 50, Color.Khaki, Color.Olive, "RoundedRectangle"),
                    (ShapeType.Triangle, 70, 60, Color.Coral, Color.DarkRed, "Triangle"),
                    (ShapeType.Diamond, 70, 70, Color.Plum, Color.Purple, "Diamond"),
                    (ShapeType.Hexagon, 90, 60, Color.SandyBrown, Color.SaddleBrown, "Hexagon"),
                    (ShapeType.RightArrow, 100, 40, Color.CornflowerBlue, Color.SteelBlue, "RightArrow"),
                    (ShapeType.LeftArrow, 100, 40, Color.Gold, Color.DarkGoldenrod, "LeftArrow"),
                    (ShapeType.UpArrow, 60, 90, Color.LightPink, Color.HotPink, "UpArrow"),
                    (ShapeType.DownArrow, 60, 90, Color.LightGray, Color.DimGray, "DownArrow"),
                    (ShapeType.Heart, 80, 70, Color.Pink, Color.HotPink, "Heart"),
                    (ShapeType.Cloud, 110, 70, Color.WhiteSmoke, Color.Gray, "Cloud"),
                    (ShapeType.Donut, 90, 90, Color.Goldenrod, Color.Maroon, "Donut"),
                    (ShapeType.Can, 80, 100, Color.LightSteelBlue, Color.SteelBlue, "Can"),
                    (ShapeType.Cube, 90, 90, Color.MediumPurple, Color.Indigo, "Cube"),
                };

                OfficeIMO.Examples.Utils.AnchoredDiagram.BuildGrid(
                    document,
                    shapes,
                    cols: 5,
                    startXpt: 30,
                    startYpt: 80,
                    colStepPt: 110,
                    rowStepPt: 100,
                    addLabels: true,
                    addHorizontalConnectors: true,
                    addVerticalConnectors: true,
                    elbowConnector: (0, 7),
                    legend: "Legend: → row neighbor, ↓ column neighbor; labels above; anchored shapes on grid."
                );

                document.Save(openWord);
                OfficeIMO.Examples.Utils.Validation.ValidateDoc(filePath);
            }
        }
    }
}
