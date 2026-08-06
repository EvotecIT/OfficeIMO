using System;
using OfficeIMO.Word;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Examples.Word {
    internal static partial class Shapes {
        internal static void Example_AnchoredShapesGrid(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with anchored DrawingML shapes grid");
            string filePath = System.IO.Path.Combine(folderPath, "DocumentWithAnchoredShapesGrid.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Anchored DrawingML Shapes (tight grid)");

                // Define shapes succinctly and build the anchored grid via a helper
                var shapes = new List<(WordShapeType type, double wPt, double hPt, Color fill, Color stroke, string label)> {
                    (WordShapeType.Rectangle, 90, 50, Color.LightSkyBlue, Color.DarkBlue, "Rectangle"),
                    (WordShapeType.Ellipse, 80, 50, Color.LightGreen, Color.DarkGreen, "Ellipse"),
                    (WordShapeType.RoundedRectangle, 90, 50, Color.Khaki, Color.Olive, "RoundedRectangle"),
                    (WordShapeType.Triangle, 70, 60, Color.Coral, Color.DarkRed, "Triangle"),
                    (WordShapeType.Diamond, 70, 70, Color.Plum, Color.Purple, "Diamond"),
                    (WordShapeType.Hexagon, 90, 60, Color.SandyBrown, Color.SaddleBrown, "Hexagon"),
                    (WordShapeType.RightArrow, 100, 40, Color.CornflowerBlue, Color.SteelBlue, "RightArrow"),
                    (WordShapeType.LeftArrow, 100, 40, Color.Gold, Color.DarkGoldenrod, "LeftArrow"),
                    (WordShapeType.UpArrow, 60, 90, Color.LightPink, Color.HotPink, "UpArrow"),
                    (WordShapeType.DownArrow, 60, 90, Color.LightGray, Color.DimGray, "DownArrow"),
                    (WordShapeType.Heart, 80, 70, Color.Pink, Color.HotPink, "Heart"),
                    (WordShapeType.Cloud, 110, 70, Color.WhiteSmoke, Color.Gray, "Cloud"),
                    (WordShapeType.Donut, 90, 90, Color.Goldenrod, Color.Maroon, "Donut"),
                    (WordShapeType.Can, 80, 100, Color.LightSteelBlue, Color.SteelBlue, "Can"),
                    (WordShapeType.Cube, 90, 90, Color.MediumPurple, Color.Indigo, "Cube"),
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

                document.Save();
                if (openWord) document.OpenInApplication();
                OfficeIMO.Examples.Utils.Validation.ValidateDoc(filePath);
            }
        }
    }
}
