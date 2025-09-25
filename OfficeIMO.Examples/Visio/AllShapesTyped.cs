using System;
using System.IO;
using OfficeIMO.Visio;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Visio {
    /// <summary>
    /// Same as AllShapes but authored in the canonical Visio style:
    /// shapes are "typed" via masters and page instances carry minimal deltas.
    /// </summary>
    public static class AllShapesTyped {
        public static void Example_AllShapes_Typed(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - All shapes (typed, master-based)");
            string filePath = Path.Combine(folderPath, "All Shapes - Typed.vsdx");

            var doc = VisioDocument.Create(filePath);
            // Typed example: rely on masters; emit full instance cells for clarity
            doc.UseMastersByDefault = true;
            doc.WriteMasterDeltasOnly = false;
            doc.RequestRecalcOnOpen();

            // Prefer canonical masters from the provided asset to match real Visio output
            try {
                string baseDir = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "VisioTemplates"));
                string[] candidates = {
                    Path.Combine(baseDir, "DrawingWithShapes.vsdx"),
                    Path.Combine(baseDir, "DrawingWithLotsOfShapresAndArrows.vsdx")
                };
                foreach (var c in candidates) if (File.Exists(c)) doc.UseMastersFromTemplate(c);
            } catch { /* optional */ }

            var page = doc.AddPage("All Shapes (Typed)", 29.7, 21.0, VisioMeasurementUnit.Centimeters);

            double y1 = 15;
            double y2 = 9;

            var rect = page.AddRectangle(4.0, y1, 4.0, 2.5, "Rectangle");
            rect.FillColor = Color.LightBlue; rect.LineColor = Color.DarkBlue; rect.LineWeight = 0.02;
            var square = page.AddSquare(10.0, y1, 3.5, "Square");
            square.FillColor = Color.Beige; square.LineColor = Color.Brown; square.LineWeight = 0.02;
            var circle = page.AddCircle(16.0, y1, 3.5, "Circle", VisioMeasurementUnit.Centimeters);
            circle.FillColor = Color.LightGreen; circle.LineColor = Color.DarkGreen; circle.LineWeight = 0.02;
            var ellipse = page.AddEllipse(22.0, y1, 4.5, 3.0, "Ellipse");
            ellipse.FillColor = Color.MistyRose; ellipse.LineColor = Color.IndianRed; ellipse.LineWeight = 0.02;
            var diamond = page.AddDiamond(27.0, y1, 3.5, 3.0, "Diamond");
            diamond.FillColor = Color.Lavender; diamond.LineColor = Color.MediumPurple; diamond.LineWeight = 0.02;

            var triangle = page.AddTriangle(6.5, y2, 4.0, 3.0, "Triangle");
            triangle.FillColor = Color.LightGoldenrodYellow; triangle.LineColor = Color.Goldenrod; triangle.LineWeight = 0.02;
            var endShape = page.AddRectangle(18.0, y2, 4.5, 3.0, "End");
            endShape.FillColor = Color.LightGray; endShape.LineColor = Color.DimGray; endShape.LineWeight = 0.02;

            // No explicit connection-point handling; AddConnector ensures side glue internally when sides are specified

            page.AddConnector(rect, square, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left).LineColor = Color.Blue;
            page.AddConnector(square, circle, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left).LineColor = Color.DarkCyan;
            page.AddConnector(circle, ellipse, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left).LineColor = Color.DarkOrange;
            page.AddConnector(diamond, endShape, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left).LineColor = Color.DarkRed;

            // Page 3: connector variants (straight, right-angle, curved, dynamic)
            var page3 = doc.AddPage("Connectors", 29.7, 21.0, VisioMeasurementUnit.Centimeters);
            VisioShape MakeBox(double x, double y, double w, double h, string text, Color fill, Color line) {
                var s = page3.AddRectangle(x, y, w, h, text);
                s.FillColor = fill; s.LineColor = line; s.LineWeight = 0.02;
                return s;
            }
            double wB = 3.5, hB = 2.5;
            double yA = 16.0, yB = 13.5, yC = 11.0, yD = 8.5;
            double yB_R = yB + 1.0;
            double yC_R = yC - 1.0;
            double yD_R = yD + 0.5;
            double xL = 7.0, xR = 18.0;
            var aL = MakeBox(xL, yA, wB, hB, "Straight-L", Color.FromRgb(230, 240, 255), Color.SteelBlue);
            var aR = MakeBox(xR, yA, wB, hB, "Straight-R", Color.FromRgb(230, 240, 255), Color.SteelBlue);
            var bL = MakeBox(xL, yB, wB, hB, "Right-L", Color.FromRgb(255, 240, 230), Color.Peru);
            var bR = MakeBox(xR, yB_R, wB, hB, "Right-R", Color.FromRgb(255, 240, 230), Color.Peru);
            var cL = MakeBox(xL, yC, wB, hB, "Curved-L", Color.FromRgb(235, 255, 235), Color.ForestGreen);
            var cR = MakeBox(xR, yC_R, wB, hB, "Curved-R", Color.FromRgb(235, 255, 235), Color.ForestGreen);
            var dL = MakeBox(xL, yD, wB, hB, "Dynamic-L", Color.FromRgb(245, 245, 245), Color.DimGray);
            var dR = MakeBox(xR, yD_R, wB, hB, "Dynamic-R", Color.FromRgb(245, 245, 245), Color.DimGray);

            VisioConnector Glue(VisioShape left, VisioShape right, ConnectorKind kind, Color color) {
                var c = page3.AddConnector(left, right, kind, VisioSide.Right, VisioSide.Left);
                c.LineColor = color; c.LineWeight = 0.02; return c;
            }

            Glue(aL, aR, ConnectorKind.Straight, Color.SteelBlue);
            Glue(bL, bR, ConnectorKind.RightAngle, Color.Teal);
            Glue(cL, cR, ConnectorKind.Curved, Color.DarkOrange);
            Glue(dL, dR, ConnectorKind.Dynamic, Color.DarkRed);

            doc.Save();
            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
