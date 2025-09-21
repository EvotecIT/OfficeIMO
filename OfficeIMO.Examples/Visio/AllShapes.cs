using System;
using System.IO;
using OfficeIMO.Visio;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Visio {
    /// <summary>
    /// Creates a single diagram that showcases all built-in shapes (one by one)
    /// and a few connector types for quick visual inspection.
    /// </summary>
    public static class AllShapes {
        public static void Example_AllShapes(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - All shapes (one by one)");
            string filePath = Path.Combine(folderPath, "All Shapes.vsdx");

            var doc = VisioDocument.Create(filePath);
            // Prefer masters for built-in shapes to match Visio geometry
            doc.UseMastersByDefault = true;
            doc.WriteMasterDeltasOnly = false;
            try {
                string baseDir = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "VisioTemplates"));
                string[] candidates = {
                    Path.Combine(baseDir, "DrawingWithShapes.vsdx"),
                    Path.Combine(baseDir, "DrawingWithLotsOfShapresAndArrows.vsdx")
                };
                foreach (var c in candidates) if (File.Exists(c)) doc.UseMastersFromTemplate(c);
            } catch { /* optional */ }
            doc.RequestRecalcOnOpen();

            // A4 landscape (29.7 x 21.0 cm) — aligns with templates; page.DefaultUnit=Centimeters
            var page = doc.AddPage("All Shapes", 29.7, 21.0, VisioMeasurementUnit.Centimeters);

            // Row 1: Basic shapes (values in centimeters)
            double y1 = 15; // upper row
            double y2 = 9;  // lower row

            var rect = page.AddRectangle(4.0, y1, 4.0, 2.5, "Rectangle"); rect.FillColor = Color.LightBlue; rect.LineColor = Color.DarkBlue; rect.LineWeight = 0.02;
            var square = page.AddSquare(10.0, y1, 3.5, "Square"); square.FillColor = Color.Beige; square.LineColor = Color.Brown; square.LineWeight = 0.02;
            var circle = page.AddCircle(16.0, y1, 3.5, "Circle", VisioMeasurementUnit.Centimeters); circle.FillColor = Color.LightGreen; circle.LineColor = Color.DarkGreen; circle.LineWeight = 0.02;
            var ellipse = page.AddEllipse(22.0, y1, 4.5, 3.0, "Ellipse"); ellipse.FillColor = Color.MistyRose; ellipse.LineColor = Color.IndianRed; ellipse.LineWeight = 0.02;
            var diamond = page.AddDiamond(27.0, y1, 3.5, 3.0, "Diamond"); diamond.FillColor = Color.Lavender; diamond.LineColor = Color.MediumPurple; diamond.LineWeight = 0.02;

            // Row 2: Triangle plus a few connectors (cm)
            var triangle = page.AddTriangle(6.5, y2, 4.0, 3.0, "Triangle"); triangle.FillColor = Color.LightGoldenrodYellow; triangle.LineColor = Color.Goldenrod; triangle.LineWeight = 0.02;
            var endShape = page.AddRectangle(18.0, y2, 4.5, 3.0, "End"); endShape.FillColor = Color.LightGray; endShape.LineColor = Color.DimGray; endShape.LineWeight = 0.02;

            // No explicit connection-point handling; AddConnector ensures side glue internally when sides are specified

            // Connectors (explicit end points to side centers)
            page.AddConnector(rect, square, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left).LineColor = Color.Blue;
            page.AddConnector(square, circle, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left).LineColor = Color.DarkCyan;
            page.AddConnector(circle, ellipse, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left).LineColor = Color.DarkOrange;
            page.AddConnector(diamond, endShape, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left).LineColor = Color.DarkRed;

            // Page 2: rotations and angled connectors
            var page2 = doc.AddPage("Rotations", 29.7, 21.0, VisioMeasurementUnit.Centimeters);
            double y3 = 12;
            var r2 = page2.AddRectangle(5, y3, 4, 2, "Rect 30deg"); r2.NameU = "Rectangle"; r2.Angle = Math.PI / 6;
            var d2 = page2.AddDiamond(12, y3, 3.5, 3.0, "Diamond -45deg"); d2.Angle = -Math.PI / 4;
            var c2 = page2.AddCircle(20, y3, 3.0, "Circle 15deg"); c2.Angle = Math.PI / 12;
            // angled connectors among rotated shapes (use API to glue to shape sides)
            page2.AddConnector(r2, d2, ConnectorKind.RightAngle).LineColor = Color.Gray;
            page2.AddConnector(d2, c2, ConnectorKind.Straight).LineColor = Color.Gray;
            page2.AddConnector(c2, r2, ConnectorKind.Dynamic).LineColor = Color.Gray;

            // Save
            // Page 3: connector variants (straight, right-angle, curved, dynamic)
            var page3 = doc.AddPage("Connectors", 29.7, 21.0, VisioMeasurementUnit.Centimeters);
            // Add a labeled rectangle box quickly
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
