using System;
using System.IO;
using OfficeIMO.Visio;
using Color = OfficeIMO.OfficeColor;

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

            // Learn supported master names from sample VSDX files; generated output still comes from OfficeIMO code.
            try {
                string baseDir = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "VisioTemplates"));
                string[] candidates = {
                    Path.Combine(baseDir, "DrawingWithShapes.vsdx"),
                    Path.Combine(baseDir, "DrawingWithLotsOfShapresAndArrows.vsdx")
                };
                foreach (var c in candidates) if (File.Exists(c)) doc.LearnMastersFromVsdx(c);
            } catch { /* optional */ }

            var page = doc.AddPage("All Shapes (Typed)", OfficeLength.Centimeters(29.7), OfficeLength.Centimeters(21.0), scaleUnit: VisioMeasurementUnit.Centimeters);

            double y1 = 15;
            double y2 = 9;

            var rect = page.AddRectangle(OfficeLength.Centimeters(4.0), OfficeLength.Centimeters(y1), OfficeLength.Centimeters(4.0), OfficeLength.Centimeters(2.5), "Rectangle");
            rect.FillColor = Color.LightBlue; rect.LineColor = Color.DarkBlue; rect.LineWeight = 0.02;
            var square = page.AddSquare(OfficeLength.Centimeters(10.0), OfficeLength.Centimeters(y1), OfficeLength.Centimeters(3.5), "Square");
            square.FillColor = Color.Beige; square.LineColor = Color.Brown; square.LineWeight = 0.02;
            var circle = page.AddCircle(OfficeLength.Centimeters(16.0), OfficeLength.Centimeters(y1), OfficeLength.Centimeters(3.5), "Circle");
            circle.FillColor = Color.LightGreen; circle.LineColor = Color.DarkGreen; circle.LineWeight = 0.02;
            var ellipse = page.AddEllipse(OfficeLength.Centimeters(22.0), OfficeLength.Centimeters(y1), OfficeLength.Centimeters(4.5), OfficeLength.Centimeters(3.0), "Ellipse");
            ellipse.FillColor = Color.MistyRose; ellipse.LineColor = Color.IndianRed; ellipse.LineWeight = 0.02;
            var diamond = page.AddDiamond(OfficeLength.Centimeters(27.0), OfficeLength.Centimeters(y1), OfficeLength.Centimeters(3.5), OfficeLength.Centimeters(3.0), "Diamond");
            diamond.FillColor = Color.Lavender; diamond.LineColor = Color.MediumPurple; diamond.LineWeight = 0.02;

            var triangle = page.AddTriangle(OfficeLength.Centimeters(6.5), OfficeLength.Centimeters(y2), OfficeLength.Centimeters(4.0), OfficeLength.Centimeters(3.0), "Triangle");
            triangle.FillColor = Color.LightGoldenrodYellow; triangle.LineColor = Color.Goldenrod; triangle.LineWeight = 0.02;
            var endShape = page.AddRectangle(OfficeLength.Centimeters(18.0), OfficeLength.Centimeters(y2), OfficeLength.Centimeters(4.5), OfficeLength.Centimeters(3.0), "End");
            endShape.FillColor = Color.LightGray; endShape.LineColor = Color.DimGray; endShape.LineWeight = 0.02;

            // No explicit connection-point handling; AddConnector ensures side glue internally when sides are specified

            page.AddConnector(rect, square, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left).LineColor = Color.Blue;
            page.AddConnector(square, circle, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left).LineColor = Color.DarkCyan;
            page.AddConnector(circle, ellipse, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left).LineColor = Color.DarkOrange;
            page.AddConnector(diamond, endShape, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left).LineColor = Color.DarkRed;

            // Page 3: connector variants (straight, right-angle, curved, dynamic)
            var page3 = doc.AddPage("Connectors", OfficeLength.Centimeters(29.7), OfficeLength.Centimeters(21.0), scaleUnit: VisioMeasurementUnit.Centimeters);
            VisioShape MakeBox(double x, double y, double w, double h, string text, Color fill, Color line) {
                var s = page3.AddRectangle(OfficeLength.Centimeters(x), OfficeLength.Centimeters(y), OfficeLength.Centimeters(w), OfficeLength.Centimeters(h), text);
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
