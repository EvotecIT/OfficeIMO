using System;
using System.IO;
using OfficeIMO.Visio;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Visio {
    /// <summary>
    /// Comprehensive example demonstrating various shapes with colors, styles, and connectors.
    /// </summary>
    public static class ComprehensiveColoredShapes {
        public static void Example_ComprehensiveColoredShapes(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Comprehensive Colored Shapes");
            string filePath = Path.Combine(folderPath, "Comprehensive Colored Shapes.vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            document.RequestRecalcOnOpen();
            VisioPage page = document.AddPage("Colored Shapes", 29.7, 21, VisioMeasurementUnit.Centimeters);

            // Row 1: Basic colored rectangles
            var blueRect = new VisioShape("1", 1.5, 7, 1.5, 1, "Blue") {
                FillColor = Color.LightBlue,
                LineColor = Color.DarkBlue,
                LineWeight = 0.02,
                LinePattern = 1, // Solid
                FillPattern = 1  // Solid
            };
            page.Shapes.Add(blueRect);

            var greenRect = new VisioShape("2", 3.5, 7, 1.5, 1, "Green") {
                FillColor = Color.LightGreen,
                LineColor = Color.DarkGreen,
                LineWeight = 0.03,
                LinePattern = 1, // Solid
                FillPattern = 1  // Solid
            };
            page.Shapes.Add(greenRect);

            var redRect = new VisioShape("3", 5.5, 7, 1.5, 1, "Red") {
                FillColor = Color.LightPink,
                LineColor = Color.DarkRed,
                LineWeight = 0.025,
                LinePattern = 1, // Solid
                FillPattern = 1  // Solid
            };
            page.Shapes.Add(redRect);

            var yellowRect = new VisioShape("4", 7.5, 7, 1.5, 1, "Yellow") {
                FillColor = Color.LightYellow,
                LineColor = Color.Orange,
                LineWeight = 0.02,
                LinePattern = 1, // Solid
                FillPattern = 1  // Solid
            };
            page.Shapes.Add(yellowRect);

            var purpleRect = new VisioShape("5", 9.5, 7, 1.5, 1, "Purple") {
                FillColor = Color.Lavender,
                LineColor = Color.Purple,
                LineWeight = 0.03,
                LinePattern = 1, // Solid
                FillPattern = 1  // Solid
            };
            page.Shapes.Add(purpleRect);

            // Row 2: Different line patterns
            var dashedRect = new VisioShape("6", 1.5, 5, 1.5, 1, "Dashed") {
                FillColor = Color.LightCyan,
                LineColor = Color.Black,
                LineWeight = 0.02,
                LinePattern = 2, // Dashed
                FillPattern = 1  // Solid
            };
            page.Shapes.Add(dashedRect);

            var thickBorderRect = new VisioShape("7", 3.5, 5, 1.5, 1, "Thick") {
                FillColor = Color.Beige,
                LineColor = Color.Brown,
                LineWeight = 0.08, // Very thick
                LinePattern = 1, // Solid
                FillPattern = 1  // Solid
            };
            page.Shapes.Add(thickBorderRect);

            var noFillRect = new VisioShape("8", 5.5, 5, 1.5, 1, "No Fill") {
                FillColor = Color.Transparent,
                LineColor = Color.Navy,
                LineWeight = 0.04,
                LinePattern = 1, // Solid
                FillPattern = 0  // No fill
            };
            page.Shapes.Add(noFillRect);

            var gradientSimRect = new VisioShape("9", 7.5, 5, 1.5, 1, "Gradient") {
                FillColor = Color.FromRgb(200, 220, 255), // Light blue gradient simulation
                LineColor = Color.SteelBlue,
                LineWeight = 0.025,
                LinePattern = 1, // Solid
                FillPattern = 1  // Solid
            };
            page.Shapes.Add(gradientSimRect);

            var grayRect = new VisioShape("10", 9.5, 5, 1.5, 1, "Gray") {
                FillColor = Color.LightGray,
                LineColor = Color.DarkGray,
                LineWeight = 0.02,
                LinePattern = 1, // Solid
                FillPattern = 1  // Solid
            };
            page.Shapes.Add(grayRect);

            // Row 3: Larger shapes with data
            var dataShape1 = new VisioShape("11", 2.5, 3, 2, 1.5, "Data Box 1") {
                FillColor = Color.FromRgb(255, 245, 230), // Light peach
                LineColor = Color.Coral,
                LineWeight = 0.03,
                LinePattern = 1, // Solid
                FillPattern = 1  // Solid
            };
            dataShape1.Data["Type"] = "Process";
            dataShape1.Data["Status"] = "Active";
            dataShape1.Data["Owner"] = "Team A";
            page.Shapes.Add(dataShape1);

            var dataShape2 = new VisioShape("12", 5.5, 3, 2, 1.5, "Data Box 2") {
                FillColor = Color.FromRgb(240, 255, 240), // Honeydew
                LineColor = Color.ForestGreen,
                LineWeight = 0.03,
                LinePattern = 1, // Solid
                FillPattern = 1  // Solid
            };
            dataShape2.Data["Type"] = "Database";
            dataShape2.Data["Status"] = "Online";
            dataShape2.Data["Size"] = "Large";
            page.Shapes.Add(dataShape2);

            var dataShape3 = new VisioShape("13", 8.5, 3, 2, 1.5, "Data Box 3") {
                FillColor = Color.FromRgb(255, 240, 245), // Lavender blush
                LineColor = Color.MediumVioletRed,
                LineWeight = 0.03,
                LinePattern = 1, // Solid
                FillPattern = 1  // Solid
            };
            dataShape3.Data["Type"] = "Output";
            dataShape3.Data["Status"] = "Complete";
            dataShape3.Data["Format"] = "PDF";
            page.Shapes.Add(dataShape3);

            // Add connectors between data shapes
            var connector1 = new VisioConnector(dataShape1, dataShape2) {
                LineColor = Color.Blue,
                LineWeight = 0.02,
                LinePattern = 1,
                EndArrow = EndArrow.Arrow
            };
            page.Connectors.Add(connector1);

            var connector2 = new VisioConnector(dataShape2, dataShape3) {
                LineColor = Color.Green,
                LineWeight = 0.02,
                LinePattern = 1,
                EndArrow = EndArrow.Arrow
            };
            page.Connectors.Add(connector2);

            // Row 4: Special effects
            var shadowShape = new VisioShape("14", 2, 1, 1.8, 1.2, "Shadow") {
                FillColor = Color.FromRgb(240, 240, 240),
                LineColor = Color.Black,
                LineWeight = 0.01,
                LinePattern = 1, // Solid
                FillPattern = 1  // Solid
            };
            page.Shapes.Add(shadowShape);

            var roundedShape = new VisioShape("15", 4.5, 1, 1.8, 1.2, "Rounded") {
                FillColor = Color.SkyBlue,
                LineColor = Color.DodgerBlue,
                LineWeight = 0.025,
                LinePattern = 1, // Solid
                FillPattern = 1  // Solid
            };
            page.Shapes.Add(roundedShape);

            var doubleLineShape = new VisioShape("16", 7, 1, 1.8, 1.2, "Double") {
                FillColor = Color.Wheat,
                LineColor = Color.SaddleBrown,
                LineWeight = 0.015,
                LinePattern = 3, // Double line pattern if supported
                FillPattern = 1  // Solid
            };
            page.Shapes.Add(doubleLineShape);

            var transparentShape = new VisioShape("17", 9.5, 1, 1.8, 1.2, "50% Trans") {
                FillColor = Color.FromRgba(0, 0, 255, 128), // Semi-transparent blue
                LineColor = Color.Blue,
                LineWeight = 0.02,
                LinePattern = 1, // Solid
                FillPattern = 1  // Solid
            };
            page.Shapes.Add(transparentShape);

            document.Save();

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}