using System;
using System.IO;
using OfficeIMO.Visio;
using Color = OfficeIMO.OfficeColor;

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
            VisioPage page = document.AddPage("Colored Shapes", OfficeLength.Centimeters(29.7), OfficeLength.Centimeters(21), scaleUnit: VisioMeasurementUnit.Centimeters);

            // Row 1: Basic colored rectangles
            var blueRect = new VisioShape("1", OfficeLength.Inches(1.5), OfficeLength.Inches(7), OfficeLength.Inches(1.5), OfficeLength.Inches(1), "Blue");
            page.Shapes.Add(blueRect);

            var greenRect = new VisioShape("2", OfficeLength.Inches(3.5), OfficeLength.Inches(7), OfficeLength.Inches(1.5), OfficeLength.Inches(1), "Green");
            page.Shapes.Add(greenRect);

            var redRect = new VisioShape("3", OfficeLength.Inches(5.5), OfficeLength.Inches(7), OfficeLength.Inches(1.5), OfficeLength.Inches(1), "Red");
            page.Shapes.Add(redRect);

            var yellowRect = new VisioShape("4", OfficeLength.Inches(7.5), OfficeLength.Inches(7), OfficeLength.Inches(1.5), OfficeLength.Inches(1), "Yellow");
            page.Shapes.Add(yellowRect);

            var purpleRect = new VisioShape("5", OfficeLength.Inches(9.5), OfficeLength.Inches(7), OfficeLength.Inches(1.5), OfficeLength.Inches(1), "Purple");
            page.Shapes.Add(purpleRect);

            // Row 2: Different line patterns
            var dashedRect = new VisioShape("6", OfficeLength.Inches(1.5), OfficeLength.Inches(5), OfficeLength.Inches(1.5), OfficeLength.Inches(1), "Dashed");
            page.Shapes.Add(dashedRect);

            var thickBorderRect = new VisioShape("7", OfficeLength.Inches(3.5), OfficeLength.Inches(5), OfficeLength.Inches(1.5), OfficeLength.Inches(1), "Thick");
            page.Shapes.Add(thickBorderRect);

            var noFillRect = new VisioShape("8", OfficeLength.Inches(5.5), OfficeLength.Inches(5), OfficeLength.Inches(1.5), OfficeLength.Inches(1), "No Fill");
            page.Shapes.Add(noFillRect);

            var gradientSimRect = new VisioShape("9", OfficeLength.Inches(7.5), OfficeLength.Inches(5), OfficeLength.Inches(1.5), OfficeLength.Inches(1), "Gradient");
            page.Shapes.Add(gradientSimRect);

            var grayRect = new VisioShape("10", OfficeLength.Inches(9.5), OfficeLength.Inches(5), OfficeLength.Inches(1.5), OfficeLength.Inches(1), "Gray");
            page.Shapes.Add(grayRect);

            // Row 3: Larger shapes with data
            var dataShape1 = new VisioShape("11", OfficeLength.Inches(2.5), OfficeLength.Inches(3), OfficeLength.Inches(2), OfficeLength.Inches(1.5), "Data Box 1");
            dataShape1.Data["Type"] = "Process";
            dataShape1.Data["Status"] = "Active";
            dataShape1.Data["Owner"] = "Team A";
            page.Shapes.Add(dataShape1);

            var dataShape2 = new VisioShape("12", OfficeLength.Inches(5.5), OfficeLength.Inches(3), OfficeLength.Inches(2), OfficeLength.Inches(1.5), "Data Box 2");
            dataShape2.Data["Type"] = "Database";
            dataShape2.Data["Status"] = "Online";
            dataShape2.Data["Size"] = "Large";
            page.Shapes.Add(dataShape2);

            var dataShape3 = new VisioShape("13", OfficeLength.Inches(8.5), OfficeLength.Inches(3), OfficeLength.Inches(2), OfficeLength.Inches(1.5), "Data Box 3");
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
            var shadowShape = new VisioShape("14", OfficeLength.Inches(2), OfficeLength.Inches(1), OfficeLength.Inches(1.8), OfficeLength.Inches(1.2), "Shadow");
            page.Shapes.Add(shadowShape);

            var roundedShape = new VisioShape("15", OfficeLength.Inches(4.5), OfficeLength.Inches(1), OfficeLength.Inches(1.8), OfficeLength.Inches(1.2), "Rounded");
            page.Shapes.Add(roundedShape);

            var doubleLineShape = new VisioShape("16", OfficeLength.Inches(7), OfficeLength.Inches(1), OfficeLength.Inches(1.8), OfficeLength.Inches(1.2), "Double");
            page.Shapes.Add(doubleLineShape);

            var transparentShape = new VisioShape("17", OfficeLength.Inches(9.5), OfficeLength.Inches(1), OfficeLength.Inches(1.8), OfficeLength.Inches(1.2), "50% Trans");
            page.Shapes.Add(transparentShape);

            document.Save();

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}