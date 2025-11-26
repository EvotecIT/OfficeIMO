using System;
using System.IO;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates table cell manipulation and row/column management.
    /// </summary>
    public static class TablesPowerPoint {
        public static void Example_PowerPointTables(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Table operations");
            string filePath = Path.Combine(folderPath, "Table Operations.pptx");

            using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
            // Slide 1: basic table
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTable table = slide.AddTable(3, 3);
            table.GetCell(0, 0).Text = "Product";
            table.GetCell(0, 1).Text = "Q1";
            table.GetCell(0, 2).Text = "Q2";
            table.GetCell(1, 0).Text = "Alpha";
            table.GetCell(1, 1).Text = "12";
            table.GetCell(1, 2).Text = "15";
            table.GetCell(2, 0).Text = "Beta";
            table.GetCell(2, 1).Text = "9";
            table.GetCell(2, 2).Text = "11";

            // Slide 2: clustered column chart
            PowerPointSlide slide2 = presentation.AddSlide();
            slide2.AddChart();

            // Slide 3: content slide
            PowerPointSlide slide3 = presentation.AddSlide();
            slide3.AddTextBox("Notes for tables and charts");

            // Slide 4: pie chart
            PowerPointSlide slide4 = presentation.AddSlide();
            slide4.AddChart();

            // Slide 5: picture with caption
            PowerPointSlide slide5 = presentation.AddSlide();
            string imagePath = Path.Combine(AppContext.BaseDirectory, "Images", "BackgroundImage.png");
            if (File.Exists(imagePath)) {
                slide5.AddPicture(imagePath, 914400L, 914400L, 5486400L, 3200400L);
            } else {
                slide5.AddTextBox("(image placeholder)");
            }

            // Slide 6: summary table
            PowerPointSlide slide6 = presentation.AddSlide();
            PowerPointTable summary = slide6.AddTable(2, 2);
            summary.GetCell(0, 0).Text = "Metric";
            summary.GetCell(0, 1).Text = "Value";
            summary.GetCell(1, 0).Text = "Total";
            summary.GetCell(1, 1).Text = "48";

            presentation.Save();

            Helpers.Open(filePath, openPowerPoint);
        }
    }
}
