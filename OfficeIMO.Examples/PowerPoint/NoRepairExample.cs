using System;
using System.IO;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Example demonstrating PowerPoint creation without repair issues.
    /// </summary>
    public static class NoRepairExample {
        public static void Example() {
            Console.WriteLine("[*] Creating PowerPoint presentation without repair issues");

            string filePath = Path.Combine(Path.GetTempPath(), "NoRepair_" + Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                // Add a title slide
                PowerPointSlide titleSlide = presentation.AddSlide();
                titleSlide.AddTitle("Welcome to OfficeIMO");
                titleSlide.AddTextBox("A reliable PowerPoint generation library");
                titleSlide.BackgroundColor = "F0F0F0";

                // Add a content slide with various elements
                PowerPointSlide contentSlide = presentation.AddSlide();
                contentSlide.AddTitle("Features");

                var textBox = contentSlide.AddTextBox("Key capabilities:");
                textBox.AddBullet("Create presentations programmatically");
                textBox.AddBullet("Add text, images, tables, and charts");
                textBox.AddBullet("Customize formatting and styles");
                textBox.AddBullet("No repair dialog issues!");
                textBox.ApplyAutoSpacing(lineSpacingMultiplier: 1.15, spaceAfterPoints: 2);

                // Add a picture if available
                if (File.Exists(imagePath)) {
                    contentSlide.AddPicture(imagePath, 4572000, 2286000, 3048000, 2286000);
                }

                // Add a table
                var table = contentSlide.AddTable(3, 3, 457200, 3657600, 4572000, 2286000);
                for (int row = 0; row < 3; row++) {
                    for (int col = 0; col < 3; col++) {
                        table.GetCell(row, col).Text = $"Cell {row+1},{col+1}";
                    }
                }

                // Add a chart slide
                PowerPointSlide chartSlide = presentation.AddSlide();
                chartSlide.AddTitle("Performance Metrics");
                chartSlide.AddChart();

                // Add slide with notes
                PowerPointSlide notesSlide = presentation.AddSlide();
                notesSlide.AddTitle("Additional Information");
                notesSlide.AddTextBox("This slide contains speaker notes");
                // notesSlide.Notes.Text = "These are speaker notes that won't be visible during the presentation.\n" +
                //                         "They can contain additional information for the presenter.";

                // Save the presentation
                presentation.Save();

                Console.WriteLine($"[+] Presentation created successfully: {filePath}");
                Console.WriteLine("[+] The presentation should open without any repair dialog!");
            }
        }
    }
}
