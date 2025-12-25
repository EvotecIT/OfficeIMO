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

            using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
            const double marginCm = 1.5;
            const double gutterCm = 1.0;
            PowerPointLayoutBox content = presentation.SlideSize.GetContentBoxCm(marginCm);
            double bodyTopCm = content.TopCm + 1.9;
            double bodyHeightCm = content.HeightCm - 1.9;
            PowerPointLayoutBox[] columns = presentation.SlideSize.GetColumnsCm(2, marginCm, gutterCm);
            PowerPointLayoutBox leftColumn = new(columns[0].Left, PowerPointUnits.FromCentimeters(bodyTopCm), columns[0].Width,
                PowerPointUnits.FromCentimeters(bodyHeightCm));
            PowerPointLayoutBox rightColumn = new(columns[1].Left, PowerPointUnits.FromCentimeters(bodyTopCm), columns[1].Width,
                PowerPointUnits.FromCentimeters(bodyHeightCm));

            // Slide 1: title
            PowerPointSlide cover = presentation.AddSlide();
            PowerPointTextBox title = cover.AddTitleCm("No-Repair PowerPoint",
                content.LeftCm, content.TopCm, content.WidthCm, 1.4);
            if (title.Paragraphs.Count > 0) {
                PowerPointTextStyle.Title.WithColor("1F4E79").Apply(title.Paragraphs[0]);
            }
            PowerPointTextBox subtitle = cover.AddTextBoxCm(
                "Built with OfficeIMO.PowerPoint — pure Open XML, no manual fixes.",
                content.LeftCm, content.TopCm + 1.9, content.WidthCm, 1.2);
            subtitle.ApplyTextStyle(PowerPointTextStyle.Body);
            subtitle.SetTextMarginsCm(0.2, 0.1, 0.2, 0.1);

            // Slide 2: features + image
            PowerPointSlide features = presentation.AddSlide();
            features.AddTitleCm("What you can build", content.LeftCm, content.TopCm, content.WidthCm, 1.4);
            PowerPointTextBox bullets = features.AddTextBoxCm(string.Empty,
                leftColumn.LeftCm, leftColumn.TopCm, leftColumn.WidthCm, leftColumn.HeightCm);
            bullets.SetTextMarginsCm(0.3, 0.2, 0.3, 0.2);
            bullets.AddBullets(new[] {
                "Clean PPTX output without repair dialogs",
                "Text boxes, shapes, images, tables, and charts",
                "Typed helpers for layouts and measurements",
                "Speaker notes and slide transitions"
            });
            bullets.ApplyAutoSpacing(lineSpacingMultiplier: 1.15, spaceAfterPoints: 2);

            if (File.Exists(imagePath)) {
                features.AddPicture(imagePath, rightColumn.Left, rightColumn.Top, rightColumn.Width, rightColumn.Height);
            } else {
                PowerPointTextBox placeholder = features.AddTextBoxCm("(image placeholder)",
                    rightColumn.LeftCm, rightColumn.TopCm, rightColumn.WidthCm, 2.0);
                placeholder.SetTextMarginsCm(0.2, 0.2, 0.2, 0.2);
            }

            // Slide 3: data snapshot
            PowerPointSlide dataSlide = presentation.AddSlide();
            dataSlide.AddTitleCm("Data snapshot", content.LeftCm, content.TopCm, content.WidthCm, 1.4);
            var rows = new[] {
                new MetricRow("Alpha", 12, 15),
                new MetricRow("Beta", 9, 11),
                new MetricRow("Gamma", 6, 9)
            };

            PowerPointTable table = dataSlide.AddTableCm(
                rows,
                columns: new[] {
                    PowerPointTableColumn<MetricRow>.Create("Product", r => r.Product).WithWidthCm(3.2),
                    PowerPointTableColumn<MetricRow>.Create("Q1", r => r.Q1),
                    PowerPointTableColumn<MetricRow>.Create("Q2", r => r.Q2)
                },
                includeHeaders: true,
                leftCm: leftColumn.LeftCm,
                topCm: leftColumn.TopCm,
                widthCm: leftColumn.WidthCm,
                heightCm: leftColumn.HeightCm);
            table.ApplyStyle(PowerPointTableStylePreset.Default);
            table.SetColumnWidthsEvenly();
            table.HeaderRow = true;
            table.BandedRows = true;

            PowerPointChart chart = dataSlide.AddChart(rows, r => r.Product,
                rightColumn.Left, rightColumn.Top, rightColumn.Width, rightColumn.Height,
                new PowerPointChartSeriesDefinition<MetricRow>("Q1", r => r.Q1),
                new PowerPointChartSeriesDefinition<MetricRow>("Q2", r => r.Q2));
            chart.SetTitle("Quarterly results");

            // Slide 4: notes
            PowerPointSlide notesSlide = presentation.AddSlide();
            notesSlide.AddTitleCm("Notes", content.LeftCm, content.TopCm, content.WidthCm, 1.4);
            PowerPointTextBox notesBody = notesSlide.AddTextBoxCm(
                "This slide contains speaker notes that are not visible in the main view.",
                content.LeftCm, content.TopCm + 1.9, content.WidthCm, 1.4);
            notesBody.ApplyTextStyle(PowerPointTextStyle.Body);
            notesBody.SetTextMarginsCm(0.2, 0.2, 0.2, 0.2);
            notesSlide.Notes.Text = "Speaker notes live in the notes section of each slide.";

            presentation.Save();

            Console.WriteLine($"[+] Presentation created successfully: {filePath}");
            Console.WriteLine("[+] The presentation should open without any repair dialog!");
        }

        private sealed record MetricRow(string Product, double Q1, double Q2);
    }
}
