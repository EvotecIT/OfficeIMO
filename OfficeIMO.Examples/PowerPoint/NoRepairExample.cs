using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
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

            using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
            const double marginCm = 1.5;
            const double gutterCm = 1.0;
            PowerPointLayoutBox content = presentation.SlideSize.GetContentBox(OfficeIMO.OfficeLength.Centimeters(marginCm));
            double bodyTopCm = content.Top.ToCentimeters() + 1.9;
            double bodyHeightCm = content.Height.ToCentimeters() - 1.9;
            PowerPointLayoutBox[] columns = presentation.SlideSize.GetColumns(2,OfficeIMO.OfficeLength.Centimeters(marginCm),OfficeIMO.OfficeLength.Centimeters(gutterCm));
            PowerPointLayoutBox leftColumn = new(OfficeIMO.OfficeLength.Emus(columns[0].Left.ToEmus()),OfficeIMO.OfficeLength.Emus(PowerPointUnits.FromCentimeters(bodyTopCm)),OfficeIMO.OfficeLength.Emus(columns[0].Width.ToEmus()),                OfficeIMO.OfficeLength.Emus(PowerPointUnits.FromCentimeters(bodyHeightCm)));
            PowerPointLayoutBox rightColumn = new(OfficeIMO.OfficeLength.Emus(columns[1].Left.ToEmus()),OfficeIMO.OfficeLength.Emus(PowerPointUnits.FromCentimeters(bodyTopCm)),OfficeIMO.OfficeLength.Emus(columns[1].Width.ToEmus()),                OfficeIMO.OfficeLength.Emus(PowerPointUnits.FromCentimeters(bodyHeightCm)));

            // Slide 1: title
            PowerPointSlide cover = presentation.AddSlide();
            PowerPointTextBox title = cover.AddTitleCm("No-Repair PowerPoint",
                content.Left.ToCentimeters(), content.Top.ToCentimeters(), content.Width.ToCentimeters(), 1.4);
            if (title.Paragraphs.Count > 0) {
                PowerPointTextStyle.Title.WithColor("1F4E79").Apply(title.Paragraphs[0]);
            }
            PowerPointTextBox subtitle = cover.AddTextBoxCm(
                "Built with OfficeIMO.PowerPoint — pure Open XML, no manual fixes.",
                content.Left.ToCentimeters(), content.Top.ToCentimeters() + 1.9, content.Width.ToCentimeters(), 1.2);
            subtitle.ApplyTextStyle(PowerPointTextStyle.Body);
            subtitle.SetTextMarginsCm(0.2, 0.1, 0.2, 0.1);

            // Slide 2: features + image
            PowerPointSlide features = presentation.AddSlide();
            features.AddTitleCm("What you can build", content.Left.ToCentimeters(), content.Top.ToCentimeters(), content.Width.ToCentimeters(), 1.4);
            PowerPointTextBox bullets = features.AddTextBoxCm(string.Empty,
                leftColumn.Left.ToCentimeters(), leftColumn.Top.ToCentimeters(), leftColumn.Width.ToCentimeters(), leftColumn.Height.ToCentimeters());
            bullets.SetTextMarginsCm(0.3, 0.2, 0.3, 0.2);
            bullets.AddBullets(new[] {
                "Clean PPTX output without repair dialogs",
                "Text boxes, shapes, images, tables, and charts",
                "Typed helpers for layouts and measurements",
                "Speaker notes and slide transitions"
            });
            bullets.ApplyAutoSpacing(lineSpacingMultiplier: 1.15, spaceAfterPoints: 2);

            if (File.Exists(imagePath)) {
                features.AddImage(imagePath, rightColumn.Left.ToEmus(), rightColumn.Top.ToEmus(), rightColumn.Width.ToEmus(), rightColumn.Height.ToEmus());
            } else {
                PowerPointTextBox placeholder = features.AddTextBoxCm("(image placeholder)",
                    rightColumn.Left.ToCentimeters(), rightColumn.Top.ToCentimeters(), rightColumn.Width.ToCentimeters(), 2.0);
                placeholder.SetTextMarginsCm(0.2, 0.2, 0.2, 0.2);
            }

            // Slide 3: data snapshot
            PowerPointSlide dataSlide = presentation.AddSlide();
            dataSlide.AddTitleCm("Data snapshot", content.Left.ToCentimeters(), content.Top.ToCentimeters(), content.Width.ToCentimeters(), 1.4);
            var rows = new[] {
                new MetricRow("Alpha", 12, 15),
                new MetricRow("Beta", 9, 11),
                new MetricRow("Gamma", 6, 9)
            };

            PowerPointTable table = dataSlide.AddTableCm(
                rows,
                columns: new[] {
                    PowerPointTableColumn<MetricRow>.Create("Product", r => r.Product).WithWidth(global::OfficeIMO.OfficeLength.Centimeters(3.2)),
                    PowerPointTableColumn<MetricRow>.Create("Q1", r => r.Q1),
                    PowerPointTableColumn<MetricRow>.Create("Q2", r => r.Q2)
                },
                includeHeaders: true,
                leftCm: leftColumn.Left.ToCentimeters(),
                topCm: leftColumn.Top.ToCentimeters(),
                widthCm: leftColumn.Width.ToCentimeters(),
                heightCm: leftColumn.Height.ToCentimeters());
            table.ApplyStyle(PowerPointTableStylePreset.Default);
            table.SetColumnWidthsEvenly();
            table.HeaderRow = true;
            table.BandedRows = true;

            OfficeChartData chartData = new OfficeChartData(rows.Select(r => r.Product), new[] {
                new OfficeSeries("Q1", rows.Select(r => (double)r.Q1)),
                new OfficeSeries("Q2", rows.Select(r => (double)r.Q2))
            });
            PowerPointChart chart = dataSlide.AddChart(OfficeChartType.ColumnClustered, chartData,
                rightColumn.Left.ToEmus(), rightColumn.Top.ToEmus(), rightColumn.Width.ToEmus(), rightColumn.Height.ToEmus());
            chart.SetTitle("Quarterly results");

            // Slide 4: notes
            PowerPointSlide notesSlide = presentation.AddSlide();
            notesSlide.AddTitleCm("Notes", content.Left.ToCentimeters(), content.Top.ToCentimeters(), content.Width.ToCentimeters(), 1.4);
            PowerPointTextBox notesBody = notesSlide.AddTextBoxCm(
                "This slide contains speaker notes that are not visible in the main view.",
                content.Left.ToCentimeters(), content.Top.ToCentimeters() + 1.9, content.Width.ToCentimeters(), 1.4);
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
