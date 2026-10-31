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
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "Kulek.jpg");

            using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
            const double marginCm = 1.5;
            const double gutterCm = 1.0;
            PowerPointLayoutBox content = presentation.SlideSize.GetContentBox(OfficeIMO.OfficeLength.Centimeters(marginCm));
            double bodyTopCm = content.Top.ToCentimeters() + 1.9;
            double bodyHeightCm = content.Height.ToCentimeters() - 1.9;
            PowerPointLayoutBox[] columns = presentation.SlideSize.GetColumns(2,OfficeIMO.OfficeLength.Centimeters(marginCm),OfficeIMO.OfficeLength.Centimeters(gutterCm));
            PowerPointLayoutBox leftColumn = new(OfficeIMO.OfficeLength.Emus(columns[0].Left.ToEmus()),global::OfficeIMO.OfficeLength.Centimeters(bodyTopCm),OfficeIMO.OfficeLength.Emus(columns[0].Width.ToEmus()),                global::OfficeIMO.OfficeLength.Centimeters(bodyHeightCm));
            PowerPointLayoutBox rightColumn = new(OfficeIMO.OfficeLength.Emus(columns[1].Left.ToEmus()),global::OfficeIMO.OfficeLength.Centimeters(bodyTopCm),OfficeIMO.OfficeLength.Emus(columns[1].Width.ToEmus()),                global::OfficeIMO.OfficeLength.Centimeters(bodyHeightCm));

            // Slide 1: title
            PowerPointSlide cover = presentation.AddSlide();
            PowerPointTextBox title = cover.AddTitle("No-Repair PowerPoint",                global::OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(1.4));
            if (title.Paragraphs.Count > 0) {
                PowerPointTextStyle.Title.WithColor("1F4E79").Apply(title.Paragraphs[0]);
            }
            PowerPointTextBox subtitle = cover.AddTextBox(
                "Built with OfficeIMO.PowerPoint — pure Open XML, no manual fixes.",                global::OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters() + 1.9),global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(1.2));
            subtitle.ApplyTextStyle(PowerPointTextStyle.Body);
            subtitle.SetTextMargins(global::OfficeIMO.OfficeLength.Centimeters(0.2), global::OfficeIMO.OfficeLength.Centimeters(0.1), global::OfficeIMO.OfficeLength.Centimeters(0.2), global::OfficeIMO.OfficeLength.Centimeters(0.1));

            // Slide 2: features + image
            PowerPointSlide features = presentation.AddSlide();
            features.AddTitle("What you can build",global::OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(1.4));
            PowerPointTextBox bullets = features.AddTextBox(string.Empty,                global::OfficeIMO.OfficeLength.Centimeters(leftColumn.Left.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(leftColumn.Top.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(leftColumn.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(leftColumn.Height.ToCentimeters()));
            bullets.SetTextMargins(global::OfficeIMO.OfficeLength.Centimeters(0.3), global::OfficeIMO.OfficeLength.Centimeters(0.2), global::OfficeIMO.OfficeLength.Centimeters(0.3), global::OfficeIMO.OfficeLength.Centimeters(0.2));
            bullets.AddBullets(new[] {
                "Clean PPTX output without repair dialogs",
                "Text boxes, shapes, images, tables, and charts",
                "Typed helpers for layouts and measurements",
                "Speaker notes and slide transitions"
            });
            bullets.ApplyAutoSpacing(lineSpacingMultiplier: 1.15, spaceAfter: global::OfficeIMO.OfficeLength.Points(2));

            if (File.Exists(imagePath)) {
                features.AddImage(imagePath, rightColumn.Left, rightColumn.Top, rightColumn.Width, rightColumn.Height);
            } else {
                PowerPointTextBox placeholder = features.AddTextBox("(image placeholder)",                    global::OfficeIMO.OfficeLength.Centimeters(rightColumn.Left.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(rightColumn.Top.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(rightColumn.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(2.0));
                placeholder.SetTextMargins(global::OfficeIMO.OfficeLength.Centimeters(0.2), global::OfficeIMO.OfficeLength.Centimeters(0.2), global::OfficeIMO.OfficeLength.Centimeters(0.2), global::OfficeIMO.OfficeLength.Centimeters(0.2));
            }

            // Slide 3: data snapshot
            PowerPointSlide dataSlide = presentation.AddSlide();
            dataSlide.AddTitle("Data snapshot",global::OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(1.4));
            var rows = new[] {
                new MetricRow("Alpha", 12, 15),
                new MetricRow("Beta", 9, 11),
                new MetricRow("Gamma", 6, 9)
            };

            PowerPointTable table = dataSlide.AddTable(
                rows,                columns: new[] {
                    PowerPointTableColumn<MetricRow>.Create("Product", r => r.Product).WithWidth(global::OfficeIMO.OfficeLength.Centimeters(3.2)),
                    PowerPointTableColumn<MetricRow>.Create("Q1", r => r.Q1),
                    PowerPointTableColumn<MetricRow>.Create("Q2", r => r.Q2)
                },                includeHeaders: true,left: global::OfficeIMO.OfficeLength.Centimeters(leftColumn.Left.ToCentimeters()),top: global::OfficeIMO.OfficeLength.Centimeters(leftColumn.Top.ToCentimeters()),width: global::OfficeIMO.OfficeLength.Centimeters(leftColumn.Width.ToCentimeters()),height: global::OfficeIMO.OfficeLength.Centimeters(leftColumn.Height.ToCentimeters()));
            table.ApplyStyle(PowerPointTableStylePreset.Default);
            table.SetColumnWidthsEvenly();
            table.HeaderRow = true;
            table.BandedRows = true;

            OfficeChartData chartData = new OfficeChartData(rows.Select(r => r.Product), new[] {
                new OfficeSeries("Q1", rows.Select(r => (double)r.Q1)),
                new OfficeSeries("Q2", rows.Select(r => (double)r.Q2))
            });
            PowerPointChart chart = dataSlide.AddChart(OfficeChartType.ColumnClustered, chartData, new OfficeIMO.PowerPoint.PowerPointChartOptions { Left = rightColumn.Left, Top = rightColumn.Top, Width = rightColumn.Width, Height = rightColumn.Height });
            chart.SetTitle("Quarterly results");

            // Slide 4: notes
            PowerPointSlide notesSlide = presentation.AddSlide();
            notesSlide.AddTitle("Notes",global::OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(1.4));
            PowerPointTextBox notesBody = notesSlide.AddTextBox(
                "This slide contains speaker notes that are not visible in the main view.",                global::OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters() + 1.9),global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(1.4));
            notesBody.ApplyTextStyle(PowerPointTextStyle.Body);
            notesBody.SetTextMargins(global::OfficeIMO.OfficeLength.Centimeters(0.2), global::OfficeIMO.OfficeLength.Centimeters(0.2), global::OfficeIMO.OfficeLength.Centimeters(0.2), global::OfficeIMO.OfficeLength.Centimeters(0.2));
            notesSlide.Notes.Text = "Speaker notes live in the notes section of each slide.";

            presentation.Save();

            Console.WriteLine($"[+] Presentation created successfully: {filePath}");
            Console.WriteLine("[+] The presentation should open without any repair dialog!");
        }

        private sealed record MetricRow(string Product, double Q1, double Q2);
    }
}
