using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Data;
using OfficeIMO.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates advanced slide features such as backgrounds, transitions, charts, and tables.
    /// </summary>
    public static class AdvancedPowerPoint {
        public static void Example_AdvancedPowerPoint(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Advanced features");
            string filePath = Path.Combine(folderPath, "Advanced PowerPoint.pptx");
            using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
            const double marginCm = 1.5;
            const double gutterCm = 1.0;
            const double titleHeightCm = 1.4;
            const double bodyGapCm = 0.8;
            PowerPointLayoutBox content = presentation.SlideSize.GetContentBox(OfficeIMO.OfficeLength.Centimeters(marginCm));
            PowerPointLayoutBox titleBox = new OfficeIMO.PowerPoint.PowerPointLayoutBox(
                OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(titleHeightCm));
            PowerPointLayoutBox bodyBox = new OfficeIMO.PowerPoint.PowerPointLayoutBox(
                OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()),                OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters() + titleHeightCm + bodyGapCm),                OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),                OfficeIMO.OfficeLength.Centimeters(content.Height.ToCentimeters() - titleHeightCm - bodyGapCm));
            PowerPointLayoutBox[] columns = bodyBox.SplitColumns(2,OfficeIMO.OfficeLength.Centimeters(gutterCm));
            PowerPointLayoutBox leftColumn = columns[0];
            PowerPointLayoutBox rightColumn = columns[1];

            List<KpiRow> kpis = new() {
                new KpiRow { Metric = "Revenue", Current = 2.4, Target = 2.8 },
                new KpiRow { Metric = "Retention", Current = 91.2, Target = 92.5 },
                new KpiRow { Metric = "Net Promoter Score", Current = 38, Target = 42 }
            };

            // Slide 1: title
            PowerPointSlide slide1 = presentation.AddSlide();
            PowerPointTextBox title = slide1.AddTitle("Advanced PowerPoint Example", titleBox);
            title.FontSize = 30;
            title.Color = "1F4E79";
            PowerPointLayoutBox introBox = new OfficeIMO.PowerPoint.PowerPointLayoutBox(
                OfficeIMO.OfficeLength.Centimeters(bodyBox.Left.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(bodyBox.Top.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(bodyBox.Width.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(1.1));
            PowerPointTextBox intro = slide1.AddTextBox(
                "Shows layout, charts, tables, images, notes, and transitions.",
                introBox);
            intro.FontSize = 18;
            intro.Color = "1F4E79";
            slide1.BackgroundColor = "F7F7F7";
            slide1.Transition = PowerPointSlideTransition.Fade;

            // Slide 2: highlights with image
            PowerPointSlide slide2 = presentation.AddSlide();
            slide2.AddTitle("Highlights", titleBox);
            slide2.Transition = PowerPointSlideTransition.PushLeft;
            PowerPointTextBox highlights = slide2.AddTextBox("This slide demonstrates:", leftColumn);
            highlights.TextAutoFit = PowerPointTextAutoFit.Normal;
            highlights.SetTextMarginsCm(0.3, 0.2, 0.3, 0.2);
            highlights.AddBullet("Positioned elements");
            highlights.AddBullet("Tables created from objects");
            highlights.AddBullet("Charts with real data");
            highlights.AddBullet("Notes and transitions");
            highlights.ApplyAutoSpacing(lineSpacingMultiplier: 1.15, spaceAfterPoints: 2);

            const double captionHeightCm = 1.0;
            const double captionGapCm = 0.4;
            double mediaHeightCm = rightColumn.Height.ToCentimeters() - captionHeightCm - captionGapCm;
            PowerPointLayoutBox mediaBox = new OfficeIMO.PowerPoint.PowerPointLayoutBox(
                OfficeIMO.OfficeLength.Centimeters(rightColumn.Left.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(rightColumn.Top.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(rightColumn.Width.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(mediaHeightCm));
            PowerPointLayoutBox captionBox = new OfficeIMO.PowerPoint.PowerPointLayoutBox(
                OfficeIMO.OfficeLength.Centimeters(rightColumn.Left.ToCentimeters()),                OfficeIMO.OfficeLength.Centimeters(rightColumn.Top.ToCentimeters() + mediaHeightCm + captionGapCm),                OfficeIMO.OfficeLength.Centimeters(rightColumn.Width.ToCentimeters()),                OfficeIMO.OfficeLength.Centimeters(captionHeightCm));

            string imagePath = Path.Combine(AppContext.BaseDirectory, "Images", "BackgroundImage.png");
            if (File.Exists(imagePath)) {
                slide2.AddImage(imagePath, mediaBox);
                PowerPointTextBox caption = slide2.AddTextBox("Illustration scaled to fit.", captionBox);
                caption.FontSize = 12;
                caption.Color = "666666";
            } else {
                PowerPointTextBox placeholder = slide2.AddTextBox("(image placeholder)", mediaBox);
                placeholder.FillColor = "F3F3F3";
                placeholder.OutlineColor = "CCCCCC";
                placeholder.TextVerticalAlignment = PowerPointTextVerticalAlignment.Center;
                placeholder.ApplyTextStyle(PowerPointTextStyle.Body.WithColor("777777"));
            }

            // Slide 3: KPI table
            PowerPointSlide slide3 = presentation.AddSlide();
            slide3.AddTitle("KPI Table", titleBox);
            slide3.Transition = PowerPointSlideTransition.Wipe;
            PowerPointTable kpiTable = slide3.AddTable(
                kpis,                o => {
                    o.HeaderCase = HeaderCase.Title;
                    o.PinFirst("Metric");
                },                includeHeaders: true,left: global::OfficeIMO.OfficeLength.Centimeters(bodyBox.Left.ToCentimeters()),top: global::OfficeIMO.OfficeLength.Centimeters(bodyBox.Top.ToCentimeters()),width: global::OfficeIMO.OfficeLength.Centimeters(bodyBox.Width.ToCentimeters()),height: global::OfficeIMO.OfficeLength.Centimeters(bodyBox.Height.ToCentimeters()));
            kpiTable.ApplyStyle(PowerPointTableStylePreset.Default);
            kpiTable.SetColumnWidths(global::OfficeIMO.OfficeLength.Points(220),global::OfficeIMO.OfficeLength.Points(120),global::OfficeIMO.OfficeLength.Points(120));
            kpiTable.SetRowHeight(0,global::OfficeIMO.OfficeLength.Points(28));
            for (int c = 0; c < kpiTable.Columns; c++) {
                PowerPointTableCell header = kpiTable.GetCell(0, c);
                header.Bold = true;
                header.HorizontalAlignment = PowerPointTextAlignment.Center;
                header.VerticalAlignment = PowerPointTextVerticalAlignment.Center;
            }

            // Slide 4: KPI chart
            PowerPointSlide slide4 = presentation.AddSlide();
            slide4.AddTitle("KPI Chart", titleBox);
            slide4.Transition = PowerPointSlideTransition.CombHorizontal;
            OfficeChartData chartData = new OfficeChartData(
                kpis.Select(k => k.Metric),
                new[] {
                    new OfficeSeries("Current", kpis.Select(k => (double)k.Current)),
                    new OfficeSeries("Target", kpis.Select(k => (double)k.Target))
                });
            PowerPointChart chart = slide4.AddChartCm(
                OfficeChartType.ColumnClustered, chartData,
                bodyBox.Left.ToCentimeters(),
                bodyBox.Top.ToCentimeters(),
                bodyBox.Width.ToCentimeters(),
                bodyBox.Height.ToCentimeters());
            chart.SetTitle("Current vs Target")
                .SetLegend(OfficeChartLegendPosition.Right)
                .SetDataLabels(showValue: true)
                .SetCategoryAxisTitle("Metric")
                .SetValueAxisTitle("Score");
            slide4.Notes.Text = "Targets are shown alongside current values.";

            presentation.Save();

            if (openPowerPoint) ExampleFileLauncher.Open(filePath);
        }

        private sealed class KpiRow {
            public string Metric { get; set; } = string.Empty;
            public double Current { get; set; }
            public double Target { get; set; }
        }
    }
}
