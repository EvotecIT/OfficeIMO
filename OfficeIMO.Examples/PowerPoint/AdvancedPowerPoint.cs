using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
            using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
            const double marginCm = 1.5;
            const double gutterCm = 1.0;
            const double titleHeightCm = 1.4;
            const double bodyGapCm = 0.8;
            PowerPointLayoutBox content = presentation.SlideSize.GetContentBoxCm(marginCm);
            PowerPointLayoutBox titleBox = PowerPointLayoutBox.FromCentimeters(
                content.LeftCm, content.TopCm, content.WidthCm, titleHeightCm);
            PowerPointLayoutBox bodyBox = PowerPointLayoutBox.FromCentimeters(
                content.LeftCm,
                content.TopCm + titleHeightCm + bodyGapCm,
                content.WidthCm,
                content.HeightCm - titleHeightCm - bodyGapCm);
            PowerPointLayoutBox[] columns = bodyBox.SplitColumnsCm(2, gutterCm);
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
            PowerPointLayoutBox introBox = PowerPointLayoutBox.FromCentimeters(
                bodyBox.LeftCm, bodyBox.TopCm, bodyBox.WidthCm, 1.1);
            PowerPointTextBox intro = slide1.AddTextBox(
                "Shows layout, charts, tables, images, notes, and transitions.",
                introBox);
            intro.FontSize = 18;
            intro.Color = "1F4E79";
            slide1.BackgroundColor = "F7F7F7";
            slide1.Transition = SlideTransition.Fade;

            // Slide 2: highlights with image
            PowerPointSlide slide2 = presentation.AddSlide();
            slide2.AddTitle("Highlights", titleBox);
            slide2.Transition = SlideTransition.PushLeft;
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
            double mediaHeightCm = rightColumn.HeightCm - captionHeightCm - captionGapCm;
            PowerPointLayoutBox mediaBox = PowerPointLayoutBox.FromCentimeters(
                rightColumn.LeftCm, rightColumn.TopCm, rightColumn.WidthCm, mediaHeightCm);
            PowerPointLayoutBox captionBox = PowerPointLayoutBox.FromCentimeters(
                rightColumn.LeftCm,
                rightColumn.TopCm + mediaHeightCm + captionGapCm,
                rightColumn.WidthCm,
                captionHeightCm);

            string imagePath = Path.Combine(AppContext.BaseDirectory, "Images", "BackgroundImage.png");
            if (File.Exists(imagePath)) {
                slide2.AddPicture(imagePath, mediaBox);
                PowerPointTextBox caption = slide2.AddTextBox("Illustration scaled to fit.", captionBox);
                caption.FontSize = 12;
                caption.Color = "666666";
            } else {
                PowerPointTextBox placeholder = slide2.AddTextBox("(image placeholder)", mediaBox);
                placeholder.FillColor = "F3F3F3";
                placeholder.OutlineColor = "CCCCCC";
                placeholder.TextVerticalAlignment = A.TextAnchoringTypeValues.Center;
                placeholder.ApplyTextStyle(PowerPointTextStyle.Body.WithColor("777777"));
            }

            // Slide 3: KPI table
            PowerPointSlide slide3 = presentation.AddSlide();
            slide3.AddTitle("KPI Table", titleBox);
            slide3.Transition = SlideTransition.Wipe;
            PowerPointTable kpiTable = slide3.AddTableCm(
                kpis,
                o => {
                    o.HeaderCase = HeaderCase.Title;
                    o.PinFirst("Metric");
                },
                includeHeaders: true,
                leftCm: bodyBox.LeftCm,
                topCm: bodyBox.TopCm,
                widthCm: bodyBox.WidthCm,
                heightCm: bodyBox.HeightCm);
            kpiTable.ApplyStyle(PowerPointTableStylePreset.Default);
            kpiTable.SetColumnWidthsPoints(220, 120, 120);
            kpiTable.SetRowHeightPoints(0, 28);
            for (int c = 0; c < kpiTable.Columns; c++) {
                PowerPointTableCell header = kpiTable.GetCell(0, c);
                header.Bold = true;
                header.HorizontalAlignment = A.TextAlignmentTypeValues.Center;
                header.VerticalAlignment = A.TextAnchoringTypeValues.Center;
            }

            // Slide 4: KPI chart
            PowerPointSlide slide4 = presentation.AddSlide();
            slide4.AddTitle("KPI Chart", titleBox);
            slide4.Transition = SlideTransition.CombHorizontal;
            PowerPointChartData chartData = PowerPointChartData.From(
                kpis,
                k => k.Metric,
                new PowerPointChartSeriesDefinition<KpiRow>("Current", k => k.Current),
                new PowerPointChartSeriesDefinition<KpiRow>("Target", k => k.Target));
            PowerPointChart chart = slide4.AddChartCm(
                chartData,
                bodyBox.LeftCm,
                bodyBox.TopCm,
                bodyBox.WidthCm,
                bodyBox.HeightCm);
            chart.SetTitle("Current vs Target")
                .SetLegend(C.LegendPositionValues.Right)
                .SetDataLabels(showValue: true)
                .SetCategoryAxisTitle("Metric")
                .SetValueAxisTitle("Score");
            slide4.Notes.Text = "Targets are shown alongside current values.";

            presentation.Save();

            Helpers.Open(filePath, openPowerPoint);
        }

        private sealed class KpiRow {
            public string Metric { get; set; } = string.Empty;
            public double Current { get; set; }
            public double Target { get; set; }
        }
    }
}
