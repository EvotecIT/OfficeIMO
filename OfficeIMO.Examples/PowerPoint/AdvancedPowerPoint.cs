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
            double slideHeightCm = presentation.SlideSize.HeightCm;
            PowerPointLayoutBox content = presentation.SlideSize.GetContentBoxCm(marginCm);
            double bodyTopCm = 3.6;
            double bodyHeightCm = slideHeightCm - bodyTopCm - marginCm;
            long bodyTop = PowerPointUnits.FromCentimeters(bodyTopCm);
            long bodyHeight = PowerPointUnits.FromCentimeters(bodyHeightCm);
            PowerPointLayoutBox[] columns = presentation.SlideSize.GetColumnsCm(2, marginCm, gutterCm);
            PowerPointLayoutBox leftColumn = new(columns[0].Left, bodyTop, columns[0].Width, bodyHeight);
            PowerPointLayoutBox rightColumn = new(columns[1].Left, bodyTop, columns[1].Width, bodyHeight);

            List<KpiRow> kpis = new() {
                new KpiRow { Metric = "Revenue", Current = 2.4, Target = 2.8 },
                new KpiRow { Metric = "Retention", Current = 91.2, Target = 92.5 },
                new KpiRow { Metric = "Net Promoter Score", Current = 38, Target = 42 }
            };

            // Slide 1: title
            PowerPointSlide slide1 = presentation.AddSlide();
            PowerPointTextBox title = slide1.AddTitleCm("Advanced PowerPoint Example", marginCm, marginCm, content.WidthCm, titleHeightCm);
            title.FontSize = 30;
            title.Color = "1F4E79";
            slide1.AddTextBoxCm("Shows layout, charts, tables, images, notes, and transitions.", marginCm, 3.2, content.WidthCm, 1.1);
            slide1.BackgroundColor = "F7F7F7";
            slide1.Transition = SlideTransition.Fade;

            // Slide 2: highlights with image
            PowerPointSlide slide2 = presentation.AddSlide();
            slide2.AddTitleCm("Highlights", marginCm, marginCm, content.WidthCm, titleHeightCm);
            slide2.Transition = SlideTransition.PushLeft;
            PowerPointTextBox highlights = slide2.AddTextBox("This slide demonstrates:",
                leftColumn.Left, leftColumn.Top, leftColumn.Width, leftColumn.Height);
            highlights.AddBullet("Positioned elements");
            highlights.AddBullet("Tables created from objects");
            highlights.AddBullet("Charts with real data");
            highlights.AddBullet("Notes and transitions");
            highlights.ApplyAutoSpacing(lineSpacingMultiplier: 1.15, spaceAfterPoints: 2);

            string imagePath = Path.Combine(AppContext.BaseDirectory, "Images", "BackgroundImage.png");
            if (File.Exists(imagePath)) {
                slide2.AddPicture(imagePath, rightColumn.Left, rightColumn.Top, rightColumn.Width, rightColumn.Height);
            }

            // Slide 3: KPI table
            PowerPointSlide slide3 = presentation.AddSlide();
            slide3.AddTitleCm("KPI Table", marginCm, marginCm, content.WidthCm, titleHeightCm);
            slide3.Transition = SlideTransition.Wipe;
            double tableTopCm = bodyTopCm;
            double tableHeightCm = slideHeightCm - tableTopCm - marginCm;
            PowerPointTable kpiTable = slide3.AddTableCm(
                kpis,
                o => {
                    o.HeaderCase = HeaderCase.Title;
                    o.PinFirst("Metric");
                },
                includeHeaders: true,
                leftCm: marginCm,
                topCm: tableTopCm,
                widthCm: content.WidthCm,
                heightCm: tableHeightCm);
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
            slide4.AddTitleCm("KPI Chart", marginCm, marginCm, content.WidthCm, titleHeightCm);
            slide4.Transition = SlideTransition.CombHorizontal;
            PowerPointChartData chartData = PowerPointChartData.From(
                kpis,
                k => k.Metric,
                new PowerPointChartSeriesDefinition<KpiRow>("Current", k => k.Current),
                new PowerPointChartSeriesDefinition<KpiRow>("Target", k => k.Target));
            double chartTopCm = bodyTopCm;
            double chartHeightCm = slideHeightCm - chartTopCm - marginCm;
            PowerPointChart chart = slide4.AddChartCm(chartData, marginCm, chartTopCm, content.WidthCm, chartHeightCm);
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
