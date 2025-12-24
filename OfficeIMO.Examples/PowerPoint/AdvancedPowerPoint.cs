using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates advanced slide features such as backgrounds, transitions, charts, and tables.
    /// </summary>
    public static class AdvancedPowerPoint {
        public static void Example_AdvancedPowerPoint(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Advanced features");
            string filePath = Path.Combine(folderPath, "Advanced PowerPoint.pptx");
            using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);

            const long slideWidth = 12192000L;
            const long margin = 914400L;
            const long contentWidth = slideWidth - (2 * margin);

            List<KpiRow> kpis = new() {
                new KpiRow { Metric = "Revenue", Current = 2.4, Target = 2.8 },
                new KpiRow { Metric = "Retention", Current = 91.2, Target = 92.5 },
                new KpiRow { Metric = "Net Promoter Score", Current = 38, Target = 42 }
            };

            // Slide 1: title
            PowerPointSlide slide1 = presentation.AddSlide();
            slide1.AddTitle("Advanced PowerPoint Example");
            slide1.AddTextBox("Shows layout, charts, tables, images, and notes", margin, 2011680L, contentWidth, 914400L);
            slide1.BackgroundColor = "F7F7F7";
            slide1.Transition = SlideTransition.Fade;

            // Slide 2: highlights with image
            PowerPointSlide slide2 = presentation.AddSlide();
            slide2.AddTitle("Highlights");
            PowerPointTextBox highlights = slide2.AddTextBox("This slide demonstrates:");
            highlights.AddBullet("Positioned elements");
            highlights.AddBullet("Tables created from objects");
            highlights.AddBullet("Charts with real data");
            highlights.AddBullet("Notes and transitions");

            string imagePath = Path.Combine(AppContext.BaseDirectory, "Images", "BackgroundImage.png");
            if (File.Exists(imagePath)) {
                slide2.AddPicture(imagePath, 6400000L, 1700000L, 4572000L, 2570400L);
            }

            // Slide 3: KPI table
            PowerPointSlide slide3 = presentation.AddSlide();
            slide3.AddTitle("KPI Table");
            PowerPointTable kpiTable = slide3.AddTable(
                kpis,
                o => {
                    o.HeaderCase = HeaderCase.Title;
                    o.PinFirst("Metric");
                },
                includeHeaders: true,
                left: margin,
                top: 1828800L,
                width: contentWidth,
                height: 2286000L);
            kpiTable.BandedRows = true;

            // Slide 4: KPI chart
            PowerPointSlide slide4 = presentation.AddSlide();
            slide4.AddTitle("KPI Chart");
            PowerPointChartData chartData = new(
                kpis.Select(k => k.Metric),
                new[] {
                    new PowerPointChartSeries("Current", kpis.Select(k => k.Current)),
                    new PowerPointChartSeries("Target", kpis.Select(k => k.Target))
                });
            slide4.AddChart(chartData, margin, 1524000L, contentWidth, 3810000L);
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

