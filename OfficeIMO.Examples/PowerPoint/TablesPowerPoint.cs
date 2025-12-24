using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates table cell manipulation and row/column management.
    /// </summary>
    public static class TablesPowerPoint {
        public static void Example_PowerPointTables(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Table operations");
            string filePath = Path.Combine(folderPath, "Table Operations.pptx");
            using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);

            const double marginCm = 1.5;
            const double gutterCm = 1.0;
            double slideHeightCm = presentation.SlideSize.HeightCm;
            PowerPointLayoutBox content = presentation.SlideSize.GetContentBoxCm(marginCm);
            double bodyTopCm = 3.8;
            double bodyHeightCm = slideHeightCm - bodyTopCm - marginCm;
            long bodyTop = PowerPointUnits.FromCentimeters(bodyTopCm);
            long bodyHeight = PowerPointUnits.FromCentimeters(bodyHeightCm);
            PowerPointLayoutBox[] columns = presentation.SlideSize.GetColumnsCm(2, marginCm, gutterCm);
            PowerPointLayoutBox leftColumn = new(columns[0].Left, bodyTop, columns[0].Width, bodyHeight);
            PowerPointLayoutBox rightColumn = new(columns[1].Left, bodyTop, columns[1].Width, bodyHeight);

            List<SalesRecord> data = new() {
                new SalesRecord { Product = "Alpha", Q1 = 12, Q2 = 15, Q3 = 18, Q4 = 20 },
                new SalesRecord { Product = "Beta", Q1 = 9, Q2 = 11, Q3 = 13, Q4 = 14 },
                new SalesRecord { Product = "Gamma", Q1 = 6, Q2 = 9, Q3 = 12, Q4 = 16 }
            };

            // Slide 1: title
            PowerPointSlide titleSlide = presentation.AddSlide();
            titleSlide.AddTitle("Quarterly Sales Overview");
            titleSlide.AddTextBoxCm("Tables and charts created from data objects", marginCm, 3.2, content.WidthCm, 1.1);

            // Slide 2: table from objects
            PowerPointSlide tableSlide = presentation.AddSlide();
            tableSlide.AddTitle("Sales by Product");
            PowerPointTable table = tableSlide.AddTableCm(
                data,
                o => {
                    o.HeaderCase = HeaderCase.Title;
                    o.PinFirst("Product");
                },
                includeHeaders: true,
                leftCm: marginCm,
                topCm: bodyTopCm,
                widthCm: content.WidthCm,
                heightCm: bodyHeightCm);
            table.BandedRows = true;
            table.SetColumnWidthsPoints(200, 90, 90, 90, 90);
            table.SetRowHeightPoints(0, 28);
            for (int r = 1; r < table.Rows; r++) {
                table.SetRowHeightPoints(r, 24);
            }

            for (int c = 0; c < table.Columns; c++) {
                PowerPointTableCell header = table.GetCell(0, c);
                header.FillColor = "1F4E79";
                header.Color = "FFFFFF";
                header.Bold = true;
                header.FontSize = 12;
                header.HorizontalAlignment = A.TextAlignmentTypeValues.Center;
                header.VerticalAlignment = A.TextAnchoringTypeValues.Center;
                header.SetBorders(TableCellBorders.All, "FFFFFF", 1);
                header.PaddingLeftPoints = 4;
                header.PaddingRightPoints = 4;
                header.PaddingTopPoints = 3;
                header.PaddingBottomPoints = 3;
            }

            for (int r = 1; r < table.Rows; r++) {
                for (int c = 0; c < table.Columns; c++) {
                    PowerPointTableCell cell = table.GetCell(r, c);
                    cell.FontSize = 11;
                    cell.SetBorders(TableCellBorders.All, "D9D9D9", 0.5);
                    cell.PaddingLeftPoints = 3;
                    cell.PaddingRightPoints = 3;
                }
            }

            // Slide 3: chart from the same dataset
            PowerPointSlide chartSlide = presentation.AddSlide();
            chartSlide.AddTitle("Quarterly Performance");
            PowerPointChartData chartData = new(
                data.Select(d => d.Product),
                new[] {
                    new PowerPointChartSeries("Q1", data.Select(d => (double)d.Q1)),
                    new PowerPointChartSeries("Q2", data.Select(d => (double)d.Q2)),
                    new PowerPointChartSeries("Q3", data.Select(d => (double)d.Q3)),
                    new PowerPointChartSeries("Q4", data.Select(d => (double)d.Q4))
                });
            chartSlide.AddChartCm(chartData, marginCm, bodyTopCm, content.WidthCm, bodyHeightCm);
            chartSlide.Notes.Text = "Chart and table share the same source data.";

            // Slide 4: table + chart side by side
            PowerPointSlide comboSlide = presentation.AddSlide();
            comboSlide.AddTitle("Table and Chart (Compact)");
            PowerPointTable compactTable = comboSlide.AddTableCm(
                data,
                o => {
                    o.HeaderCase = HeaderCase.Title;
                    o.PinFirst("Product");
                },
                includeHeaders: true,
                leftCm: leftColumn.LeftCm,
                topCm: leftColumn.TopCm,
                widthCm: leftColumn.WidthCm,
                heightCm: leftColumn.HeightCm);
            compactTable.BandedRows = true;
            compactTable.SetRowHeightPoints(0, 24);

            comboSlide.AddChartCm(chartData, rightColumn.LeftCm, rightColumn.TopCm, rightColumn.WidthCm,
                rightColumn.HeightCm);

            // Slide 5: totals by quarter
            PowerPointSlide totalsSlide = presentation.AddSlide();
            totalsSlide.AddTitle("Totals by Quarter");
            var quarterLabels = new[] { "Q1", "Q2", "Q3", "Q4" };
            var totals = new[] {
                data.Sum(d => d.Q1),
                data.Sum(d => d.Q2),
                data.Sum(d => d.Q3),
                data.Sum(d => d.Q4)
            };
            PowerPointChartData totalsData = new(
                quarterLabels,
                new[] { new PowerPointChartSeries("Total", totals.Select(t => (double)t)) });
            totalsSlide.AddChartCm(totalsData, marginCm, bodyTopCm, content.WidthCm, bodyHeightCm);

            // Slide 6: picture with caption
            PowerPointSlide imageSlide = presentation.AddSlide();
            imageSlide.AddTitle("Brand Assets");
            string imagePath = Path.Combine(AppContext.BaseDirectory, "Images", "BackgroundImage.png");
            if (File.Exists(imagePath)) {
                imageSlide.AddPictureCm(imagePath, marginCm, bodyTopCm, content.WidthCm, bodyHeightCm);
            } else {
                imageSlide.AddTextBoxCm("(image placeholder)", marginCm, bodyTopCm, content.WidthCm, 2.5);
            }

            // Slide 7: summary table
            PowerPointSlide summarySlide = presentation.AddSlide();
            summarySlide.AddTitle("Summary");
            var summaryRows = new[] {
                new SummaryRow { Metric = "Total Q1", Value = totals[0] },
                new SummaryRow { Metric = "Total Q2", Value = totals[1] },
                new SummaryRow { Metric = "Total Q3", Value = totals[2] },
                new SummaryRow { Metric = "Total Q4", Value = totals[3] }
            };
            PowerPointTable summary = summarySlide.AddTableCm(
                summaryRows,
                o => o.HeaderCase = HeaderCase.Title,
                includeHeaders: true,
                leftCm: marginCm,
                topCm: bodyTopCm,
                widthCm: content.WidthCm,
                heightCm: bodyHeightCm);
            summary.BandedRows = true;

            PowerPointTableCell summaryHeader = summary.GetCell(0, 0);
            summaryHeader.FillColor = "2F5597";
            summaryHeader.SetBorders(TableCellBorders.All, "FFFFFF", 1);
            summaryHeader.PaddingLeftPoints = 4;

            presentation.Save();

            Helpers.Open(filePath, openPowerPoint);
        }

        private sealed class SalesRecord {
            public string Product { get; set; } = string.Empty;
            public int Q1 { get; set; }
            public int Q2 { get; set; }
            public int Q3 { get; set; }
            public int Q4 { get; set; }
        }

        private sealed class SummaryRow {
            public string Metric { get; set; } = string.Empty;
            public int Value { get; set; }
        }
    }
}

