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
            const double titleHeightCm = 1.4;
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
            IReadOnlyList<PowerPointTableStyleInfo> tableStyles = presentation.TableStyles;
            PowerPointTableStyleInfo? accentStyle = FindStyle(tableStyles, "Medium Style 2 - Accent 1");
            PowerPointTableStyleInfo? darkStyle = FindStyle(tableStyles, "Dark Style 1 - Accent 2");
            PowerPointTableStyleInfo? lightStyle = FindStyle(tableStyles, "Light Style 1");

            // Slide 1: title
            PowerPointSlide titleSlide = presentation.AddSlide();
            PowerPointTextBox title = titleSlide.AddTitleCm("Quarterly Sales Overview", marginCm, marginCm, content.WidthCm, titleHeightCm);
            title.FontSize = 32;
            title.Color = "1F4E79";
            titleSlide.AddTextBoxCm("Tables created from data objects with built-in styles.", marginCm, 3.2, content.WidthCm, 1.1);

            // Slide 2: table from objects
            PowerPointSlide tableSlide = presentation.AddSlide();
            tableSlide.AddTitleCm("Sales by Product", marginCm, marginCm, content.WidthCm, titleHeightCm);
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
            table.ApplyStyle(new PowerPointTableStylePreset(styleId: accentStyle?.StyleId, firstRow: true, bandedRows: true));
            table.SetColumnWidthsPoints(200, 90, 90, 90, 90);
            table.SetRowHeightPoints(0, 28);
            for (int r = 1; r < table.Rows; r++) {
                table.SetRowHeightPoints(r, 24);
            }

            for (int c = 0; c < table.Columns; c++) {
                PowerPointTableCell header = table.GetCell(0, c);
                header.Bold = true;
                header.FontSize = 12;
                header.HorizontalAlignment = A.TextAlignmentTypeValues.Center;
                header.VerticalAlignment = A.TextAnchoringTypeValues.Center;
                header.PaddingLeftPoints = 4;
                header.PaddingRightPoints = 4;
                header.PaddingTopPoints = 3;
                header.PaddingBottomPoints = 3;
            }

            for (int r = 1; r < table.Rows; r++) {
                for (int c = 0; c < table.Columns; c++) {
                    PowerPointTableCell cell = table.GetCell(r, c);
                    cell.FontSize = 11;
                    cell.VerticalAlignment = A.TextAnchoringTypeValues.Center;
                    if (c > 0) {
                        cell.HorizontalAlignment = A.TextAlignmentTypeValues.Center;
                    }
                    cell.PaddingLeftPoints = 3;
                    cell.PaddingRightPoints = 3;
                }
            }

                        // Slide 3: merged cells
            PowerPointSlide mergeSlide = presentation.AddSlide();
            mergeSlide.AddTitleCm("Merged Cells", marginCm, marginCm, content.WidthCm, titleHeightCm);
            PowerPointTable mergeTable = mergeSlide.AddTableCm(
                rows: 4,
                columns: 4,
                leftCm: marginCm,
                topCm: bodyTopCm,
                widthCm: content.WidthCm,
                heightCm: bodyHeightCm);
            mergeTable.ApplyStyle(new PowerPointTableStylePreset(styleId: accentStyle?.StyleId, firstRow: true, bandedRows: true));
            mergeTable.SetColumnWidthsEvenly();
            mergeTable.SetRowHeightPoints(0, 28);

            mergeTable.GetCell(0, 0).Text = "2025 Sales (Merged Header)";
            mergeTable.MergeCells(0, 0, 0, mergeTable.Columns - 1);

            mergeTable.GetCell(1, 0).Text = "Category";
            mergeTable.GetCell(1, 1).Text = "Q1";
            mergeTable.GetCell(1, 2).Text = "Q2";
            mergeTable.GetCell(1, 3).Text = "Q3";

            mergeTable.GetCell(2, 0).Text = "Hardware";
            mergeTable.GetCell(2, 1).Text = "120";
            mergeTable.GetCell(2, 2).Text = "140";
            mergeTable.GetCell(2, 3).Text = "160";
            mergeTable.GetCell(3, 1).Text = "90";
            mergeTable.GetCell(3, 2).Text = "110";
            mergeTable.GetCell(3, 3).Text = "130";
            mergeTable.MergeCells(2, 0, 3, 0);

            for (int r = 0; r < mergeTable.Rows; r++) {
                for (int c = 0; c < mergeTable.Columns; c++) {
                    PowerPointTableCell cell = mergeTable.GetCell(r, c);
                    cell.FontSize = r == 0 ? 12 : 11;
                    cell.VerticalAlignment = A.TextAnchoringTypeValues.Center;
                    if (r <= 1 || c > 0) {
                        cell.HorizontalAlignment = A.TextAlignmentTypeValues.Center;
                    }
                    cell.PaddingLeftPoints = 3;
                    cell.PaddingRightPoints = 3;
                }
            }
            // Slide 4: built-in styles showcase
            PowerPointSlide stylesSlide = presentation.AddSlide();
            stylesSlide.AddTitleCm("Built-in Table Styles", marginCm, marginCm, content.WidthCm, titleHeightCm);
            const double styleLabelHeightCm = 0.6;
            const double styleLabelGapCm = 0.2;
            double styleTableTopCm = bodyTopCm + styleLabelHeightCm + styleLabelGapCm;
            double styleTableHeightCm = bodyHeightCm - styleLabelHeightCm - styleLabelGapCm;
            stylesSlide.AddTextBoxCm($"Style: {accentStyle?.Name ?? "Default"}", leftColumn.LeftCm, bodyTopCm,
                leftColumn.WidthCm, styleLabelHeightCm);
            PowerPointTable leftStyleTable = stylesSlide.AddTableCm(
                data,
                o => {
                    o.HeaderCase = HeaderCase.Title;
                    o.PinFirst("Product");
                },
                includeHeaders: true,
                leftCm: leftColumn.LeftCm,
                topCm: styleTableTopCm,
                widthCm: leftColumn.WidthCm,
                heightCm: styleTableHeightCm);
            leftStyleTable.ApplyStyle(new PowerPointTableStylePreset(styleId: accentStyle?.StyleId, firstRow: true, bandedRows: true));
            leftStyleTable.SetColumnWidthsEvenly();

            stylesSlide.AddTextBoxCm($"Style: {darkStyle?.Name ?? "Banded Columns"}", rightColumn.LeftCm, bodyTopCm,
                rightColumn.WidthCm, styleLabelHeightCm);
            PowerPointTable rightStyleTable = stylesSlide.AddTableCm(
                data,
                o => {
                    o.HeaderCase = HeaderCase.Title;
                    o.PinFirst("Product");
                },
                includeHeaders: true,
                leftCm: rightColumn.LeftCm,
                topCm: styleTableTopCm,
                widthCm: rightColumn.WidthCm,
                heightCm: styleTableHeightCm);
            rightStyleTable.ApplyStyle(new PowerPointTableStylePreset(styleId: darkStyle?.StyleId, firstRow: true, bandedColumns: true));
            rightStyleTable.SetColumnWidthsEvenly();

            // Slide 5: chart from the same dataset
            PowerPointSlide chartSlide = presentation.AddSlide();
            chartSlide.AddTitleCm("Quarterly Performance", marginCm, marginCm, content.WidthCm, titleHeightCm);
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

            // Slide 6: table + chart side by side
            PowerPointSlide comboSlide = presentation.AddSlide();
            comboSlide.AddTitleCm("Table and Chart (Compact)", marginCm, marginCm, content.WidthCm, titleHeightCm);
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
            compactTable.ApplyStyle(new PowerPointTableStylePreset(styleId: lightStyle?.StyleId, firstRow: true, bandedRows: true));
            double compactWidthPoints = leftColumn.WidthPoints;
            double productWidthPoints = compactWidthPoints * 0.38;
            double quarterWidthPoints = (compactWidthPoints - productWidthPoints) / 4;
            compactTable.SetColumnWidthsPoints(productWidthPoints, quarterWidthPoints, quarterWidthPoints,
                quarterWidthPoints, quarterWidthPoints);
            compactTable.SetRowHeightPoints(0, 20);
            for (int r = 1; r < compactTable.Rows; r++) {
                compactTable.SetRowHeightPoints(r, 18);
            }
            for (int r = 0; r < compactTable.Rows; r++) {
                for (int c = 0; c < compactTable.Columns; c++) {
                    PowerPointTableCell cell = compactTable.GetCell(r, c);
                    cell.FontSize = r == 0 ? 10 : 9;
                    cell.VerticalAlignment = A.TextAnchoringTypeValues.Center;
                    if (c > 0) {
                        cell.HorizontalAlignment = A.TextAlignmentTypeValues.Center;
                    }
                }
            }

            comboSlide.AddChartCm(chartData, rightColumn.LeftCm, rightColumn.TopCm, rightColumn.WidthCm,
                rightColumn.HeightCm);

            // Slide 7: totals by quarter
            PowerPointSlide totalsSlide = presentation.AddSlide();
            totalsSlide.AddTitleCm("Totals by Quarter", marginCm, marginCm, content.WidthCm, titleHeightCm);
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

            // Slide 8: picture with caption
            PowerPointSlide imageSlide = presentation.AddSlide();
            imageSlide.AddTitleCm("Brand Assets", marginCm, marginCm, content.WidthCm, titleHeightCm);
            string imagePath = Path.Combine(AppContext.BaseDirectory, "Images", "BackgroundImage.png");
            if (File.Exists(imagePath)) {
                imageSlide.AddPictureCm(imagePath, marginCm, bodyTopCm, content.WidthCm, bodyHeightCm);
            } else {
                imageSlide.AddTextBoxCm("(image placeholder)", marginCm, bodyTopCm, content.WidthCm, 2.5);
            }

            // Slide 9: summary table
            PowerPointSlide summarySlide = presentation.AddSlide();
            summarySlide.AddTitleCm("Summary", marginCm, marginCm, content.WidthCm, titleHeightCm);
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
            summary.ApplyStyle(new PowerPointTableStylePreset(styleId: accentStyle?.StyleId, firstRow: true, bandedRows: true));

            PowerPointTableCell summaryHeader = summary.GetCell(0, 0);
            summaryHeader.Bold = true;
            summaryHeader.PaddingLeftPoints = 4;

            presentation.Save();

            Helpers.Open(filePath, openPowerPoint);
        }

        private static PowerPointTableStyleInfo? FindStyle(IReadOnlyList<PowerPointTableStyleInfo> styles, string name) {
            if (styles == null || styles.Count == 0) {
                return null;
            }

            foreach (PowerPointTableStyleInfo style in styles) {
                if (style.Name.Contains(name, StringComparison.OrdinalIgnoreCase)) {
                    return style;
                }
            }

            return null;
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


