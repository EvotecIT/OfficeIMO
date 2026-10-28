using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Data;
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
            using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
            const double marginCm = 1.5;
            const double gutterCm = 1.0;
            const double titleHeightCm = 1.4;
            double slideHeightCm = presentation.SlideSize.Height.ToCentimeters();
            PowerPointLayoutBox content = presentation.SlideSize.GetContentBox(OfficeIMO.OfficeLength.Centimeters(marginCm));
            double bodyTopCm = 3.8;
            double bodyHeightCm = slideHeightCm - bodyTopCm - marginCm;
            long bodyTop = PowerPointUnits.FromCentimeters(bodyTopCm);
            long bodyHeight = PowerPointUnits.FromCentimeters(bodyHeightCm);
            PowerPointLayoutBox[] columns = presentation.SlideSize.GetColumns(2,OfficeIMO.OfficeLength.Centimeters(marginCm),OfficeIMO.OfficeLength.Centimeters(gutterCm));
            PowerPointLayoutBox leftColumn = new(OfficeIMO.OfficeLength.Emus(columns[0].Left.ToEmus()),OfficeIMO.OfficeLength.Emus(bodyTop),OfficeIMO.OfficeLength.Emus(columns[0].Width.ToEmus()),OfficeIMO.OfficeLength.Emus(bodyHeight));
            PowerPointLayoutBox rightColumn = new(OfficeIMO.OfficeLength.Emus(columns[1].Left.ToEmus()),OfficeIMO.OfficeLength.Emus(bodyTop),OfficeIMO.OfficeLength.Emus(columns[1].Width.ToEmus()),OfficeIMO.OfficeLength.Emus(bodyHeight));

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
            PowerPointTextBox title = titleSlide.AddTitle("Quarterly Sales Overview",global::OfficeIMO.OfficeLength.Centimeters(marginCm),global::OfficeIMO.OfficeLength.Centimeters(marginCm),global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(titleHeightCm));
            title.FontSize = global::OfficeIMO.OfficeLength.Points(32);
            title.Color = "1F4E79";
            titleSlide.AddTextBox("Tables created from data objects with built-in styles.",global::OfficeIMO.OfficeLength.Centimeters(marginCm),global::OfficeIMO.OfficeLength.Centimeters(3.2),global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(1.1));

            // Slide 2: table from objects
            PowerPointSlide tableSlide = presentation.AddSlide();
            tableSlide.AddTitle("Sales by Product",global::OfficeIMO.OfficeLength.Centimeters(marginCm),global::OfficeIMO.OfficeLength.Centimeters(marginCm),global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(titleHeightCm));
            PowerPointTable table = tableSlide.AddTable(
                data,                o => {
                    o.HeaderCase = HeaderCase.Title;
                    o.PinFirst("Product");
                },                includeHeaders: true,left: global::OfficeIMO.OfficeLength.Centimeters(marginCm),top: global::OfficeIMO.OfficeLength.Centimeters(bodyTopCm),width: global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),height: global::OfficeIMO.OfficeLength.Centimeters(bodyHeightCm));
            table.ApplyStyle(new PowerPointTableStylePreset(styleId: accentStyle?.StyleId, firstRow: true, bandedRows: true));
            table.SetColumnWidths(global::OfficeIMO.OfficeLength.Points(200),global::OfficeIMO.OfficeLength.Points(90),global::OfficeIMO.OfficeLength.Points(90),global::OfficeIMO.OfficeLength.Points(90),global::OfficeIMO.OfficeLength.Points(90));
            table.SetRowHeight(0,global::OfficeIMO.OfficeLength.Points(28));
            for (int r = 1; r < table.Rows; r++) {
                table.SetRowHeight(r,global::OfficeIMO.OfficeLength.Points(24));
            }

            for (int c = 0; c < table.Columns; c++) {
                PowerPointTableCell header = table.GetCell(0, c);
                header.Bold = true;
                header.FontSize = global::OfficeIMO.OfficeLength.Points(12);
                header.HorizontalAlignment = PowerPointTextAlignment.Center;
                header.VerticalAlignment = PowerPointTextVerticalAlignment.Center;
                header.PaddingLeft = global::OfficeIMO.OfficeLength.Points(4);
                header.PaddingRight = global::OfficeIMO.OfficeLength.Points(4);
                header.PaddingTop = global::OfficeIMO.OfficeLength.Points(3);
                header.PaddingBottom = global::OfficeIMO.OfficeLength.Points(3);
            }

            for (int r = 1; r < table.Rows; r++) {
                for (int c = 0; c < table.Columns; c++) {
                    PowerPointTableCell cell = table.GetCell(r, c);
                    cell.FontSize = global::OfficeIMO.OfficeLength.Points(11);
                    cell.VerticalAlignment = PowerPointTextVerticalAlignment.Center;
                    if (c > 0) {
                        cell.HorizontalAlignment = PowerPointTextAlignment.Center;
                    }
                    cell.PaddingLeft = global::OfficeIMO.OfficeLength.Points(3);
                    cell.PaddingRight = global::OfficeIMO.OfficeLength.Points(3);
                }
            }

                        // Slide 3: merged cells
            PowerPointSlide mergeSlide = presentation.AddSlide();
            mergeSlide.AddTitle("Merged Cells",global::OfficeIMO.OfficeLength.Centimeters(marginCm),global::OfficeIMO.OfficeLength.Centimeters(marginCm),global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(titleHeightCm));
            PowerPointTable mergeTable = mergeSlide.AddTable(
                rows: 4,                columns: 4,left: global::OfficeIMO.OfficeLength.Centimeters(marginCm),top: global::OfficeIMO.OfficeLength.Centimeters(bodyTopCm),width: global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),height: global::OfficeIMO.OfficeLength.Centimeters(bodyHeightCm));
            mergeTable.ApplyStyle(new PowerPointTableStylePreset(styleId: accentStyle?.StyleId, firstRow: true, bandedRows: true));
            mergeTable.SetColumnWidthsEvenly();
            mergeTable.SetRowHeight(0,global::OfficeIMO.OfficeLength.Points(28));

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
                    cell.FontSize = global::OfficeIMO.OfficeLength.Points(r == 0 ? 12 : 11);
                    cell.VerticalAlignment = PowerPointTextVerticalAlignment.Center;
                    if (r <= 1 || c > 0) {
                        cell.HorizontalAlignment = PowerPointTextAlignment.Center;
                    }
                    cell.PaddingLeft = global::OfficeIMO.OfficeLength.Points(3);
                    cell.PaddingRight = global::OfficeIMO.OfficeLength.Points(3);
                }
            }
            // Slide 4: built-in styles showcase
            PowerPointSlide stylesSlide = presentation.AddSlide();
            stylesSlide.AddTitle("Built-in Table Styles",global::OfficeIMO.OfficeLength.Centimeters(marginCm),global::OfficeIMO.OfficeLength.Centimeters(marginCm),global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(titleHeightCm));
            const double styleLabelHeightCm = 0.6;
            const double styleLabelGapCm = 0.2;
            double styleTableTopCm = bodyTopCm + styleLabelHeightCm + styleLabelGapCm;
            double styleTableHeightCm = bodyHeightCm - styleLabelHeightCm - styleLabelGapCm;
            stylesSlide.AddTextBox($"Style: {accentStyle?.Name ?? "Default"}",global::OfficeIMO.OfficeLength.Centimeters(leftColumn.Left.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(bodyTopCm),                global::OfficeIMO.OfficeLength.Centimeters(leftColumn.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(styleLabelHeightCm));
            PowerPointTable leftStyleTable = stylesSlide.AddTable(
                data,                o => {
                    o.HeaderCase = HeaderCase.Title;
                    o.PinFirst("Product");
                },                includeHeaders: true,left: global::OfficeIMO.OfficeLength.Centimeters(leftColumn.Left.ToCentimeters()),top: global::OfficeIMO.OfficeLength.Centimeters(styleTableTopCm),width: global::OfficeIMO.OfficeLength.Centimeters(leftColumn.Width.ToCentimeters()),height: global::OfficeIMO.OfficeLength.Centimeters(styleTableHeightCm));
            leftStyleTable.ApplyStyle(new PowerPointTableStylePreset(styleId: accentStyle?.StyleId, firstRow: true, bandedRows: true));
            leftStyleTable.SetColumnWidthsEvenly();

            stylesSlide.AddTextBox($"Style: {darkStyle?.Name ?? "Banded Columns"}",global::OfficeIMO.OfficeLength.Centimeters(rightColumn.Left.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(bodyTopCm),                global::OfficeIMO.OfficeLength.Centimeters(rightColumn.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(styleLabelHeightCm));
            PowerPointTable rightStyleTable = stylesSlide.AddTable(
                data,                o => {
                    o.HeaderCase = HeaderCase.Title;
                    o.PinFirst("Product");
                },                includeHeaders: true,left: global::OfficeIMO.OfficeLength.Centimeters(rightColumn.Left.ToCentimeters()),top: global::OfficeIMO.OfficeLength.Centimeters(styleTableTopCm),width: global::OfficeIMO.OfficeLength.Centimeters(rightColumn.Width.ToCentimeters()),height: global::OfficeIMO.OfficeLength.Centimeters(styleTableHeightCm));
            rightStyleTable.ApplyStyle(new PowerPointTableStylePreset(styleId: darkStyle?.StyleId, firstRow: true, bandedColumns: true));
            rightStyleTable.SetColumnWidthsEvenly();

            // Slide 5: chart from the same dataset
            PowerPointSlide chartSlide = presentation.AddSlide();
            chartSlide.AddTitle("Quarterly Performance",global::OfficeIMO.OfficeLength.Centimeters(marginCm),global::OfficeIMO.OfficeLength.Centimeters(marginCm),global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(titleHeightCm));
            OfficeChartData chartData = new(
                data.Select(d => d.Product),
                new[] {
                    new OfficeSeries("Q1", data.Select(d => (double)d.Q1)),
                    new OfficeSeries("Q2", data.Select(d => (double)d.Q2)),
                    new OfficeSeries("Q3", data.Select(d => (double)d.Q3)),
                    new OfficeSeries("Q4", data.Select(d => (double)d.Q4))
                });
            chartSlide.AddChart(OfficeChartType.ColumnClustered,chartData,global::OfficeIMO.OfficeLength.Centimeters(marginCm),global::OfficeIMO.OfficeLength.Centimeters(bodyTopCm),                global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(bodyHeightCm));
            chartSlide.Notes.Text = "Chart and table share the same source data.";

            // Slide 6: table + chart side by side
            PowerPointSlide comboSlide = presentation.AddSlide();
            comboSlide.AddTitle("Table and Chart (Compact)",global::OfficeIMO.OfficeLength.Centimeters(marginCm),global::OfficeIMO.OfficeLength.Centimeters(marginCm),global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(titleHeightCm));
            PowerPointTable compactTable = comboSlide.AddTable(
                data,                o => {
                    o.HeaderCase = HeaderCase.Title;
                    o.PinFirst("Product");
                },                includeHeaders: true,left: global::OfficeIMO.OfficeLength.Centimeters(leftColumn.Left.ToCentimeters()),top: global::OfficeIMO.OfficeLength.Centimeters(leftColumn.Top.ToCentimeters()),width: global::OfficeIMO.OfficeLength.Centimeters(leftColumn.Width.ToCentimeters()),height: global::OfficeIMO.OfficeLength.Centimeters(leftColumn.Height.ToCentimeters()));
            compactTable.ApplyStyle(new PowerPointTableStylePreset(styleId: lightStyle?.StyleId, firstRow: true, bandedRows: true));
            double compactWidthPoints = leftColumn.Width.ToPoints();
            double productWidthPoints = compactWidthPoints * 0.38;
            double quarterWidthPoints = (compactWidthPoints - productWidthPoints) / 4;
            compactTable.SetColumnWidths(global::OfficeIMO.OfficeLength.Points(productWidthPoints),global::OfficeIMO.OfficeLength.Points(quarterWidthPoints),global::OfficeIMO.OfficeLength.Points(quarterWidthPoints),                global::OfficeIMO.OfficeLength.Points(quarterWidthPoints),global::OfficeIMO.OfficeLength.Points(quarterWidthPoints));
            compactTable.SetRowHeight(0,global::OfficeIMO.OfficeLength.Points(20));
            for (int r = 1; r < compactTable.Rows; r++) {
                compactTable.SetRowHeight(r,global::OfficeIMO.OfficeLength.Points(18));
            }
            for (int r = 0; r < compactTable.Rows; r++) {
                for (int c = 0; c < compactTable.Columns; c++) {
                    PowerPointTableCell cell = compactTable.GetCell(r, c);
                    cell.FontSize = global::OfficeIMO.OfficeLength.Points(r == 0 ? 10 : 9);
                    cell.VerticalAlignment = PowerPointTextVerticalAlignment.Center;
                    if (c > 0) {
                        cell.HorizontalAlignment = PowerPointTextAlignment.Center;
                    }
                }
            }

            comboSlide.AddChart(OfficeChartType.ColumnClustered,chartData,global::OfficeIMO.OfficeLength.Centimeters(rightColumn.Left.ToCentimeters()),                global::OfficeIMO.OfficeLength.Centimeters(rightColumn.Top.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(rightColumn.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(rightColumn.Height.ToCentimeters()));

            // Slide 7: totals by quarter
            PowerPointSlide totalsSlide = presentation.AddSlide();
            totalsSlide.AddTitle("Totals by Quarter",global::OfficeIMO.OfficeLength.Centimeters(marginCm),global::OfficeIMO.OfficeLength.Centimeters(marginCm),global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(titleHeightCm));
            var quarterLabels = new[] { "Q1", "Q2", "Q3", "Q4" };
            var totals = new[] {
                data.Sum(d => d.Q1),
                data.Sum(d => d.Q2),
                data.Sum(d => d.Q3),
                data.Sum(d => d.Q4)
            };
            OfficeChartData totalsData = new(
                quarterLabels,
                new[] { new OfficeSeries("Total", totals.Select(t => (double)t)) });
            totalsSlide.AddChart(OfficeChartType.ColumnClustered,totalsData,global::OfficeIMO.OfficeLength.Centimeters(marginCm),global::OfficeIMO.OfficeLength.Centimeters(bodyTopCm),                global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(bodyHeightCm));

            // Slide 8: picture with caption
            PowerPointSlide imageSlide = presentation.AddSlide();
            imageSlide.AddTitle("Brand Assets",global::OfficeIMO.OfficeLength.Centimeters(marginCm),global::OfficeIMO.OfficeLength.Centimeters(marginCm),global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(titleHeightCm));
            string imagePath = Path.Combine(AppContext.BaseDirectory, "Images", "BackgroundImage.png");
            if (File.Exists(imagePath)) {
                imageSlide.AddImage(imagePath, OfficeIMO.OfficeLength.Centimeters(marginCm), OfficeIMO.OfficeLength.Centimeters(bodyTopCm), OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()), OfficeIMO.OfficeLength.Centimeters(bodyHeightCm));
            } else {
                imageSlide.AddTextBox("(image placeholder)",global::OfficeIMO.OfficeLength.Centimeters(marginCm),global::OfficeIMO.OfficeLength.Centimeters(bodyTopCm),global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(2.5));
            }

            // Slide 9: summary table
            PowerPointSlide summarySlide = presentation.AddSlide();
            summarySlide.AddTitle("Summary",global::OfficeIMO.OfficeLength.Centimeters(marginCm),global::OfficeIMO.OfficeLength.Centimeters(marginCm),global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(titleHeightCm));
            var summaryRows = new[] {
                new SummaryRow { Metric = "Total Q1", Value = totals[0] },
                new SummaryRow { Metric = "Total Q2", Value = totals[1] },
                new SummaryRow { Metric = "Total Q3", Value = totals[2] },
                new SummaryRow { Metric = "Total Q4", Value = totals[3] }
            };
            PowerPointTable summary = summarySlide.AddTable(
                summaryRows,                o => o.HeaderCase = HeaderCase.Title,                includeHeaders: true,left: global::OfficeIMO.OfficeLength.Centimeters(marginCm),top: global::OfficeIMO.OfficeLength.Centimeters(bodyTopCm),width: global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),height: global::OfficeIMO.OfficeLength.Centimeters(bodyHeightCm));
            summary.ApplyStyle(new PowerPointTableStylePreset(styleId: accentStyle?.StyleId, firstRow: true, bandedRows: true));

            PowerPointTableCell summaryHeader = summary.GetCell(0, 0);
            summaryHeader.Bold = true;
            summaryHeader.PaddingLeft = global::OfficeIMO.OfficeLength.Points(4);

            presentation.Save();

            if (openPowerPoint) ExampleFileLauncher.Open(filePath);
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


