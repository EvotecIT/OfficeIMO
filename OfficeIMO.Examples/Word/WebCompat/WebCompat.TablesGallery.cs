using System;
using OfficeIMO.Word;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Examples.Word {
    internal static partial class WebCompat {
        public static void Example_TablesGallery(string folderPath, bool openWord) {
            string filePath = System.IO.Path.Combine(folderPath, "WebCompat-Tables.docx");
            Console.WriteLine("[*] Generating: " + filePath);

            using var doc = WordDocument.Create(filePath);

            // 1) Auto table (no explicit widths) – baseline behavior
            doc.AddParagraph("Auto table (no explicit widths)");
            var tAuto = doc.AddTable(2, 3, WordTableStyle.TableGrid);
            tAuto.Rows[0].Cells[0].AddParagraph("A1", true);
            tAuto.Rows[0].Cells[1].AddParagraph("A2", true);
            tAuto.Rows[0].Cells[2].AddParagraph("A3", true);
            tAuto.Rows[1].Cells[0].AddParagraph("B1", true);
            tAuto.Rows[1].Cells[1].AddParagraph("B2", true);
            tAuto.Rows[1].Cells[2].AddParagraph("B3", true);

            // 2) Percent widths (10/90) – typical case
            doc.AddParagraph().AddText("10/90 percent widths");
            var tPct = doc.AddTable(2, 2, WordTableStyle.TableGrid);
            tPct.WidthType = TableWidthUnitValues.Pct; tPct.Width = 5000; // 100%
            tPct.ColumnWidthType = TableWidthUnitValues.Pct; tPct.ColumnWidth = new() { 500, 4500 };
            tPct.Rows[0].Cells[0].AddParagraph("10%", true);
            tPct.Rows[0].Cells[1].AddParagraph("90%", true);
            tPct.Rows[1].Cells[0].AddParagraph("10%", true);
            tPct.Rows[1].Cells[1].AddParagraph("90%", true);

            // 3) DXA widths smaller than container – previously looked half-width online
            doc.AddParagraph().AddText("DXA widths (sum smaller than container)");
            var tDxaSmall = doc.AddTable(2, 2, WordTableStyle.TableGrid);
            tDxaSmall.WidthType = TableWidthUnitValues.Pct; tDxaSmall.Width = 5000; // 100%
            tDxaSmall.ColumnWidthType = TableWidthUnitValues.Dxa; tDxaSmall.ColumnWidth = new() { 2400, 2400 };
            tDxaSmall.Rows[0].Cells[0].AddParagraph("Left", true);
            tDxaSmall.Rows[0].Cells[1].AddParagraph("Right", true);

            // 4) Merged header spanning 4 cols with detailed row below (tests grid inference)
            doc.AddParagraph().AddText("Merged header over 4 columns");
            var tMerge = doc.AddTable(2, 4, WordTableStyle.TableGrid);
            // Merge first row 4 cells
            tMerge.Rows[0].Cells[0].AddParagraph("Header spanning 4", true);
            tMerge.Rows[0].Cells[0].MergeHorizontally(3);
            // Set explicit percent widths for data row
            tMerge.ColumnWidthType = TableWidthUnitValues.Pct; tMerge.ColumnWidth = new() { 1000, 1000, 1000, 2000 };
            for (int i = 0; i < 4; i++) tMerge.Rows[1].Cells[i].AddParagraph($"C{i+1}", true);

            // 5) Many columns (7) to test rounding
            doc.AddParagraph().AddText("7 columns (percent)");
            var t7 = doc.AddTable(2, 7, WordTableStyle.TableGrid);
            t7.WidthType = TableWidthUnitValues.Pct; t7.Width = 5000;
            t7.ColumnWidthType = TableWidthUnitValues.Pct; t7.ColumnWidth = new() { 700,700,700,700,700,700, 1000 };
            for (int c = 0; c < 7; c++) t7.Rows[0].Cells[c].AddParagraph($"H{c+1}", true);
            for (int c = 0; c < 7; c++) t7.Rows[1].Cells[c].AddParagraph($"D{c+1}", true);

            doc.Save(openWord);
        }
    }
}

