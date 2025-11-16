using System;
using OfficeIMO.Word;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Examples.Word {
    internal static partial class WebCompat {
        public static void Example_TablesGallery(string folderPath, bool openWord) {
            string filePath = System.IO.Path.Combine(folderPath, "WebCompat-Tables.docx");
            Console.WriteLine("[*] Generating: " + filePath);

            using var doc = WordDocument.Create(filePath);

            // 1) Auto vs AutoFit (same content) – easy side-by-side comparison
            doc.AddParagraph("Auto table (no explicit widths) vs AutoFit to Window (100%) — same content");
            doc.AddParagraph("Auto (no explicit widths)");
            var tAuto = doc.AddTable(2, 3, WordTableStyle.TableGrid);
            // Content rows
            tAuto.Rows[0].Cells[0].AddParagraph("Cell A1", true);
            tAuto.Rows[0].Cells[1].AddParagraph("Cell A2", true);
            tAuto.Rows[0].Cells[2].AddParagraph("Cell A3", true);
            tAuto.Rows[1].Cells[0].AddParagraph("Cell B1", true);
            tAuto.Rows[1].Cells[1].AddParagraph("Cell B2", true);
            tAuto.Rows[1].Cells[2].AddParagraph("Cell B3", true);

            // AutoFit version placed right after, same content
            doc.AddParagraph("AutoFit to Window (100%)");
            var tAutoFit = doc.AddTable(2, 3, WordTableStyle.TableGrid);
            tAutoFit.AutoFitToWindow(); // sets pct 100% and fixed layout
            // Content rows
            tAutoFit.Rows[0].Cells[0].AddParagraph("Cell A1", true);
            tAutoFit.Rows[0].Cells[1].AddParagraph("Cell A2", true);
            tAutoFit.Rows[0].Cells[2].AddParagraph("Cell A3", true);
            tAutoFit.Rows[1].Cells[0].AddParagraph("Cell B1", true);
            tAutoFit.Rows[1].Cells[1].AddParagraph("Cell B2", true);
            tAutoFit.Rows[1].Cells[2].AddParagraph("Cell B3", true);

            // 2) Percent widths (10/90) – typical case
            doc.AddParagraph().AddText("10/90 percent widths");
            var tPct = doc.AddTable(2, 2, WordTableStyle.TableGrid);
            tPct.WidthType = TableWidthUnitValues.Pct; tPct.Width = 5000; // 100%
            tPct.ColumnWidthType = TableWidthUnitValues.Pct; tPct.ColumnWidth = new() { 500, 4500 };
            // Content rows
            tPct.Rows[0].Cells[0].AddParagraph("10%", true);
            tPct.Rows[0].Cells[1].AddParagraph("90%", true);
            tPct.Rows[1].Cells[0].AddParagraph("10%", true);
            tPct.Rows[1].Cells[1].AddParagraph("90%", true);

            // 3) DXA widths smaller than container – previously looked half-width online
            doc.AddParagraph().AddText("DXA widths (sum smaller than container)");
            var tDxaSmall = doc.AddTable(2, 2, WordTableStyle.TableGrid);
            tDxaSmall.WidthType = TableWidthUnitValues.Pct; tDxaSmall.Width = 5000; // 100%
            tDxaSmall.ColumnWidthType = TableWidthUnitValues.Dxa; tDxaSmall.ColumnWidth = new() { 2400, 2400 };
            // Content rows
            tDxaSmall.Rows[0].Cells[0].AddParagraph("Left", true);
            tDxaSmall.Rows[0].Cells[1].AddParagraph("Right", true);
            tDxaSmall.Rows[1].Cells[0].AddParagraph("", true);
            tDxaSmall.Rows[1].Cells[1].AddParagraph("", true);

            // 3a) Side-by-side split: Percent 30/70 vs DXA equivalent (renders identically online)
            doc.AddParagraph().AddText("30/70 split – Percent vs DXA equivalent");
            // Percent version
            doc.AddParagraph().AddText("30/70 (percent)");
            var t3070Pct = doc.AddTable(2, 2, WordTableStyle.TableGrid);
            t3070Pct.WidthType = TableWidthUnitValues.Pct; t3070Pct.Width = 5000; // 100%
            t3070Pct.ColumnWidthType = TableWidthUnitValues.Pct; t3070Pct.ColumnWidth = new() { 1500, 3500 };
            // Content rows
            t3070Pct.Rows[0].Cells[0].AddParagraph("30%", true);
            t3070Pct.Rows[0].Cells[1].AddParagraph("70%", true);
            t3070Pct.Rows[1].Cells[0].AddParagraph("", true);
            t3070Pct.Rows[1].Cells[1].AddParagraph("", true);
            // DXA version (any 3:7 proportion will be normalized to container width on save)
            doc.AddParagraph().AddText("30/70 (DXA equivalent)");
            var t3070Dxa = doc.AddTable(2, 2, WordTableStyle.TableGrid);
            t3070Dxa.WidthType = TableWidthUnitValues.Pct; t3070Dxa.Width = 5000; // 100%
            t3070Dxa.ColumnWidthType = TableWidthUnitValues.Dxa; t3070Dxa.ColumnWidth = new() { 3000, 7000 };
            // Content rows
            t3070Dxa.Rows[0].Cells[0].AddParagraph("~30% (DXA)", true);
            t3070Dxa.Rows[0].Cells[1].AddParagraph("~70% (DXA)", true);
            t3070Dxa.Rows[1].Cells[0].AddParagraph("", true);
            t3070Dxa.Rows[1].Cells[1].AddParagraph("", true);

            // 4) Merged header: 1 cell in first row (spans 4), then 4 data columns
            // 4) True merged header (hMerge → gridSpan on Save)
            // 4a) Auto (no table width)
            doc.AddParagraph().AddText("Merged header (true) — Auto (no table width)");
            var tMergeTrueAuto = doc.AddTable(2, 4, WordTableStyle.TableGrid);
            tMergeTrueAuto.Rows[0].Cells[0].AddParagraph("Header spanning 4", true);
            tMergeTrueAuto.Rows[0].Cells[0].MergeHorizontally(3);
            tMergeTrueAuto.Rows[0].Cells[0].ShadingFillColorHex = "e6e6e6";
            tMergeTrueAuto.Rows[0].Cells[0].Paragraphs[0].ParagraphAlignment = JustificationValues.Center;
            tMergeTrueAuto.Rows[0].Cells[0].VerticalAlignment = TableVerticalAlignmentValues.Center;
            tMergeTrueAuto.Rows[0].Cells[0].Borders.RightStyle = BorderValues.Single;
            for (int i = 0; i < 4; i++) tMergeTrueAuto.Rows[1].Cells[i].AddParagraph($"C{i+1}", true);
            tMergeTrueAuto.ColumnWidthType = TableWidthUnitValues.Pct; tMergeTrueAuto.ColumnWidth = new() { 500, 1000, 1000, 2500 };

            // 4b) AutoFit to Window (100%) — same content, full width
            doc.AddParagraph().AddText("Merged header (true) — AutoFit to Window (100%)");
            var tMergeTrueFit = doc.AddTable(2, 4, WordTableStyle.TableGrid);
            tMergeTrueFit.AutoFitToWindow();
            tMergeTrueFit.Rows[0].Cells[0].AddParagraph("Header spanning 4", true);
            tMergeTrueFit.Rows[0].Cells[0].MergeHorizontally(3);
            tMergeTrueFit.Rows[0].Cells[0].ShadingFillColorHex = "e6e6e6";
            tMergeTrueFit.Rows[0].Cells[0].Paragraphs[0].ParagraphAlignment = JustificationValues.Center;
            tMergeTrueFit.Rows[0].Cells[0].VerticalAlignment = TableVerticalAlignmentValues.Center;
            tMergeTrueFit.Rows[0].Cells[0].Borders.RightStyle = BorderValues.Single;
            for (int i = 0; i < 4; i++) tMergeTrueFit.Rows[1].Cells[i].AddParagraph($"C{i+1}", true);
            tMergeTrueFit.ColumnWidthType = TableWidthUnitValues.Pct; tMergeTrueFit.ColumnWidth = new() { 500, 1000, 1000, 2500 };

            // 5) Many columns (7) to test rounding (make last column very wide for contrast)
            doc.AddParagraph().AddText("7 columns (percent)");
            var t7 = doc.AddTable(2, 7, WordTableStyle.TableGrid);
            t7.WidthType = TableWidthUnitValues.Pct; t7.Width = 5000;
            t7.ColumnWidthType = TableWidthUnitValues.Pct; t7.ColumnWidth = new() { 200,200,200,200,200,200, 3800 };
            for (int c = 0; c < 7; c++) t7.Rows[0].Cells[c].AddParagraph($"H{c+1}", true);
            for (int c = 0; c < 7; c++) t7.Rows[1].Cells[c].AddParagraph($"D{c+1}", true);

            // 6) Two-row header with column groups (A,B) expanding to four columns
            doc.AddParagraph().AddText("Two-row header (A,B groups) → 4 columns");
            var tGroups = doc.AddTable(3, 4, WordTableStyle.TableGrid);
            // First header row: A spans 2, B spans 2
            tGroups.Rows[0].Cells[0].AddParagraph("Group A", true);
            tGroups.Rows[0].Cells[0].MergeHorizontally(1);
            tGroups.Rows[0].Cells[2].AddParagraph("Group B", true);
            tGroups.Rows[0].Cells[2].MergeHorizontally(1);
            // Second header row: A1 A2 B1 B2
            tGroups.Rows[1].Cells[0].AddParagraph("A1", true);
            tGroups.Rows[1].Cells[1].AddParagraph("A2", true);
            tGroups.Rows[1].Cells[2].AddParagraph("B1", true);
            tGroups.Rows[1].Cells[3].AddParagraph("B2", true);
            // Data row
            tGroups.Rows[2].Cells[0].AddParagraph("v1", true);
            tGroups.Rows[2].Cells[1].AddParagraph("v2", true);
            tGroups.Rows[2].Cells[2].AddParagraph("v3", true);
            tGroups.Rows[2].Cells[3].AddParagraph("v4", true);

            // 7) Narrow layouts — options users can replicate online
            // 7a) Preferred width 60% and centered
            doc.AddParagraph().AddText("Narrow layout: preferred width 60% (centered)");
            var tNarrowPct = doc.AddTable(2, 2, WordTableStyle.TableGrid);
            // Use fixed layout at 60% to make Word Online honor the percentage strictly
            tNarrowPct.SetTableLayout(WordTableLayoutType.FixedWidth, 60);
            tNarrowPct.Alignment = TableRowAlignmentValues.Center;
            tNarrowPct.ColumnWidthType = TableWidthUnitValues.Pct; tNarrowPct.ColumnWidth = new() { 2500, 2500 };
            tNarrowPct.Rows[0].Cells[0].AddParagraph("L", true);
            tNarrowPct.Rows[0].Cells[1].AddParagraph("R", true);
            tNarrowPct.Rows[1].Cells[0].AddParagraph("L", true);
            tNarrowPct.Rows[1].Cells[1].AddParagraph("R", true);

            // 7b) Container table (1x1) that constrains inner table like a narrow block
            doc.AddParagraph().AddText("Narrow layout: 1×1 container table (centered) with inner table");
            var container = doc.AddTable(1, 1, WordTableStyle.TableGrid);
            container.SetTableLayout(WordTableLayoutType.FixedWidth, 60); // 60%
            container.Alignment = TableRowAlignmentValues.Center;
            // Build the inner table inside the single cell
            var inner = container.Rows[0].Cells[0].AddTable(2, 2, WordTableStyle.TableGrid);
            inner.SetTableLayout(WordTableLayoutType.FixedWidth, 100); // inner 100% of container cell
            inner.ColumnWidthType = TableWidthUnitValues.Pct; inner.ColumnWidth = new() { 2500, 2500 };
            inner.Rows[0].Cells[0].AddParagraph("Inner L", true);
            inner.Rows[0].Cells[1].AddParagraph("Inner R", true);
            inner.Rows[1].Cells[0].AddParagraph("Inner L", true);
            inner.Rows[1].Cells[1].AddParagraph("Inner R", true);

            // 7c) Narrow preferred width 60% but left-aligned
            doc.AddParagraph().AddText("Narrow layout: preferred width 60% (left-aligned)");
            var tNarrowPctLeft = doc.AddTable(2, 2, WordTableStyle.TableGrid);
            tNarrowPctLeft.SetTableLayout(WordTableLayoutType.FixedWidth, 60); // 60%
            tNarrowPctLeft.Alignment = TableRowAlignmentValues.Left;
            tNarrowPctLeft.ColumnWidthType = TableWidthUnitValues.Pct; tNarrowPctLeft.ColumnWidth = new() { 2500, 2500 };
            tNarrowPctLeft.Rows[0].Cells[0].AddParagraph("L", true);
            tNarrowPctLeft.Rows[0].Cells[1].AddParagraph("R", true);
            tNarrowPctLeft.Rows[1].Cells[0].AddParagraph("L", true);
            tNarrowPctLeft.Rows[1].Cells[1].AddParagraph("R", true);

            // 7d) Container with fixed DXA width (5 inches) and centered
            doc.AddParagraph().AddText("Narrow layout: 1×1 container (fixed 5\" width, centered) with inner table");
            var containerFixed = doc.AddTable(1, 1, WordTableStyle.TableGrid);
            containerFixed.SetTableLayout(WordTableLayoutType.FixedWidth, null); // switch to Fixed layout
            containerFixed.WidthType = TableWidthUnitValues.Dxa; containerFixed.Width = 7200; // 5 inches * 1440 twips
            containerFixed.Alignment = TableRowAlignmentValues.Center;
            var innerFixed = containerFixed.Rows[0].Cells[0].AddTable(2, 2, WordTableStyle.TableGrid);
            innerFixed.SetTableLayout(WordTableLayoutType.FixedWidth, 100); // inner 100% of container cell
            innerFixed.WidthType = TableWidthUnitValues.Pct; innerFixed.Width = 5000;
            innerFixed.ColumnWidthType = TableWidthUnitValues.Pct; innerFixed.ColumnWidth = new() { 2500, 2500 };
            innerFixed.Rows[0].Cells[0].AddParagraph("Inner L", true);
            innerFixed.Rows[0].Cells[1].AddParagraph("Inner R", true);
            innerFixed.Rows[1].Cells[0].AddParagraph("Inner L", true);
            innerFixed.Rows[1].Cells[1].AddParagraph("Inner R", true);

            doc.Save(openWord);
        }
    }
}
