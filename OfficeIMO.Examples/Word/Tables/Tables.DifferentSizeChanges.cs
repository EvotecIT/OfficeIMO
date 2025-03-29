using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word;

internal static partial class Tables {
    internal static void Example_DifferentTableSizes(string folderPath, bool openWord) {
        Console.WriteLine("[*] Creating standard document with tables of different sizes");
        string filePath = System.IO.Path.Combine(folderPath, "Document with Tables of different sizes.docx");
        using (WordDocument document = WordDocument.Create(filePath)) {
            document.AddParagraph("Demonstrating Various Table Layouts and Sizes").Bold = true;
            document.AddParagraph();

            // --- Basic Auto Width Table (Default Behavior) ---
            document.AddParagraph("Table 1: Default Table (Auto width, visually fits window)").Style = WordParagraphStyles.Heading1;
            WordTable wordTable = document.AddTable(3, 4, WordTableStyle.PlainTable1);
            wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Default";
            Console.WriteLine($"Table 1 - Default Layout: {wordTable.LayoutMode}");
            document.AddParagraph();

            // --- Percentage Width & Alignment ---
            document.AddParagraph("Table 2: Percentage Width (Invalid Value) & Centered").Style = WordParagraphStyles.Heading1;
            WordTable wordTable1 = document.AddTable(2, 6, WordTableStyle.PlainTable1);
            wordTable1.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
            wordTable1.Rows[1].Cells[0].Paragraphs[0].Text = "Longer text example to show wrapping...";
            // Setting width=100 with Pct is invalid (should be 5000 for 100%)
            // Word will likely treat this strangely or default
            wordTable1.WidthType = TableWidthUnitValues.Pct;
            wordTable1.Width = 100;
            wordTable1.Alignment = TableRowAlignmentValues.Center;
            Console.WriteLine($"Table 2 - Invalid Pct Width Layout: {wordTable1.LayoutMode}");
            document.AddParagraph();

            // --- Default Width (AutoFit to Window Visual) ---
            document.AddParagraph("Table 3: Explicitly Default (Visually AutoFit to Window)").Style = WordParagraphStyles.Heading1;
            WordTable wordTable2 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
            wordTable2.Rows[0].Cells[0].Paragraphs[0].Text = "Default";
            Console.WriteLine($"Table 3 - Explicit Default Layout: {wordTable2.LayoutMode}");
            document.AddParagraph();

            // --- Full Width (100% Percentage) ---
            document.AddParagraph("Table 4: Full Width (100% Pct)").Style = WordParagraphStyles.Heading1;
            WordTable wordTable3 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
            wordTable3.WidthType = TableWidthUnitValues.Pct;
            wordTable3.Width = 5000; // 5000 = 100%
            wordTable3.Rows[0].Cells[0].Paragraphs[0].Text = "100%";
            Console.WriteLine($"Table 4 - 100% Pct Layout: {wordTable3.LayoutMode}");
            document.AddParagraph();

            // --- Specific Percentage Width ---
            document.AddParagraph("Table 5: 50% Width (Pct)").Style = WordParagraphStyles.Heading1;
            WordTable wordTable4 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
            wordTable4.WidthType = TableWidthUnitValues.Pct;
            wordTable4.Width = 2500; // 50 * 50 = 2500 = 50%
            wordTable4.Rows[0].Cells[0].Paragraphs[0].Text = "50%";
            Console.WriteLine($"Table 5 - 50% Pct Layout: {wordTable4.LayoutMode}");
            document.AddParagraph();

            // --- Setting Width and then AutoFitting to Window ---
            document.AddParagraph("Table 6: Initially 50%, then AutoFit to Window").Style = WordParagraphStyles.Heading1;
            WordTable wordTable5 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
            wordTable5.WidthType = TableWidthUnitValues.Pct;
            wordTable5.Width = 2500; // Set to 50%
            Console.WriteLine($"Table 6 - Before AutoFit Window: {wordTable5.LayoutMode}");
            wordTable5.AutoFitToWindow(); // Now make it 100% with distributed columns
            wordTable5.Rows[0].Cells[0].Paragraphs[0].Text = "AutoFit Window";
            Console.WriteLine($"Table 6 - After AutoFit Window: {wordTable5.LayoutMode}");
            document.AddParagraph();

            // --- Using SetFixedWidth Method ---
            document.AddParagraph("Table 7: Set Fixed Width 50% (Method)").Style = WordParagraphStyles.Heading1;
            WordTable wordTable6 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
            wordTable6.SetFixedWidth(50);
            wordTable6.Rows[0].Cells[0].Paragraphs[0].Text = "Fixed 50%";
            Console.WriteLine($"Table 7 - SetFixedWidth(50): {wordTable6.LayoutMode}");
            document.AddParagraph();

            document.AddParagraph("Table 8: Set Fixed Width 75% (Method)").Style = WordParagraphStyles.Heading1;
            WordTable wordTable7 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
            wordTable7.SetFixedWidth(75);
            wordTable7.Rows[0].Cells[0].Paragraphs[0].Text = "Fixed 75%";
            Console.WriteLine($"Table 8 - SetFixedWidth(75): {wordTable7.LayoutMode}");
            document.AddParagraph();

            // --- Setting Individual Column Widths (May interfere with table-level settings) ---
            document.AddParagraph("Table 9: Individual Column Widths (Pct)").Style = WordParagraphStyles.Heading1;
            WordTable wordTable8 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
            wordTable8.Rows[0].Cells[0].Paragraphs[0].Text = "Col Widths";
            // Setting individual widths often requires careful management
            wordTable8.ColumnWidth = new List<int>() { 1000, 500, 500, 750 }; // Example Pct widths
            wordTable8.ColumnWidthType = TableWidthUnitValues.Pct;
            // Need to ensure table width matches column sum if Fixed/Pct
            wordTable8.Width = 1000 + 500 + 500 + 750; // Sum = 2750 (55%)
            wordTable8.WidthType = TableWidthUnitValues.Pct;
            wordTable8.LayoutType = TableLayoutValues.Fixed;
            Console.WriteLine($"Table 9 - Individual Pct Cols: {wordTable8.LayoutMode}");
            document.AddParagraph();

            document.AddParagraph("Table 10: More Individual Column Widths (Pct)").Style = WordParagraphStyles.Heading1;
            WordTable wordTable9 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
            wordTable9.Rows[0].Cells[0].Paragraphs[0].Text = "More Cols";
            wordTable9.ColumnWidth = new List<int>() { 1000, 500, 500, 750 };
            wordTable9.ColumnWidthType = TableWidthUnitValues.Pct;
            wordTable9.WidthType = TableWidthUnitValues.Pct; // Setting table type is good practice
            wordTable9.Width = 2750; // Ensure table width matches
            wordTable9.LayoutType = TableLayoutValues.Fixed;
            Console.WriteLine($"Table 10 - More Pct Cols: {wordTable9.LayoutMode}");
            document.AddParagraph();

            // --- AutoFit Window after setting column widths ---
            document.AddParagraph("Table 11: Set Col Widths, then AutoFit Window").Style = WordParagraphStyles.Heading1;
            WordTable wordTable10 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
            wordTable10.Rows[0].Cells[0].Paragraphs[0].Text = "Cols then Fit";
            wordTable10.ColumnWidth = new List<int>() { 1000, 500, 500, 750 };
            wordTable10.ColumnWidthType = TableWidthUnitValues.Pct;
            Console.WriteLine($"Table 11 - Before AutoFit Window: {wordTable10.LayoutMode}");
            wordTable10.AutoFitToWindow(); // Overrides column settings for 100% width
            Console.WriteLine($"Table 11 - After AutoFit Window: {wordTable10.LayoutMode}");
            document.AddParagraph();

            // --- Manually setting distributed widths ---
            document.AddParagraph("Table 12: Manually Distributed Columns (Approx 41%)").Style = WordParagraphStyles.Heading1;
            WordTable wordTable11 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
            wordTable11.Rows[0].Cells[0].Paragraphs[0].Text = "Manual Dist";
            // Sum = 2748 (approx 55%), let's aim for 4 columns at ~13.75% = 687.5, use 687
            wordTable11.ColumnWidth = new List<int>() { 687, 687, 687, 687 }; // Sum = 2748
            wordTable11.ColumnWidthType = TableWidthUnitValues.Pct;
            wordTable11.Width = 2748; // Set table width to match sum
            wordTable11.WidthType = TableWidthUnitValues.Pct;
            wordTable11.LayoutType = TableLayoutValues.Fixed;
            Console.WriteLine($"Table 12 - Manual Distribution: {wordTable11.LayoutMode}");
            document.AddParagraph();

            // --- Setting 100% via equal Column Widths ---
            document.AddParagraph("Table 13: 100% via Equal Column Widths").Style = WordParagraphStyles.Heading1;
            WordTable wordTable12 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
            wordTable12.Rows[0].Cells[0].Paragraphs[0].Text = "100% via Cols";
            wordTable12.ColumnWidth = new List<int>() { 1250, 1250, 1250, 1250 }; // 4 * 1250 = 5000 (100%)
            wordTable12.ColumnWidthType = TableWidthUnitValues.Pct;
            wordTable12.Width = 5000; // Set table width to match
            wordTable12.WidthType = TableWidthUnitValues.Pct;
            // Optional: Could set LayoutType = Fixed, but Window behavior is implicit here
            Console.WriteLine($"Table 13 - 100% via Cols: {wordTable12.LayoutMode}");
            document.AddParagraph();

            // --- Demonstrating Layout Changes with GetCurrentLayoutType / LayoutMode ---
            document.AddParagraph("Table 14: Demonstrate Layout Changes (AutoFit)").Style = WordParagraphStyles.Heading1;
            WordTable wordTable13 = document.AddTable(4, 4, WordTableStyle.PlainTable1);
            wordTable13.Rows[0].Cells[0].Paragraphs[0].Text = "Content";
            wordTable13.Rows[1].Cells[1].Paragraphs[0].Text = "More Content";
            wordTable13.Rows[2].Cells[2].Paragraphs[0].Text = "Wider Content Here";
            Console.WriteLine($"Table 14 - Initial: {wordTable13.LayoutMode}");
            // Set via raw property
            wordTable13.LayoutType = TableLayoutValues.Autofit;
            Console.WriteLine($"Table 14 - After LayoutType=Autofit: {wordTable13.LayoutMode}");
            // Set via new property
            wordTable13.LayoutMode = WordTableLayoutType.AutoFitToContents;
            Console.WriteLine($"Table 14 - After LayoutMode=AutoFitToContents: {wordTable13.LayoutMode}");
            document.AddParagraph();

            document.AddParagraph("Table 15: Demonstrate Layout Changes (Fixed 50%)").Style = WordParagraphStyles.Heading1;
            WordTable wordTable14 = document.AddTable(4, 4, WordTableStyle.PlainTable1);
            wordTable14.Rows[0].Cells[0].Paragraphs[0].Text = "Fixed Content";
            Console.WriteLine($"Table 15 - Initial: {wordTable14.LayoutMode}");
            // Set Fixed Width via method
            wordTable14.SetFixedWidth(50);
            Console.WriteLine($"Table 15 - After SetFixedWidth(50): {wordTable14.LayoutMode}");
            // Set via new property (defaults to 100%)
            wordTable14.LayoutMode = WordTableLayoutType.FixedWidth;
            Console.WriteLine($"Table 15 - After LayoutMode=FixedWidth: {wordTable14.LayoutMode}");
            // Set back to 50%
            wordTable14.SetFixedWidth(50);
            Console.WriteLine($"Table 15 - After SetFixedWidth(50) again: {wordTable14.LayoutMode}");
            document.AddParagraph();

            document.AddParagraph("Table 16: Demonstrate Layout Changes (Window/Content)").Style = WordParagraphStyles.Heading1;
            WordTable wordTable15 = document.AddTable(4, 4, WordTableStyle.PlainTable1);
            wordTable15.Rows[0].Cells[0].Paragraphs[0].Text = "Content";
            wordTable15.Rows[1].Cells[1].Paragraphs[0].Text = "More Content";
            wordTable15.Rows[2].Cells[2].Paragraphs[0].Text = "Wider Content Here";
            wordTable15.Rows[0].Cells[1].Paragraphs[0].Text = "Test 1 - Long text that should not be cut off if AutoFit";
            Console.WriteLine($"Table 16 - Initial: {wordTable15.LayoutMode}");
            // Set via method
            wordTable15.AutoFitToContents();
            Console.WriteLine($"Table 16 - After AutoFitToContents(): {wordTable15.LayoutMode}");
            // Set via new property
            wordTable15.LayoutMode = WordTableLayoutType.AutoFitToWindow;
            Console.WriteLine($"Table 16 - After LayoutMode=AutoFitToWindow: {wordTable15.LayoutMode}");
            document.AddParagraph();


            document.AddParagraph("Table 17: Demonstrate Layout Changes (Window)").Style = WordParagraphStyles.Heading1;
            WordTable wordTable16 = document.AddTable(4, 4, WordTableStyle.PlainTable1);
            wordTable16.Rows[0].Cells[0].Paragraphs[0].Text = "Content";
            wordTable16.Rows[1].Cells[1].Paragraphs[0].Text = "More Content";
            wordTable16.Rows[2].Cells[2].Paragraphs[0].Text = "Wider Content Here";
            wordTable16.Rows[0].Cells[1].Paragraphs[0].Text = "Test 1 - Long text that should not be cut off if AutoFit";
            Console.WriteLine($"Table 17 - Initial: {wordTable16.LayoutMode}");
            // Set via method
            wordTable16.AutoFitToContents();
            Console.WriteLine($"Table 17 - After AutoFitToContents(): {wordTable16.LayoutMode}");
            // Set via new property
            document.AddParagraph();

            document.Save(openWord);
        }
    }
}