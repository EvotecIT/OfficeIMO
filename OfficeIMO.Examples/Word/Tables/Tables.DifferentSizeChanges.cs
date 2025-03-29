using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Tables {
        internal static void Example_DifferentTableSizes(string folderPath, bool openWord) {

            Console.WriteLine("[*] Creating standard document with tables of different sizes");
            string filePath = System.IO.Path.Combine(folderPath, "Document with Tables of different sizes.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Table 1");
                WordTable wordTable = document.AddTable(3, 4, WordTableStyle.PlainTable1);
                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";

                document.AddParagraph();
                document.AddParagraph();
                document.AddParagraph("Table 2 - Sized for 2000 width / Centered");
                WordTable wordTable1 = document.AddTable(2, 6, WordTableStyle.PlainTable1);
                wordTable1.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable1.Rows[1].Cells[0].Paragraphs[0].Text = "Test 1 - ok longer text, no autosize right?";
                wordTable1.WidthType = TableWidthUnitValues.Pct;
                wordTable1.Width = 100;
                wordTable1.Alignment = TableRowAlignmentValues.Center;

                document.AddParagraph();
                document.AddParagraph();
                document.AddParagraph("Table 3 - By default the table is autosized for full width");
                WordTable wordTable2 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
                wordTable2.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";

                document.AddParagraph();
                document.AddParagraph();
                document.AddParagraph("Table 4 - Magic number 5000 (full width)");
                WordTable wordTable3 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
                wordTable3.WidthType = TableWidthUnitValues.Pct;
                wordTable3.Width = 5000;
                wordTable3.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";

                document.AddParagraph();
                document.AddParagraph();
                document.AddParagraph("Table 5 - 50% by using 2500 width (pct)");
                WordTable wordTable4 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
                wordTable4.WidthType = TableWidthUnitValues.Pct;
                wordTable4.Width = 2500;
                wordTable4.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";

                document.AddParagraph();
                document.AddParagraph();
                document.AddParagraph("Table 6 - 50% by using 2500 width (pct), that we fix to full width");
                WordTable wordTable5 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
                // this data is temporary just to prove things work
                wordTable5.WidthType = TableWidthUnitValues.Pct;
                wordTable5.Width = 2500;
                // lets fix it for full width
                wordTable5.DistributeColumnsEvenly();

                document.AddParagraph();
                document.AddParagraph();
                document.AddParagraph("Table 7 - 50%");
                WordTable wordTable6 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
                wordTable6.SetWidthPercentage(50);

                document.AddParagraph();
                document.AddParagraph();
                document.AddParagraph("Table 8 - 75%");
                WordTable wordTable7 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
                wordTable7.SetWidthPercentage(75);

                document.AddParagraph();
                document.AddParagraph();
                document.AddParagraph("Table 9");
                WordTable wordTable8 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
                wordTable8.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable8.Rows[0].Cells[0].Width = 1000; // this will not work alone i think
                wordTable8.Rows[0].Cells[1].Width = 500; // this will not work alone i think
                wordTable8.ColumnWidth = new List<int>() { 1000, 500, 500, 750 };
                wordTable8.ColumnWidthType = TableWidthUnitValues.Pct;
                wordTable8.SetWidthPercentage(100);


                document.AddParagraph();
                document.AddParagraph();
                document.AddParagraph("Table 10 - Shows setting up different sizes for each column");
                WordTable wordTable9 = document.AddTable(3, 4, WordTableStyle.PlainTable1);

                wordTable9.ColumnWidth = new List<int>() { 1000, 500, 500, 750 };
                wordTable9.ColumnWidthType = TableWidthUnitValues.Pct;
                wordTable9.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";


                document.AddParagraph();
                document.AddParagraph();
                document.AddParagraph("Table 11 - Shows setting up different sizes for each column, but fixing it with Distribute Columns Evenly");
                WordTable wordTable10 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
                wordTable10.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable10.ColumnWidth = new List<int>() { 1000, 500, 500, 750 };
                wordTable10.ColumnWidthType = TableWidthUnitValues.Pct;
                // Lets distribute it evenly now
                wordTable10.DistributeColumnsEvenly();


                document.AddParagraph();
                document.AddParagraph();
                document.AddParagraph("Table 12 - The same as above, but manually");
                WordTable wordTable11 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
                wordTable11.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable11.ColumnWidth = new List<int>() { 687, 687, 687, 687 };
                wordTable11.ColumnWidthType = TableWidthUnitValues.Pct;
                wordTable11.Width = 2063;
                wordTable11.WidthType = TableWidthUnitValues.Pct;


                document.AddParagraph();
                document.AddParagraph();
                document.AddParagraph("Table 13 - Set the magic number by column width");
                WordTable wordTable12 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
                wordTable12.ColumnWidth = new List<int>() { 1250, 1250, 1250, 1250 };
                wordTable12.ColumnWidthType = TableWidthUnitValues.Pct;
                wordTable12.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";


                document.AddParagraph();
                document.AddParagraph();
                document.AddParagraph("Table 14");

                WordTable wordTable13 = document.AddTable(4, 4, WordTableStyle.PlainTable1);
                wordTable13.LayoutType = TableLayoutValues.Autofit;
                wordTable13.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable13.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                wordTable13.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
                wordTable13.Rows[3].Cells[0].Paragraphs[0].Text = "Test 4";

                document.AddParagraph();
                document.AddParagraph();
                document.AddParagraph("Table 15");

                WordTable wordTable14 = document.AddTable(4, 4, WordTableStyle.PlainTable1);
                wordTable14.LayoutType = TableLayoutValues.Fixed;
                wordTable14.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable14.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                wordTable14.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
                wordTable14.Rows[3].Cells[0].Paragraphs[0].Text = "Test 4";

                document.AddParagraph();
                document.AddParagraph();
                document.AddParagraph("Table 16");

                WordTable wordTable15 = document.AddTable(4, 4, WordTableStyle.PlainTable1);
                wordTable15.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable15.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                wordTable15.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
                wordTable15.Rows[3].Cells[0].Paragraphs[0].Text = "Test 4";

                wordTable15.Rows[0].Cells[1].Paragraphs[0].Text = "Test 1 - Long text that should not be cut off";
                wordTable15.Rows[1].Cells[2].Paragraphs[0].Text = "Test 2";
                wordTable15.Rows[2].Cells[2].Paragraphs[0].Text = "Test 3";
                wordTable15.Rows[3].Cells[3].Paragraphs[0].Text = "Test 4 - longer";

                var layoutType = wordTable15.GetCurrentLayoutType();
                Console.WriteLine("Current Layout Type: " + layoutType.ToString());

                wordTable15.AutoFitToContents();

                var layoutType2 = wordTable15.GetCurrentLayoutType();
                Console.WriteLine("Current Layout Type after change: " + layoutType2.ToString());

                document.Save(openWord);
            }
        }
    }
}
