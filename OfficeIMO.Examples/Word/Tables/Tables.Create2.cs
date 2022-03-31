using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Examples.Word {
    internal static partial class Tables {
        internal static void Example_Tables(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with tables");
            string filePath = System.IO.Path.Combine(folderPath, "Document with Tables.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Basic paragraph - Page 4");
                paragraph.ParagraphAlignment = JustificationValues.Center;

                WordTable wordTable = document.AddTable(3, 4, WordTableStyle.PlainTable1);
                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                wordTable.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";

                Console.WriteLine(wordTable.Style);
                Console.WriteLine(wordTable.Rows.Count);

                wordTable.Rows[1].Remove();

                Console.WriteLine(wordTable.Rows.Count);
                wordTable.Rows[1].Cells[1].Paragraphs[0].Text = "This should be in row 1st";
                wordTable.Rows[1].Cells[2].Paragraphs[0].Text = "This should be in row 1st - 2nd column";
                wordTable.Rows[1].Cells[3].Paragraphs[0].Text = "This should be in row 1st - 3rd column";
                wordTable.Rows[1].Cells[2].Remove();
                wordTable.Rows[1].Cells[2].Paragraphs[0].AddText("More text which means another paragraph 1");
                wordTable.Rows[1].Cells[2].Paragraphs[0].AddText("More text which means another paragraph 2");

                Console.WriteLine(wordTable.Rows[1].Cells[2].Paragraphs.Count);

                Console.WriteLine(wordTable.Rows.Count);
                wordTable.AddRow();
                wordTable.AddRow(7);
                wordTable.AddRow();
                wordTable.AddRow(5, 5);
                Console.WriteLine(wordTable.Rows.Count);

                wordTable.Rows[8].Cells[1].Paragraphs[0].Text = "This should be in row 8th";
                wordTable.Rows[1].Cells[2].Paragraphs[2].Text = "Change me";
                wordTable.Rows[1].Cells[2].Paragraphs[2].SetColor(SixLabors.ImageSharp.Color.Green);
                // lets overwrite style
                wordTable.Style = WordTableStyle.GridTable6ColorfulAccent1;

                Console.WriteLine("----");
                Console.WriteLine(document.Tables.Count);

                WordTable wordTable1 = document.AddTable(3, 4, WordTableStyle.GridTable5DarkAccent5);

                Console.WriteLine(document.Tables.Count);

                WordTable wordTable2 = document.AddTable(3, 4, WordTableStyle.GridTable5DarkAccent5);
                wordTable2.Remove();

                Console.WriteLine(document.Tables.Count);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.Tables[1].Remove();

                document.AddParagraph("This new table should have cells merged");

                WordTable wordTable = document.AddTable(3, 4, WordTableStyle.PlainTable1);

                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Some test";
                wordTable.Rows[0].Cells[1].Paragraphs[0].Text = "Some test 1";
                wordTable.Rows[0].Cells[2].Paragraphs[0].Text = "Some test 2";
                wordTable.Rows[0].Cells[3].Paragraphs[0].Text = "Some test 3";
                wordTable.Rows[0].Cells[1].MergeHorizontally(2, true);
                // we unmerge the cells
                //wordTable.Rows[0].Cells[2].HorizontalMerge = null;
                //wordTable.Rows[0].Cells[3].HorizontalMerge = null;
                // bring back from merge
                wordTable.Rows[0].Cells[1].SplitHorizontally(2);


                Console.WriteLine(document.Tables.Count);

                document.AddParagraph("Another table");

                wordTable = document.AddTable(7, 4, WordTableStyle.PlainTable1);

                wordTable.Rows[0].Cells[2].Paragraphs[0].Text = "Some test 0";
                wordTable.Rows[1].Cells[2].Paragraphs[0].Text = "Some test 1";
                wordTable.Rows[2].Cells[2].Paragraphs[0].Text = "Some test 2";
                wordTable.Rows[3].Cells[2].Paragraphs[0].Text = "Some test 3";
                wordTable.Rows[0].Cells[2].MergeVertically(2, true);


                document.AddHorizontalLine(BorderValues.Double, Color.Green);


                document.AddParagraph("Test");


                var paragraph = document.AddParagraph().AddHorizontalLine();

                document.AddPageBreak();

                var section = document.AddSection();

                section.AddParagraph("This is a big test");

                section.AddHorizontalLine(BorderValues.BalloonsHotAir, null, 24, 24);

                document.Save(openWord);
            }
        }
    }
}
