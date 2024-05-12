using DocumentFormat.OpenXml.Wordprocessing;

using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Tables {
        internal static void Example_BasicTables1(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with tables");
            string filePath = System.IO.Path.Combine(folderPath, "Document with Tables1.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Basic paragraph - Page 4");
                paragraph.ParagraphAlignment = JustificationValues.Center;

                document.AddParagraph();

                WordTable wordTable = document.AddTable(3, 4, WordTableStyle.PlainTable1);
                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                wordTable.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
                // align to center
                wordTable.Rows[2].Cells[3].Paragraphs[0].Text = "Center";
                wordTable.Rows[2].Cells[3].Paragraphs[0].ParagraphAlignment = JustificationValues.Center;

                // align to right
                wordTable.Rows[1].Cells[3].Paragraphs[0].Text = "Right";
                wordTable.Rows[1].Cells[3].Paragraphs[0].ParagraphAlignment = JustificationValues.Right;

                // align it on paragraph outside of table
                var paragraph1 = wordTable.Rows[0].Cells[0].Paragraphs[0].AddParagraph();
                paragraph1 = paragraph1.AddParagraph();
                paragraph1.AddText("Ok");
                paragraph1.ParagraphAlignment = JustificationValues.Center;

                var paragraph2 = wordTable.Rows[1].Cells[0].Paragraphs[0].AddParagraphAfterSelf();
                paragraph2 = paragraph2.AddParagraphAfterSelf();
                paragraph2.AddText("Ok2");

                var paragraphBefore = wordTable.Rows[1].Cells[0].Paragraphs[0].AddParagraphBeforeSelf();
                paragraphBefore = paragraphBefore.AddParagraphBeforeSelf();
                paragraphBefore.AddText("Ok but Before");


                wordTable.Rows[2].Cells[0].Paragraphs[0].AddParagraphAfterSelf().AddParagraphAfterSelf().AddParagraphAfterSelf().Text = "Works differently";



                Console.WriteLine(wordTable.Style);

                // lets overwrite style
                wordTable.Style = WordTableStyle.GridTable6ColorfulAccent1;

                document.Save(openWord);
            }
        }
    }
}
