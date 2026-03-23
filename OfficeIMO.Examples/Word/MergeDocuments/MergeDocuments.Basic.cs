using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class MergeDocuments {
        public static void Example_AppendDocument(string folderPath, bool openWord) {
            Console.WriteLine("[*] Appending one Word document into another");

            string outDir = Path.Combine(folderPath, "Word", "MergeDocuments");
            Directory.CreateDirectory(outDir);

            string destinationPath = Path.Combine(outDir, "MergeDestination.docx");
            string sourcePath = Path.Combine(outDir, "MergeSource.docx");
            string mergedPath = Path.Combine(outDir, "MergedDocument.docx");

            using (var destination = WordDocument.Create(destinationPath)) {
                var heading = destination.AddParagraph("Quarterly Report");
                heading.Style = WordParagraphStyles.Heading1;
                destination.AddParagraph("This file acts as the destination document.");

                var list = destination.AddList(WordListStyle.Numbered);
                list.AddItem("Overview");
                list.AddItem("Existing content");

                destination.Save();
            }

            using (var source = WordDocument.Create(sourcePath)) {
                var heading = source.AddParagraph("Appendix");
                heading.Style = WordParagraphStyles.Heading1;
                source.AddParagraph("This content comes from the second document.");

                var table = source.AddTable(2, 2, WordTableStyle.TableGrid);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Name";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Value";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "Alpha";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "42";

                source.Save();
            }

            using (var destination = WordDocument.Load(destinationPath))
            using (var source = WordDocument.Load(sourcePath)) {
                destination.AppendDocument(source);
                destination.SaveAs(mergedPath, openWord);
            }
        }
    }
}
