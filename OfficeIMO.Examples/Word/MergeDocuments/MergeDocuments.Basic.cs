using OfficeIMO.Word.Pdf;
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
                destination.AddTableOfContent();
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
                destination.UpdateFields(new WordFieldUpdateOptions {
                    UpdateTablesOfContents = true
                }.WithPdfLayoutPagination());
                destination.SaveCopy(mergedPath);
                if (openWord) destination.OpenInApplication(mergedPath);
            }
        }

        public static void Example_SplitDocument(string folderPath, bool openWord) {
            Console.WriteLine("[*] Splitting a Word document by heading");

            string outDir = Path.Combine(folderPath, "Word", "MergeDocuments");
            Directory.CreateDirectory(outDir);

            using var report = WordDocument.Create();
            foreach (string chapter in new[] { "Summary", "Findings", "Next steps" }) {
                report.AddParagraph(chapter).Style = WordParagraphStyles.Heading1;
                report.AddParagraph($"{chapter} text.").AddFootNote($"Source for {chapter.ToLowerInvariant()}.");
            }

            IReadOnlyList<WordDocument> pieces = report.SplitByHeading(level: 1);
            for (int i = 0; i < pieces.Count; i++) {
                using WordDocument piece = pieces[i];
                string path = Path.Combine(outDir, $"SplitChapter{i + 1}.docx");
                piece.Save(path);
                if (openWord && i == 0) piece.OpenInApplication(path);
            }
        }
    }
}
