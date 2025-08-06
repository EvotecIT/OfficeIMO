using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Shared {
    internal static partial class Shared {
        public static void Example_SharedHelpers(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "SharedHelpers.docx");

            using (var document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph(string.Empty);
                paragraph.AddFormattedText("Hello", bold: true);
                paragraph.AddFormattedText("World", italic: true);
                document.Save();

                foreach (var section in DocumentTraversal.EnumerateSections(document)) {
                    foreach (var p in section.Paragraphs) {
                        foreach (var run in FormattingHelper.GetFormattedRuns(p)) {
                            if (!string.IsNullOrEmpty(run.Text)) {
                                Console.WriteLine($"{run.Text} B:{run.Bold} I:{run.Italic}");
                            }
                        }
                    }
                }
            }

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}

