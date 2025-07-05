using System;
using System.IO;
using DocumentFormat.OpenXml.Bibliography;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class CitationsExamples {
        public static void Example_BasicCitations(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with a single citation");
            string filePath = Path.Combine(folderPath, "DocumentWithCitation.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var source = new WordBibliographySource("Src1", DataSourceValues.Book) {
                    Title = "Sample Book",
                    Author = "John Doe",
                    Year = "2024"
                };
                document.BibliographySources[source.Tag] = source;

                document.AddParagraph("Referenced text: ").AddCitation(source.Tag);
                document.Save(openWord);
            }
        }
    }
}
