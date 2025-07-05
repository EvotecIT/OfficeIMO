using System;
using System.IO;
using DocumentFormat.OpenXml.Bibliography;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class CitationsExamples {
        public static void Example_AdvancedCitations(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with multiple citations");
            string filePath = Path.Combine(folderPath, "DocumentWithMultipleCitations.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var book = new WordBibliographySource("B1", DataSourceValues.Book) {
                    Title = "C# in Depth",
                    Author = "Jon Skeet",
                    Year = "2021"
                };
                var article = new WordBibliographySource("A1", DataSourceValues.ArticleInAPeriodical) {
                    Title = "Wordprocessing With OpenXML",
                    Author = "Jane Smith",
                    Year = "2023"
                };

                document.BibliographySources[book.Tag] = book;
                document.BibliographySources[article.Tag] = article;

                document.AddParagraph("Book reference: ").AddCitation(book.Tag);
                document.AddParagraph("Article reference: ").AddCitation(article.Tag);

                document.Save(openWord);
            }
        }
    }
}
