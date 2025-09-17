using System;
using System.IO;
using DocumentFormat.OpenXml.Bibliography;
using OfficeIMO.Examples.Utils;
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

                string bookTag = Guard.NotNullOrWhiteSpace(book.Tag, "Bibliography source 'book' must define a tag.");
                string articleTag = Guard.NotNullOrWhiteSpace(article.Tag, "Bibliography source 'article' must define a tag.");

                document.BibliographySources[bookTag] = book;
                document.BibliographySources[articleTag] = article;

                document.AddParagraph("Book reference: ").AddCitation(bookTag);
                document.AddParagraph("Article reference: ").AddCitation(articleTag);

                document.Save(openWord);
            }
        }
    }
}
