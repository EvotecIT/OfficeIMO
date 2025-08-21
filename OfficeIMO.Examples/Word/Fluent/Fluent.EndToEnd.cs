using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;

namespace OfficeIMO.Examples.Word {
    internal static partial class FluentDocument {
        public static void Example_FluentEndToEnd(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating end-to-end fluent document");
            string filePath = Path.Combine(folderPath, "FluentEndToEnd.docx");
            string imagesPath = Path.Combine(Directory.GetCurrentDirectory(), "Images");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Info(i => i.Title("Quarterly Review").Author("OfficeIMO"))
                    .PageSetup(ps => ps.Orientation(PageOrientationValues.Portrait)
                                        .Size(WordPageSize.A4)
                                        .Margins(WordMargin.Normal))
                    .Paragraph(p => p.Text("Hello ").Text("World", t => t.BoldOn().Color("#ff0000")).Text("!"))
                    .List(l => l.Numbered().Item("First").Item("Second"))
                    .Table(t => t.Columns(3).PreferredWidth(percent: 100)
                        .Header("Name", "Role", "Score")
                        .Row("Alice", "Dev", 98)
                        .Row("Bob", "Ops", 91)
                        .Style(WordTableStyle.TableGrid)
                        .Align(HorizontalAlignment.Center))
                    .Image(i => i.Add(Path.Combine(imagesPath, "Kulek.jpg")).Size(100).Alt("Chart", "Quarterly chart"))
                    .End()
                    .Save(false);
            }
            Helpers.Open(filePath, openWord);
        }
    }
}
