using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;

namespace OfficeIMO.Examples.Word {
    internal static partial class FluentDocument {
        public static void Example_FluentHeadersAndFooters(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with fluent headers and footers");
            string filePath = Path.Combine(folderPath, "FluentHeadersAndFooters.docx");
            string imagesPath = Path.Combine(Directory.GetCurrentDirectory(), "Images");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .PageSetup(p => p.DifferentFirstPage().DifferentOddAndEvenPages())
                    .Header(h => h
                        .Default(d => d
                            .Paragraph("Default header")
                            .Paragraph(pb => pb.Text("Second header paragraph"))
                            .Image(Path.Combine(imagesPath, "Kulek.jpg"), 50, 50)
                            .Table(2, 2))
                        .First(f => f.Paragraph("First page header"))
                        .Even(e => e.Paragraph("Even page header")))
                    .Footer(f => f
                        .Default(d => d.Paragraph("Default footer"))
                        .First(ft => ft.Paragraph("First page footer"))
                        .Even(ev => ev.Paragraph("Even page footer")))
                    .Paragraph(p => p.Text("Body paragraph"))
                    .Section(s => s.New())
                    .Paragraph(p => p.Text("Second section paragraph"))
                    .End()
                    .Save(false);
            }

            Helpers.Open(filePath, openWord);
        }
    }
}

