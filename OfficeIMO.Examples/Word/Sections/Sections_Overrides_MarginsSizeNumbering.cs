using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;

namespace OfficeIMO.Examples.Word {
    internal partial class Sections {

        internal static void Example_Word_Fluent_Sections_Overrides_MarginsSizeNumbering(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with multiple sections and per-section overrides");
            string filePath = Path.Combine(folderPath, "Fluent_Sections_Overrides_MarginsSizeNumbering.docx");

            using (var document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .PageSetup(ps => ps
                        .Orientation(PageOrientationValues.Portrait)
                        .Size(WordPageSize.A4)
                        .Margins(WordMargin.Normal))
                    .Section(sec => sec
                        .New(SectionMarkValues.NextPage)
                            .Margins(WordMargin.Narrow)
                            .Size(WordPageSize.Legal)
                            .PageNumbering(restart: true)
                            .Paragraph(p => p.Text("Section 1"))
                            .Table(t => t.Create(1, 1).Table!.Rows[0].Cells[0].AddParagraph("Cell 1"))
                        .New(SectionMarkValues.NextPage)
                            .Margins(WordMargin.Wide)
                            .Size(WordPageSize.A3)
                            .PageNumbering(restart: false)
                            .Paragraph(p => p.Text("Section 2"))
                            .Table(t => t.Create(1, 1).Table!.Rows[0].Cells[0].AddParagraph("Cell 2")))
                    .End()
                    .Save(false);
            }
            Helpers.Open(filePath, openWord);
        }
    }
}
