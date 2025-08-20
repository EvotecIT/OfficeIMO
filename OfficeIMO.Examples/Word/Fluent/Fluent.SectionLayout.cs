using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;

namespace OfficeIMO.Examples.Word {
    internal static partial class FluentDocument {
        public static void Example_FluentSectionLayout(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with section specific layout using fluent API");
            string filePath = Path.Combine(folderPath, "FluentSectionLayout.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .PageSetup(ps => ps.Orientation(PageOrientationValues.Landscape)
                                         .Size(WordPageSize.A4)
                                         .Margins(WordMargin.Normal)
                                         .DifferentFirstPage()
                                         .DifferentOddAndEvenPages())
                    .Section(s => s
                        .New()
                            .Margins(WordMargin.Narrow)
                            .Size(WordPageSize.Legal)
                            .Columns(2)
                            .PageNumbering(NumberFormatValues.LowerRoman, restart: true)
                            .Paragraph(p => p.Text("Section 1"))
                            .Table(t => t.Create(1, 1).Table!.Rows[0].Cells[0].AddParagraph("Cell 1"))
                        .New()
                            .Margins(WordMargin.Wide)
                            .Size(WordPageSize.A3)
                            .PageNumbering(restart: true)
                            .Paragraph(p => p.Text("Section 2"))
                            .Table(t => t.Create(1, 1).Table!.Rows[0].Cells[0].AddParagraph("Cell 2")))
                    .End()
                    .Save(false);
            }
            Helpers.Open(filePath, openWord);
        }
    }
}
