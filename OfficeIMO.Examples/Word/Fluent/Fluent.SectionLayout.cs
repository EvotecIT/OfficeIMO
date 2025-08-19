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
                    .Page(p => p.SetOrientation(PageOrientationValues.Landscape)
                                 .SetPaperSize(WordPageSize.A4)
                                 .SetMargins(WordMargin.Normal)
                                 .DifferentFirstPage()
                                 .DifferentOddAndEvenPages())
                    .Section(s => {
                        s.AddSection(SectionMarkValues.NextPage)
                            .SetMargins(WordMargin.Narrow)
                            .SetPageSize(WordPageSize.Legal)
                            .SetPageNumbering(restart: true);

                        s.AddSection(SectionMarkValues.NextPage)
                            .SetMargins(WordMargin.Wide)
                            .SetPageSize(WordPageSize.A3)
                            .SetPageNumbering(restart: false);
                    });
                document.Save(false);
            }
            Helpers.Open(filePath, openWord);
        }
    }
}
