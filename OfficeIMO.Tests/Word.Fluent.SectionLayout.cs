using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact(Skip = "Fluent section layout pending fix")]
        public void Test_FluentSectionLayout() {
            string filePath = Path.Combine(_directoryWithFiles, "FluentSectionLayout.docx");
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
                            .Table(t => t.Columns(1).Row("Cell 1"))
                        .New()
                            .Margins(WordMargin.Wide)
                            .Size(WordPageSize.A3)
                            .PageNumbering(restart: true)
                            .Paragraph(p => p.Text("Section 2"))
                            .Table(t => t.Columns(1).Row("Cell 2")))
                    .End();

                Assert.Equal(3, document.Sections.Count);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal(3, document.Sections.Count);
            }
        }
    }
}
