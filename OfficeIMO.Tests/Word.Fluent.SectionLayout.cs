using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
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
                Assert.Equal("Section 1", document.Sections[1].Paragraphs[0].Text);
                Assert.Equal("Cell 1", document.Sections[1].Tables[0].Rows[0].Cells[0].Paragraphs[0].Text);
                Assert.Equal("Section 2", document.Sections[2].Paragraphs[0].Text);
                Assert.Equal("Cell 2", document.Sections[2].Tables[0].Rows[0].Cells[0].Paragraphs[0].Text);

                Assert.Equal(NumberFormatValues.LowerRoman, document.Sections[1].PageNumberType.Format!.Value);
                Assert.Equal(1, document.Sections[1].PageNumberType.Start!.Value);

                Assert.Null(document.Sections[2].PageNumberType.Format);
                Assert.Equal(1, document.Sections[2].PageNumberType.Start!.Value);
            }
        }
    }
}
