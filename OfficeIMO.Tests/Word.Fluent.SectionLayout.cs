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
                        .New(SectionMarkValues.NextPage)
                            .Margins(WordMargin.Narrow)
                            .Size(WordPageSize.Legal)
                            .PageNumbering(restart: true)
                            .Paragraph(p => p.Text("Section 1"))
                            .Table(t => t.Columns(1).Row("Cell 1"))
                        .New(SectionMarkValues.NextPage)
                            .Margins(WordMargin.Wide)
                            .Size(WordPageSize.A3)
                            .PageNumbering(restart: false)
                            .Paragraph(p => p.Text("Section 2"))
                            .Table(t => t.Columns(1).Row("Cell 2")))
                    .End();
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {

                Assert.Equal(3, document.Sections.Count);
                Assert.Equal("Section 1", document.Sections[1].Paragraphs[0].Text);
                Assert.Equal("Cell 1", document.Sections[1].Tables[0].Rows[0].Cells[0].Paragraphs[0].Text);
                Assert.Equal("Section 2", document.Sections[2].Paragraphs[0].Text);
                Assert.Equal("Cell 2", document.Sections[2].Tables[0].Rows[0].Cells[0].Paragraphs[0].Text);
            }
        }
    }
}
