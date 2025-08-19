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

            using (WordDocument document = WordDocument.Load(filePath)) {

                Assert.True(document.Sections.Count >= 1);
            }
        }
    }
}
