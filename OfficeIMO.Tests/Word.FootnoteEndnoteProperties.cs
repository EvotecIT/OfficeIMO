using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_FootnoteEndnotePropertiesRoundtrip() {
            string filePath = Path.Combine(_directoryWithFiles, "FootnoteEndnoteProperties.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddFootnoteProperties(NumberFormatValues.LowerRoman,
                                            FootnotePositionValues.PageBottom,
                                            RestartNumberValues.EachSection,
                                            5);
                document.AddEndnoteProperties(NumberFormatValues.Decimal,
                                            EndnotePositionValues.SectionEnd,
                                            RestartNumberValues.EachSection,
                                            5);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal(NumberFormatValues.LowerRoman, document.Sections[0].FootnoteProperties.NumberingFormat.Val.Value);
                Assert.Equal(FootnotePositionValues.PageBottom, document.Sections[0].FootnoteProperties.FootnotePosition.Val.Value);
                Assert.Equal(RestartNumberValues.EachSection, document.Sections[0].FootnoteProperties.NumberingRestart.Val.Value);
                Assert.Equal(5, document.Sections[0].FootnoteProperties.NumberingStart.Val.Value);

                Assert.Equal(NumberFormatValues.Decimal, document.Sections[0].EndnoteProperties.NumberingFormat.Val.Value);
                Assert.Equal(EndnotePositionValues.SectionEnd, document.Sections[0].EndnoteProperties.EndnotePosition.Val.Value);
                Assert.Equal(RestartNumberValues.EachSection, document.Sections[0].EndnoteProperties.NumberingRestart.Val.Value);
                Assert.Equal(5, document.Sections[0].EndnoteProperties.NumberingStart.Val.Value);
            }
        }
    }
}
