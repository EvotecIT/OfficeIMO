using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Tests configuring footnote and endnote properties.
    /// </summary>
    public partial class Word {
        /// <summary>
        /// Creates a document, sets footnote and endnote options, and reloads it.
        /// </summary>
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
                var section = document.Sections[0];
                var footProps = section.FootnoteProperties;
                Assert.NotNull(footProps);
                Assert.Equal(NumberFormatValues.LowerRoman, footProps!.NumberingFormat!.Val!.Value);
                Assert.Equal(FootnotePositionValues.PageBottom, footProps.FootnotePosition!.Val!.Value);
                Assert.Equal(RestartNumberValues.EachSection, footProps.NumberingRestart!.Val!.Value);
                Assert.Equal(5, footProps.NumberingStart!.Val!.Value);

                var endProps = section.EndnoteProperties;
                Assert.NotNull(endProps);
                Assert.Equal(NumberFormatValues.Decimal, endProps!.NumberingFormat!.Val!.Value);
                Assert.Equal(EndnotePositionValues.SectionEnd, endProps.EndnotePosition!.Val!.Value);
                Assert.Equal(RestartNumberValues.EachSection, endProps.NumberingRestart!.Val!.Value);
                Assert.Equal(5, endProps.NumberingStart!.Val!.Value);
            }
        }
    }
}
