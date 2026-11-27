using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Theory]
        [InlineData(WordSectionBreakType.NextPage, "nextPage")]
        [InlineData(WordSectionBreakType.NextColumn, "nextColumn")]
        [InlineData(WordSectionBreakType.Continuous, "continuous")]
        [InlineData(WordSectionBreakType.EvenPage, "evenPage")]
        [InlineData(WordSectionBreakType.OddPage, "oddPage")]
        public void Test_AddSectionMapsOfficeIMOSectionBreakTypes(WordSectionBreakType breakType, string expected) {
            string filePath = Path.Combine(_directoryWithFiles, $"SectionBreak{expected}.docx");

            using WordDocument document = WordDocument.Create(filePath);
            WordSection previousSection = document.Sections[0];
            WordSection newSection = document.AddSection(breakType);

            Assert.Equal(2, document.Sections.Count);
            Assert.Equal(WordSectionBreakType.NextPage, previousSection.BreakType);
            Assert.Equal(breakType, newSection.BreakType);
            SectionType sectionType = Assert.IsType<SectionType>(newSection._sectionProperties.GetFirstChild<SectionType>());
            Assert.Equal(expected, sectionType.Val?.InnerText);
        }

        [Fact]
        public void Test_AddSectionDoesNotExposeOpenXmlValueStructOverload() {
            var openXmlOverload = typeof(WordDocument).GetMethod(nameof(WordDocument.AddSection), new[] { typeof(SectionMarkValues?) });
            Assert.Null(openXmlOverload);
        }

        [Fact]
        public void Test_LegacyDefaultLiteralCallsRemainSourceCompatible() {
            string filePath = Path.Combine(_directoryWithFiles, "LegacyDefaultLiteralSectionBreaks.docx");

            using WordDocument document = WordDocument.Create(filePath);
            document.AddSection(default);
            document.AddSection(default);

            Assert.Equal(3, document.Sections.Count);
        }
    }

    internal static class SectionBreakTestExtensions {
        // Legacy DOC fixtures still describe section breaks using the SDK's wire values.
        internal static WordSectionBreakType ToOfficeEnum(this SectionMarkValues value) {
            if (value == SectionMarkValues.NextPage) return WordSectionBreakType.NextPage;
            if (value == SectionMarkValues.NextColumn) return WordSectionBreakType.NextColumn;
            if (value == SectionMarkValues.Continuous) return WordSectionBreakType.Continuous;
            if (value == SectionMarkValues.EvenPage) return WordSectionBreakType.EvenPage;
            if (value == SectionMarkValues.OddPage) return WordSectionBreakType.OddPage;
            throw new ArgumentOutOfRangeException(nameof(value), value, "Unsupported Word section break type.");
        }
    }
}
