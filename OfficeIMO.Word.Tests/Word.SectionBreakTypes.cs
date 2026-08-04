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
            document.AddSection(breakType);

            Assert.Equal(2, document.Sections.Count);
            SectionType sectionType = Assert.IsType<SectionType>(previousSection._sectionProperties.GetFirstChild<SectionType>());
            Assert.Equal(expected, sectionType.Val?.InnerText);
        }

        [Fact]
        public void Test_FluentSectionBuilderAcceptsOfficeIMOSectionBreakType() {
            string filePath = Path.Combine(_directoryWithFiles, "FluentContinuousSectionBreak.docx");

            using WordDocument document = WordDocument.Create(filePath);
            WordSection previousSection = document.Sections[0];
            document.AsFluent().Section(section => section.New(WordSectionBreakType.Continuous)).End();

            Assert.Equal(2, document.Sections.Count);
            SectionType sectionType = Assert.IsType<SectionType>(previousSection._sectionProperties.GetFirstChild<SectionType>());
            Assert.Equal(SectionMarkValues.Continuous, sectionType.Val?.Value);
        }

        [Fact]
        public void Test_AddSectionRetainsNullableOpenXmlOverload() {
            string filePath = Path.Combine(_directoryWithFiles, "NullableOpenXmlSectionBreak.docx");
            SectionMarkValues? breakType = SectionMarkValues.Continuous;

            var legacyOverload = typeof(WordDocument).GetMethod(nameof(WordDocument.AddSection), new[] { typeof(SectionMarkValues?) });
            Assert.NotNull(legacyOverload);
            var legacyParameter = Assert.Single(legacyOverload!.GetParameters());
            Assert.True(legacyParameter.IsOptional);
            Assert.Null(legacyParameter.DefaultValue);

            using WordDocument document = WordDocument.Create(filePath);
            WordSection previousSection = document.Sections[0];
            document.AddSection(breakType);

            SectionType sectionType = Assert.IsType<SectionType>(previousSection._sectionProperties.GetFirstChild<SectionType>());
            Assert.Equal(SectionMarkValues.Continuous, sectionType.Val?.Value);
        }
    }
}
