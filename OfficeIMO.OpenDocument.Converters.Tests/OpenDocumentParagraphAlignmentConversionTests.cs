using OfficeIMO.OpenDocument;
using OfficeIMO.Word;
using OfficeIMO.Word.OpenDocument;
using System.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class OpenDocumentParagraphAlignmentConversionTests {
    [Fact]
    public void ExistingAlignmentEnumValuesRemainBinaryCompatible() {
        Assert.Equal(0, (int)OdtParagraphAlignment.Start);
        Assert.Equal(1, (int)OdtParagraphAlignment.Center);
        Assert.Equal(2, (int)OdtParagraphAlignment.End);
        Assert.Equal(3, (int)OdtParagraphAlignment.Justify);
    }

    [Theory]
    [InlineData(OdtParagraphAlignment.Left, "Left")]
    [InlineData(OdtParagraphAlignment.Right, "Right")]
    public void OdtPhysicalAlignmentRemainsPhysicalInRightToLeftParagraphs(
        OdtParagraphAlignment alignment,
        string expectedWordAlignment) {
        OdtDocument source = OdtDocument.Create();
        OdtParagraph paragraph = source.AddParagraph("Physical alignment");
        paragraph.Alignment = alignment;
        paragraph.WritingMode = "rl-tb";

        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument target = conversion.Value;
        WordParagraphSnapshot converted = Assert.Single(target.CreateInspectionSnapshot().Sections
            .SelectMany(section => section.Elements).OfType<WordParagraphSnapshot>());

        Assert.Equal(expectedWordAlignment, converted.Alignment);
        Assert.True(converted.IsRightToLeft);
        Assert.DoesNotContain(conversion.Report.Mappings, mapping =>
            mapping.Feature == "paragraph-formatting" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);

        OdtParagraph roundTrip = Assert.Single(target.ToOpenDocumentResult().Value.Paragraphs);
        Assert.Equal(alignment, roundTrip.Alignment);
        Assert.True(roundTrip.IsRightToLeft);
    }
}
