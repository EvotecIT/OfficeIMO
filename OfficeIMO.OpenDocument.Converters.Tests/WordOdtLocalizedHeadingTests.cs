using OfficeIMO.OpenDocument;
using OfficeIMO.Word;
using OfficeIMO.Word.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class WordOdtLocalizedHeadingTests {
    [Theory]
    [InlineData("Titre2")]
    [InlineData("berschrift2")]
    public void LocalizedHeadingStyleBecomesOdtHeading(string localizedHeading2) {
        using WordDocument source = WordDocument.Create();
        source.Styles.Remove("Heading2");
        source.Styles.AddParagraphStyle(localizedHeading2, "heading 2");
        source.AddParagraph("Localized heading").SetStyleId(localizedHeading2);

        OdtDocument odt = source.ToOpenDocument();

        OdtParagraph heading = Assert.Single(odt.ContentBlocks, block => block.Paragraph?.Text == "Localized heading").Paragraph!;
        Assert.True(heading.IsHeading);
        Assert.Equal(2, heading.HeadingLevel);
    }
}
