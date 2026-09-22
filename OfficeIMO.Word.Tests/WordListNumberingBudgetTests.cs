using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Word.Tests;

public sealed class WordListNumberingBudgetTests {
    [Fact]
    public void ConversionRejectsExcessiveStyleInheritance() {
        using WordDocument document = WordDocument.Create();
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        for (int index = 0; index < 257; index++) {
            styles.Append(new Style(new BasedOn { Val = index == 256 ? "Normal" : "Chain" + (index + 1) }) {
                Type = StyleValues.Paragraph,
                StyleId = "Chain" + index,
                CustomStyle = true
            });
        }
        document.AddParagraph("Plain text").SetStyleId("Chain0");

        Assert.Throws<InvalidDataException>(() => WordDocumentTraversal.BuildResolvedListMarkers(document));
    }

    [Fact]
    public void ExportCatalogCachesMissingNumberingForRepeatedStyles() {
        using WordDocument document = WordDocument.Create();
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new Style(new BasedOn { Val = "Normal" }) {
            Type = StyleValues.Paragraph,
            StyleId = "PlainChain",
            CustomStyle = true
        });
        WordParagraph first = document.AddParagraph("First").SetStyleId("PlainChain");
        WordParagraph second = document.AddParagraph("Second").SetStyleId("PlainChain");
        WordListNumberingResolver.StyleCatalog catalog = WordListNumberingResolver.CreateStyleCatalog(document);

        Assert.False(WordListNumberingResolver.TryResolve(first, out _, catalog));
        Assert.False(WordListNumberingResolver.TryResolve(second, out _, catalog));
        Assert.Single(catalog.NumberingByStyle!);
        Assert.Single(catalog.LinkedStyleByStyle!);
    }
}
