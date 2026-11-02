using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordContentBlockTests {
    [Fact]
    public void NativeBlocksRetainTasksLinksStylesAndThemedCallouts() {
        using var document = WordDocument.Create();
        document.AddParagraph("Existing content");
        var toc = document.AddTableOfContent(title: "Contents", minLevel: 1, maxLevel: 2, insertAtTop: true);
        Assert.Equal("Contents", toc.Text);
        var callout = document.AddCallout(OfficeCalloutKind.Warning, "Attention", "Check the totals", OfficeTheme.Dark);
        var code = document.AddCodeBlock("return 1;", "csharp");
        var list = document.AddList(WordListStyle.Bulleted);
        list.AddTaskItem("Review", done: true);
        list.AddLinkItem("Source", new Uri("https://example.com/source"));

        using var stream = new MemoryStream(document.ToBytes());
        using var package = WordprocessingDocument.Open(stream, false);
        var main = package.MainDocumentPart!;
        Assert.IsType<SdtBlock>(main.Document!.Body!.FirstChild);
        Assert.Equal(OfficeTheme.Dark.Palette.Warning.ToRgbHex(), callout.ShadingFillColor!.Value.ToRgbHex());
        Assert.Equal("OfficeIMOCodeBlock", code.StyleId);
        Assert.Contains(main.StyleDefinitionsPart!.Styles!.Elements<Style>(), style => style.StyleId == code.StyleId && style.StyleRunProperties?.RunFonts != null);
        Assert.Single(main.Document!.Descendants<DocumentFormat.OpenXml.Office2010.Word.SdtContentCheckBox>());
        Assert.Contains(main.HyperlinkRelationships, relationship => relationship.Uri.AbsoluteUri == "https://example.com/source");
    }
}
