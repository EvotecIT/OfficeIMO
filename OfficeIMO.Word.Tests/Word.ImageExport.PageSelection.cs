using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class WordImageExportTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SinglePageBeyondTheEstimatedEndDoesNotReplayAnEarlierSection(bool multipleSections) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("FirstBody");
        document.AddPageBreak();
        document.AddParagraph("SecondBody");
        if (multipleSections) document.AddSection().AddParagraph("LastSectionBody");
        int count = document.GetEstimatedPageCount();
        var snapshot = document.CreateVisualSnapshot(new WordImageExportOptions { PageIndex = count });
        Assert.Contains(snapshot.Diagnostics, diagnostic => diagnostic.Code == WordImageExportDiagnosticCodes.UnsupportedPageIndex);
        Assert.DoesNotContain(snapshot.Drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text.Contains("Body"));
        Assert.DoesNotContain(snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>(), text => text.PlainText.Contains("Body"));
    }
}
