using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class HtmlWordPaperPresetTests {
    [Theory]
    [InlineData(WordPageSize.Envelope10, OfficePageOrientation.Landscape, 500, false)]
    [InlineData(WordPageSize.Tabloid, OfficePageOrientation.Portrait, 1200, true)]
    public void EditableRegionUsesTheConfiguredWordPresetHeight(WordPageSize preset, OfficePageOrientation orientation,
        int topPixels, bool positioned) {
        string html = "<div style='position:absolute;top:" + topPixels
            + "px;width:120px;height:30px'>Preset region</div>";
        HtmlToWordResult result = HtmlConversionDocument.Parse(html).ToWordDocumentResult(new HtmlToWordOptions {
            DefaultPageSize = preset, DefaultOrientation = orientation
        });
        using WordDocument word = result.Value;
        Assert.Equal(preset, word.Sections[0].PageSettings.PageSize);
        if (positioned) {
            Assert.Single(word.TextBoxes);
        } else {
            Assert.Empty(word.TextBoxes);
            Assert.NotEmpty(word.Find("Preset region", StringComparison.Ordinal));
            Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
                && diagnostic.Detail != null && diagnostic.Detail.Contains("maximumPageHeight", StringComparison.Ordinal));
        }
    }
}
