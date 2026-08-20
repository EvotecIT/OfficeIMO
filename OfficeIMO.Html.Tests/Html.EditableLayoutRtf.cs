using OfficeIMO.Html;
using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class HtmlEditableLayoutRtfTests {
    [Fact]
    public void PositionedAndFloatingRegionsReopenAsEditableRtfFrames() {
        const string html = "<style>" +
            ".positioned{position:absolute;left:32px;top:24px;width:240px;height:72px;background:#dbeafe;z-index:4}" +
            ".floating{float:right;width:120px;height:48px;background:#fef3c7}" +
            ".flex{display:flex;width:300px}</style>" +
            "<p>Ordinary flow</p><div class='positioned'>Editable positioned</div>" +
            "<div class='floating'>Editable float</div><div class='flex'><span>Flex remains</span></div>";
        HtmlToRtfResult result = HtmlConversionDocument.Parse(html).ToRtfDocumentResult();
        string rtf = result.Value.ToRtf();

        RtfReadResult reopened = RtfDocument.Read(rtf);
        RtfParagraph positioned = Assert.Single(reopened.Document.Paragraphs, paragraph =>
            paragraph.ToPlainText().Contains("Editable positioned", StringComparison.Ordinal));
        RtfParagraph floating = Assert.Single(reopened.Document.Paragraphs, paragraph =>
            paragraph.ToPlainText().Contains("Editable float", StringComparison.Ordinal));

        Assert.Equal(1200, positioned.Frame.HorizontalPositionTwips);
        Assert.Equal(1080, positioned.Frame.VerticalPositionTwips);
        Assert.Equal(3600, positioned.Frame.WidthTwips);
        Assert.Equal(-1080, positioned.Frame.HeightTwips);
        Assert.True(positioned.Frame.NoWrap);
        Assert.True(positioned.Frame.OverlayText);
        Assert.False(floating.Frame.NoWrap);
        Assert.True(floating.Frame.NoOverlap);
        Assert.Contains("Flex remains", string.Join("\n", reopened.Document.Paragraphs.Select(paragraph => paragraph.ToPlainText())), StringComparison.Ordinal);
        Assert.Contains(@"\phpg", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\pvpg", rtf, StringComparison.Ordinal);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.RegionProjected);
        Assert.True(result.Succeeded);
    }
}
