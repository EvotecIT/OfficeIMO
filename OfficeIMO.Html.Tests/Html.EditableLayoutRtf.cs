using OfficeIMO.Html;
using OfficeIMO.Rtf;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class HtmlEditableLayoutRtfTests {
    [Fact]
    public void PositionedRichContentStaysInSemanticFlow() {
        const string html = "<div style='position:absolute;width:200px;height:60px'>" +
            "<strong>Bold</strong> <a href='https://example.test'>Linked</a></div>";

        HtmlToRtfResult result = HtmlConversionDocument.Parse(html).ToRtfDocumentResult();
        string rtf = result.Value.ToRtf();

        Assert.DoesNotContain(@"\phpg", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\b ", rtf, StringComparison.Ordinal);
        Assert.Contains("HYPERLINK", rtf, StringComparison.Ordinal);
        Assert.Contains("https://example.test", rtf, StringComparison.Ordinal);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified);
    }

    [Fact]
    public void PositionedRegionRetainsEmbeddedPngPicture() {
        byte[] png = PdfPngTestImages.CreateRgbPng(4, 3);
        string image = "data:image/png;base64," + Convert.ToBase64String(png);
        string html = "<div style='position:absolute;width:180px;height:70px'>Region picture" +
            "<img alt='Hidden marker' src='" + image + "' style='display:none'>" +
            "<img alt='Region marker' src='" + image + "' style='width:24px;height:18px'></div>";

        HtmlToRtfResult result = HtmlConversionDocument.Parse(html).ToRtfDocumentResult();
        RtfReadResult reopened = RtfDocument.Read(result.Value.ToRtf());

        RtfParagraph paragraph = Assert.Single(reopened.Document.Paragraphs, item =>
            item.ToPlainText().Contains("Region picture", StringComparison.Ordinal));
        RtfImage picture = Assert.Single(paragraph.Inlines.OfType<RtfImage>());
        Assert.Equal(RtfImageFormat.Png, picture.Format);
        Assert.Equal(360, picture.DesiredWidthTwips);
        Assert.Equal(270, picture.DesiredHeightTwips);
    }

    [Fact]
    public void PrintRegionsStaySemanticWhenRenderedPageOwnershipCannotBeMapped() {
        const string html = "<div style='position:absolute;width:160px;height:40px'>Print anchor</div>" +
            "<section style='break-before:page'><p>Later page</p></section>";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html, new HtmlConversionDocumentOptions {
            Profile = HtmlConversionProfile.HighFidelityPrint
        });

        HtmlToRtfResult result = document.ToRtfDocumentResult();
        string rtf = result.Value.ToRtf();

        Assert.Contains("Print anchor", string.Join("\n", result.Value.Paragraphs.Select(
            paragraph => paragraph.ToPlainText())), StringComparison.Ordinal);
        Assert.DoesNotContain(@"\phpg", rtf, StringComparison.Ordinal);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified);
    }

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
