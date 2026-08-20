using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class HtmlEditableLayoutWordTests {
    [Fact]
    public void PositionedRegionRetainsItsForegroundPictureAndEffects() {
        string image = "data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(4, 3));
        string html = "<div style='position:absolute;width:180px;height:70px;background-image:url(\"" + image + "\")'>" +
            "Region picture<img alt='Region marker' src='" + image + "' style='width:24px;height:18px;opacity:.4'></div>";

        HtmlToWordResult result = HtmlConversionDocument.Parse(html).ToWordDocumentResult();
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.RegionImageOmitted);
        using var stream = new MemoryStream();
        result.Value.Save(stream);
        result.Value.Dispose();

        byte[] bytes = stream.ToArray();
        using WordDocument reopened = WordDocument.Load(new MemoryStream(bytes),
            new WordLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
        Assert.Single(reopened.TextBoxes, textBox => textBox.Paragraphs.Any(paragraph =>
            paragraph.Text.Contains("Region picture", StringComparison.Ordinal)));
        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        Assert.Single(package.MainDocumentPart!.ImageParts);
        using var reader = new StreamReader(package.MainDocumentPart!.GetStream());
        string documentXml = reader.ReadToEnd();
        Assert.Contains(":blip", documentXml, StringComparison.Ordinal);
        Assert.Contains(":alphaModFix amt=\"40000\"", documentXml, StringComparison.Ordinal);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.BackgroundLayersFlattened);
    }

    [Fact]
    public void PositionedAndFloatingRegionsReopenAsEditableWordAnchors() {
        const string html = "<style>" +
            ".positioned{position:absolute;left:32px;top:24px;width:240px;height:72px;background:#dbeafe;z-index:4}" +
            ".floating{float:right;width:120px;height:48px;background:#fef3c7}" +
            ".flex{display:flex;width:300px}</style>" +
            "<p>Ordinary flow</p><div class='positioned'>Editable positioned</div>" +
            "<div class='floating'>Editable float</div><div class='flex'><span>Flex remains</span></div>";
        HtmlToWordResult result = HtmlConversionDocument.Parse(html).ToWordDocumentResult();
        using var stream = new MemoryStream();
        result.Value.Save(stream);
        result.Value.Dispose();

        using WordDocument reopened = WordDocument.Load(
            new MemoryStream(stream.ToArray()),
            new WordLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
        WordTextBox positioned = Assert.Single(reopened.TextBoxes, textBox =>
            textBox.Paragraphs.Any(paragraph => paragraph.Text.Contains("Editable positioned", StringComparison.Ordinal)));
        WordTextBox floating = Assert.Single(reopened.TextBoxes, textBox =>
            textBox.Paragraphs.Any(paragraph => paragraph.Text.Contains("Editable float", StringComparison.Ordinal)));

        Assert.Equal(762000, positioned.HorizontalPositionOffset);
        Assert.Equal(685800, positioned.VerticalPositionOffset);
        Assert.Equal(2286000L, positioned.Width);
        Assert.Equal(685800L, positioned.Height);
        Assert.Equal("DBEAFE", positioned.FillColorHex);
        Assert.Equal(WordImageTextWrapping.InFrontOfText, positioned.WrapText);
        Assert.Equal(WordImageTextWrapping.Square, floating.WrapText);
        Assert.NotEmpty(reopened.Find("Flex remains", StringComparison.Ordinal));
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.RegionProjected);
        Assert.True(result.Succeeded);
    }
}
