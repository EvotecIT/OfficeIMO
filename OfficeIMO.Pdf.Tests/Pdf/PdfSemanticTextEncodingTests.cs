using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfSemanticTextEncodingTests {
    [Theory]
    [InlineData(false, "• – “quote” €")]
    [InlineData(true, "• – “quote” €")]
    [InlineData(false, "Nonbreaking\u00A0space\u00AD")]
    public void GeneratedActualTextAndMetadataKeepUnicodeInIndependentReaders(bool drawing, string text) {
        var document = PdfDocument.Create(new PdfOptions { CompressContentStreams = false }).Meta(title: text);
        if (drawing) {
            var ink = new OfficeDrawing(200D, 30D).AddText("Paint", 0D, 0D, 180D, 20D,
                new OfficeFontInfo("Helvetica", 12D));
            var scene = new OfficeDrawing(200D, 30D).AddActualTextDrawing(text, ink, 0D, 0D);
            document.Canvas(canvas => canvas.Drawing(scene, 10D, 10D, 200D, 30D));
        } else {
            document.Canvas(canvas => canvas.ActualText(text, 10D, 10D,
                logical => logical.Text("Paint", 10D, 10D, 180D, 20D)));
        }

        byte[] bytes = document.ToBytes();
        using var independent = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal(text, independent.Information.Title);
        Assert.Equal(text, Assert.Single(GetActualText(independent.GetPage(1).GetMarkedContents())));
        Assert.Equal(text.Replace('\u00A0', ' '), PdfReadDocument.Open(bytes).ExtractText().Trim());
        Assert.Equal(text, PdfInspector.Inspect(bytes).Metadata.Title);
    }

    [Theory]
    [InlineData("<809395A0>", "•ﬁŁ€", new byte[] { 0x80, 0x93, 0x95, 0xA0 })]
    [InlineData("(\\200\\223\\225\\240)", "•ﬁŁ€", new byte[] { 0x80, 0x93, 0x95, 0xA0 })]
    public void ParsedSemanticStringsUsePdfDocEncodingAndRetainGlyphBytes(string token, string text, byte[] rawBytes) {
        byte[] bytes = Encoding.ASCII.GetBytes("%PDF-1.7\n1 0 obj\n" + token + "\nendobj\n%%EOF\n");
        var parsed = Assert.IsType<PdfStringObj>(Assert.Single(PdfSyntax.ParseObjects(bytes).Map).Value.Value);

        Assert.Equal(text, parsed.Value);
        Assert.Equal(rawBytes, parsed.RawBytes);
        // Font-encoded content strings still use their font's WinAnsi mapping.
        Assert.Equal("€“•\u00A0", PdfWinAnsiEncoding.Decode(parsed.RawBytes));
    }

    [Fact]
    public void SemanticStringObjectsUseUnicodeForNonAsciiText() {
        const string text = "• €\u00A0";
        var value = new PdfStringObj(text, useTextStringEncoding: true);

        Assert.Equal(new byte[] { 0xFE, 0xFF, 0x20, 0x22, 0, 0x20, 0x20, 0xAC, 0, 0xA0 }, value.RawBytes);
        Assert.Equal(text, PdfTextString.Decode(value.RawBytes));
    }

    [Theory]
    [InlineData("Chapter\u00A0note\u00AD")]
    [InlineData("• – “quote” €")]
    public void GeneratedOutlineTitlesAndLinkCommentsKeepUnicode(string text) {
        byte[] bytes = PdfDocument.Create()
            .Canvas(canvas => canvas.Outline(text, 1, 10D))
            .Paragraph(paragraph => paragraph.Link("Link", "https://evotec.xyz/", contents: text))
            .ToBytes();

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        Assert.Equal(text, Assert.Single(info.Outlines).Title);
        Assert.Equal(text, Assert.Single(info.GetAnnotationsBySubtype("Link")).Contents);
        using var independent = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.True(independent.TryGetBookmarks(out var bookmarks));
        Assert.Equal(text, Assert.Single(bookmarks.Roots).Title);
        Assert.Equal(text, Assert.Single(independent.GetPage(1).GetAnnotations()).Content);
    }

    [Fact]
    public void SemanticStringByteBudgetMatchesTheSerializedUnicodeObject() {
        var value = new PdfStringObj("Résumé •", useTextStringEncoding: true);
        var context = new PdfPageExtractor.SerializationContext(new Dictionary<int, int>(), 0,
            new Dictionary<int, Dictionary<string, PdfObject>>());
        byte[] bytes = PdfPageExtractor.SerializeObject(value, context);

        PdfPageExtractor.EnsureSerializedObjectWithinLimit(value, context, bytes.LongLength);
        Assert.Throws<InvalidDataException>(() =>
            PdfPageExtractor.EnsureSerializedObjectWithinLimit(value, context, bytes.LongLength - 1));
    }

    private static IEnumerable<string> GetActualText(IEnumerable<UglyToad.PdfPig.Content.MarkedContentElement> elements) {
        foreach (var element in elements) {
            if (!string.IsNullOrEmpty(element.ActualText)) yield return element.ActualText!;
            else foreach (string text in GetActualText(element.Children)) yield return text;
        }
    }
}
