using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfStringRewritePreservationTests {
    [Theory]
    [InlineData("Extract")]
    [InlineData("Merge")]
    [InlineData("Optimize")]
    [InlineData("Annotation")]
    public void SupportedRewritesRetainSemanticText(string operation) {
        const string text = "Chapter\u00A0note\u00AD";
        byte[] source = PdfDocument.Create().Meta(title: text)
            .Canvas(canvas => canvas.Outline(text, 1, 10D))
            .Paragraph(paragraph => paragraph.Link("Link", "https://evotec.xyz/", contents: text))
            .ToBytes();
        AssertSemanticText(source, text);

        AssertSemanticText(Rewrite(source, operation), text);
    }

    [Theory]
    [InlineData("Extract")]
    [InlineData("Merge")]
    [InlineData("Optimize")]
    [InlineData("Annotation")]
    public void SupportedRewritesRetainBinaryIndexedPaletteAndPixels(string operation) {
        byte[] source = BuildIndexedImagePdf();
        AssertIndexedImage(source);

        AssertIndexedImage(Rewrite(source, operation));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MetadataReplacementEncodesNewSemanticValues(bool incremental) {
        const string text = "• Chapter\u00A0note\u00AD";
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Body")).ToBytes();

        byte[] output = incremental
            ? PdfIncrementalUpdater.UpdateMetadata(source, title: text)
            : PdfMetadataEditor.SynchronizeMetadata(source, title: text);

        Assert.Equal(text, PdfInspector.Inspect(output).Metadata.Title);
        using var independent = UglyToad.PdfPig.PdfDocument.Open(output);
        Assert.Equal(text, independent.Information.Title);
    }

    [Theory]
    [InlineData("<FEFF004100A000AD>", new byte[] { 0xFE, 0xFF, 0, 65, 0, 0xA0, 0, 0xAD })]
    [InlineData("<A04040>", new byte[] { 0xA0, 0x40, 0x40 })]
    [InlineData("(A\\(B)", new byte[] { 65, 40, 66 })]
    public void ParsedStringRewriteAndByteBudgetRetainOriginalBytes(string token, byte[] expected) {
        byte[] source = Encoding.ASCII.GetBytes("%PDF-1.7\n1 0 obj\n" + token + "\nendobj\n%%EOF\n");
        var value = Assert.IsType<PdfStringObj>(Assert.Single(PdfSyntax.ParseObjects(source).Map).Value.Value);
        var context = new PdfPageExtractor.SerializationContext(new Dictionary<int, int>(), 0,
            new Dictionary<int, Dictionary<string, PdfObject>>());
        byte[] serialized = PdfPageExtractor.SerializeObject(value, context);
        byte[] rewritten = Encoding.ASCII.GetBytes("%PDF-1.7\n1 0 obj\n")
            .Concat(serialized).Concat(Encoding.ASCII.GetBytes("endobj\n%%EOF\n")).ToArray();
        var parsed = Assert.IsType<PdfStringObj>(Assert.Single(PdfSyntax.ParseObjects(rewritten).Map).Value.Value);

        Assert.Equal(expected, parsed.RawBytes);
        PdfPageExtractor.EnsureSerializedObjectWithinLimit(value, context, serialized.LongLength);
        Assert.Throws<InvalidDataException>(() =>
            PdfPageExtractor.EnsureSerializedObjectWithinLimit(value, context, serialized.LongLength - 1));
    }

    [Fact]
    public void BinaryStringConstructionAndRewriteKeepAllByteValues() {
        byte[] expected = Enumerable.Range(0, 256).Select(value => (byte)value).ToArray();
        var value = new PdfStringObj(expected);
        var context = new PdfPageExtractor.SerializationContext(new Dictionary<int, int>(), 0,
            new Dictionary<int, Dictionary<string, PdfObject>>());
        byte[] serialized = PdfPageExtractor.SerializeObject(value, context);
        byte[] rewritten = Encoding.ASCII.GetBytes("%PDF-1.7\n1 0 obj\n")
            .Concat(serialized).Concat(Encoding.ASCII.GetBytes("endobj\n%%EOF\n")).ToArray();
        var parsed = Assert.IsType<PdfStringObj>(Assert.Single(PdfSyntax.ParseObjects(rewritten).Map).Value.Value);

        Assert.Equal(expected, parsed.RawBytes);
        PdfPageExtractor.EnsureSerializedObjectWithinLimit(value, context, serialized.LongLength);
        Assert.Throws<InvalidDataException>(() =>
            PdfPageExtractor.EnsureSerializedObjectWithinLimit(value, context, serialized.LongLength - 1));
    }

    private static byte[] Rewrite(byte[] source, string operation) => operation switch {
        "Extract" => PdfPageExtractor.ExtractPages(source, 1),
        "Merge" => PdfMerger.Merge(source),
        "Optimize" => PdfOptimizer.Optimize(source, new PdfOptimizationOptions { KeepOriginalWhenNotSmaller = false }).Bytes,
        "Annotation" => PdfAnnotationEditor.UpdateAnnotation(source,
            Assert.Single(PdfInspector.Inspect(source).Annotations).ObjectNumber!.Value,
            new PdfAnnotationUpdateOptions { Flags = 4 }).Bytes,
        _ => throw new ArgumentOutOfRangeException(nameof(operation))
    };

    private static void AssertSemanticText(byte[] bytes, string text) {
        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        Assert.Equal(text, info.Metadata.Title);
        Assert.Equal(text, Assert.Single(info.Outlines).Title);
        Assert.Equal(text, Assert.Single(info.GetAnnotationsBySubtype("Link")).Contents);
        using var independent = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal(text, independent.Information.Title);
        Assert.True(independent.TryGetBookmarks(out var bookmarks));
        Assert.Equal(text, Assert.Single(bookmarks.Roots).Title);
        Assert.Equal(text, Assert.Single(independent.GetPage(1).GetAnnotations()).Content);
    }

    private static void AssertIndexedImage(byte[] bytes) {
        var objects = PdfSyntax.ParseObjects(bytes).Map;
        var image = Assert.Single(objects.Values.Select(item => item.Value).OfType<PdfStream>(),
            stream => stream.Dictionary.Get<PdfName>("Subtype")?.Name == "Image");
        var colorSpace = Assert.IsType<PdfArray>(image.Dictionary.Items["ColorSpace"]);
        var lookup = Assert.IsType<PdfStringObj>(colorSpace.Items[3]);
        Assert.Equal(new byte[] { 160, 64, 64 }, lookup.RawBytes);
        PdfExtractedImage extracted = Assert.Single(PdfImageExtractor.ExtractImages(bytes));
        Assert.Equal(new byte[] { 0, 160, 64, 64 }, PdfPngTestImages.DecodePngIdat(extracted.Bytes));
    }

    private static byte[] BuildIndexedImagePdf() {
        const string content = "q\n20 0 0 20 0 0 cm\n/Im1 Do\nQ\n";
        byte[][] objects = {
            PdfPageExtractor.WrapObject(1, Encoding.ASCII.GetBytes("<< /Type /Catalog /Pages 2 0 R >>\n")),
            PdfPageExtractor.WrapObject(2, Encoding.ASCII.GetBytes("<< /Type /Pages /Count 1 /Kids [3 0 R] >>\n")),
            PdfPageExtractor.WrapObject(3, Encoding.ASCII.GetBytes("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 30 30] /Resources << /XObject << /Im1 5 0 R >> >> /Contents 4 0 R /Annots [6 0 R] >>\n")),
            PdfPageExtractor.WrapObject(4, PdfObjectBytes.WrapStreamBody("<< /Length " + content.Length + " >>\n", Encoding.ASCII.GetBytes(content))),
            PdfPageExtractor.WrapObject(5, PdfObjectBytes.WrapStreamBody("<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /BitsPerComponent 8 /ColorSpace [/Indexed /DeviceRGB 0 <A04040>] /Length 1 >>\n", new byte[] { 0 })),
            PdfPageExtractor.WrapObject(6, Encoding.ASCII.GetBytes("<< /Type /Annot /Subtype /Text /P 3 0 R /Rect [22 22 28 28] /Contents (Keep) >>\n"))
        };
        return PdfFileAssembler.Assemble(objects, 1, 0);
    }
}
