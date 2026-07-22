using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfOptimizerAdvancedTests {
    [Fact]
    public void MaximumCompression_EmitsDeterministicObjectAndXrefStreams() {
        byte[] source = PdfDocument.Create()
            .Meta(title: "Optimization profile")
            .Paragraph(p => p.Text("Object stream proof"))
            .PageBreak()
            .Paragraph(p => p.Text("Second page proof"))
            .ToBytes();

        PdfOptimizationActionResult first = PdfOptimizer.Optimize(source, PdfOptimizationProfile.MaximumCompression);
        PdfOptimizationActionResult second = PdfOptimizer.Optimize(source, PdfOptimizationProfile.MaximumCompression);
        PdfDocumentInfo info = PdfInspector.Inspect(first.Bytes);

        Assert.Equal(first.Bytes, second.Bytes);
        Assert.True(info.Security.HasXrefStreams);
        Assert.True(info.Security.HasObjectStreams);
        Assert.Contains("Object stream proof", PdfTextExtractor.ExtractAllText(first.Bytes), StringComparison.Ordinal);
        Assert.Contains("Second page proof", PdfTextExtractor.ExtractAllText(first.Bytes), StringComparison.Ordinal);
        Assert.Contains(first.Actions, static action => action.Kind == "PackObjectStreams");
        Assert.Contains(first.Actions, static action => action.Kind == "WriteXrefStream");
        Assert.True(first.PreservationReport.IsPreserved);

        var normalize = PdfOptimizationOptions.Create(PdfOptimizationProfile.Custom);
        normalize.KeepOriginalWhenNotSmaller = false;
        PdfOptimizationActionResult normalized = PdfOptimizer.Optimize(first.Bytes, normalize);
        Assert.False(PdfInspector.Inspect(normalized.Bytes).Security.HasObjectStreams);
        Assert.Contains("Object stream proof", PdfTextExtractor.ExtractAllText(normalized.Bytes), StringComparison.Ordinal);
    }

    [Fact]
    public void WebProfile_EmitsDeterministicStandardsCompliantLinearization() {
        byte[] source = PdfDocument.Create()
            .Paragraph(p => p.Text("Web optimization proof"))
            .PageBreak()
            .Paragraph(p => p.Text("Second web page"))
            .ToBytes();

        PdfOptimizationActionResult result = PdfOptimizer.Optimize(source, PdfOptimizationProfile.Web);
        PdfOptimizationActionResult repeated = PdfOptimizer.Optimize(source, PdfOptimizationProfile.Web);
        string raw = PdfEncoding.Latin1GetString(result.Bytes);

        Assert.Equal(result.Bytes, repeated.Bytes);
        Assert.Contains("/Linearized 1", raw, StringComparison.Ordinal);
        Assert.True(result.CandidateLinearized);
        Assert.Equal(2, PdfInspector.Inspect(result.Bytes).PageCount);
        Assert.Contains("Web optimization proof", PdfTextExtractor.ExtractAllText(result.Bytes), StringComparison.Ordinal);
        Assert.Contains("Second web page", PdfTextExtractor.ExtractAllText(result.Bytes), StringComparison.Ordinal);
        Assert.Contains(result.Actions, static action => action.Kind == "Linearize");
        Assert.True(result.PreservationReport.IsPreserved);

        var (objects, _) = PdfSyntax.ParseObjects(result.Bytes);
        PdfDictionary linearization = Assert.Single(
            objects.Values.Select(static item => item.Value).OfType<PdfDictionary>(),
            static dictionary => dictionary.Items.ContainsKey("Linearized"));
        Assert.Equal(result.Bytes.Length, linearization.Get<PdfNumber>("L")?.Value);
        Assert.Equal(2D, linearization.Get<PdfNumber>("N")?.Value);
        Assert.True(linearization.Get<PdfNumber>("O")?.Value > 0D);
        Assert.True(linearization.Get<PdfNumber>("E")?.Value > 0D);
        Assert.True(linearization.Get<PdfNumber>("T")?.Value > 0D);
        PdfArray hints = Assert.IsType<PdfArray>(linearization.Items["H"]);
        Assert.Equal(2, hints.Items.Count);
        int hintOffset = checked((int)Assert.IsType<PdfNumber>(hints.Items[0]).Value);
        int hintLength = checked((int)Assert.IsType<PdfNumber>(hints.Items[1]).Value);
        Assert.InRange(hintOffset, 1, result.Bytes.Length - 1);
        Assert.InRange(hintLength, 1, result.Bytes.Length - hintOffset);
        Assert.Contains("/S ", PdfEncoding.Latin1GetString(result.Bytes.AsSpan(hintOffset, hintLength).ToArray()), StringComparison.Ordinal);
        PdfStream hintStream = Assert.Single(
            objects.Values.Select(static item => item.Value).OfType<PdfStream>(),
            static stream => stream.Dictionary.Items.ContainsKey("S"));
        uint hintedFirstPageOffset = System.Buffers.Binary.BinaryPrimitives.ReadUInt32BigEndian(hintStream.Data.AsSpan(4, 4));
        int firstPageObjectId = checked((int)linearization.Get<PdfNumber>("O")!.Value);
        int actualFirstPageOffset = raw.IndexOf(firstPageObjectId.ToString(System.Globalization.CultureInfo.InvariantCulture) + " 0 obj\n", StringComparison.Ordinal);
        Assert.Equal(actualFirstPageOffset, checked((int)hintedFirstPageOffset) + hintLength);
        int mainXrefFirstEntryOffset = checked((int)linearization.Get<PdfNumber>("T")!.Value);
        Assert.StartsWith("0000000000 65535 f ", raw.Substring(mainXrefFirstEntryOffset), StringComparison.Ordinal);

        var incompatible = PdfOptimizationOptions.Create(PdfOptimizationProfile.Custom);
        incompatible.Linearize = true;
        incompatible.UseObjectStreams = true;
        NotSupportedException exception = Assert.Throws<NotSupportedException>(() => PdfOptimizer.Optimize(source, incompatible));
        Assert.Contains("classic cross-reference", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Optimize_DeduplicatesEquivalentFontsAndResourceDictionaries() {
        byte[] source = Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 2 /Kids [8 0 R 9 0 R] >>", "endobj",
            "4 0 obj", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>", "endobj",
            "5 0 obj", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>", "endobj",
            "6 0 obj", "<< /Font << /F1 4 0 R >> >>", "endobj",
            "7 0 obj", "<< /Font << /F1 5 0 R >> >>", "endobj",
            "8 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources 6 0 R /Contents 10 0 R >>", "endobj",
            "9 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources 7 0 R /Contents 11 0 R >>", "endobj",
            "10 0 obj", "<< /Length 32 >>", "stream", "BT /F1 12 Tf (First) Tj ET", "endstream", "endobj",
            "11 0 obj", "<< /Length 33 >>", "stream", "BT /F1 12 Tf (Second) Tj ET", "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 12 >>", "%%EOF"
        }));
        var options = PdfOptimizationOptions.Create(PdfOptimizationProfile.Custom);
        options.CompressUnfilteredStreams = false; options.DeduplicateIdenticalStreams = false; options.DeduplicateImages = false; options.KeepOriginalWhenNotSmaller = false;

        PdfOptimizationActionResult result = PdfOptimizer.Optimize(source, options);

        Assert.Contains(result.Actions, static action => action.Kind == "DeduplicateFont");
        Assert.Contains(result.Actions, static action => action.Kind == "DeduplicateResource");
        Assert.Equal(2, PdfInspector.Inspect(result.Bytes).PageCount);
        Assert.Contains("First", PdfTextExtractor.ExtractAllText(result.Bytes), StringComparison.Ordinal);
        Assert.Contains("Second", PdfTextExtractor.ExtractAllText(result.Bytes), StringComparison.Ordinal);
    }

    [Fact]
    public void Optimize_DeduplicatesLosslesslyEquivalentDecodedImages() {
        byte[] source = Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << /XObject << /Im1 4 0 R /Im2 5 0 R >> >> /Contents 6 0 R >>", "endobj",
            "4 0 obj", "<< /Type /XObject /Subtype /Image /Width 2 /Height 2 /ColorSpace /DeviceGray /BitsPerComponent 8 /Length 4 >>", "stream", "ABCD", "endstream", "endobj",
            "5 0 obj", "<< /Type /XObject /Subtype /Image /Width 2 /Height 2 /ColorSpace /DeviceGray /BitsPerComponent 8 /Filter /ASCIIHexDecode /Length 9 >>", "stream", "41424344>", "endstream", "endobj",
            "6 0 obj", "<< /Length 49 >>", "stream", "q 20 0 0 20 10 10 cm /Im1 Do Q q 20 0 0 20 40 10 cm /Im2 Do Q", "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 7 >>", "%%EOF"
        }));
        var options = PdfOptimizationOptions.Create(PdfOptimizationProfile.Custom);
        options.CompressUnfilteredStreams = false; options.DeduplicateIdenticalStreams = false; options.DeduplicateFonts = false; options.DeduplicateResources = false; options.KeepOriginalWhenNotSmaller = false;

        PdfOptimizationActionResult result = PdfOptimizer.Optimize(source, options);

        Assert.Contains(result.Actions, static action => action.Kind == "DeduplicateImage");
        Assert.Single(PdfReadDocument.Open(result.Bytes).Pages[0].GetImagePlacements().Select(static placement => placement.ObjectNumber).Distinct());
        Assert.True(result.PreservationReport.IsPreserved);

        options.MaximumTotalDecodedImageBytes = 4;
        PdfOptimizationActionResult limited = PdfOptimizer.Optimize(source, options);
        Assert.DoesNotContain(limited.Actions, static action => action.Kind == "DeduplicateImage");
        Assert.Contains(limited.SkippedActions, static action =>
            action.Kind == "DeduplicateImage" && action.Reason == "AggregateDecodeLimit");

        options.MaximumTotalDecodedImageBytes = 0;
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfOptimizer.Optimize(source, options));
    }
}
