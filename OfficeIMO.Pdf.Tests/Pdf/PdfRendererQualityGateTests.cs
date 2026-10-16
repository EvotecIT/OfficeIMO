using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Filters;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

/// <summary>
/// Deterministic fidelity gate: every corpus page renders at the baseline size and stays at or above the PSNR and
/// SSIM thresholds recorded in manifest.json against the checked-in Poppler baselines.
/// </summary>
public sealed class PdfRendererQualityGateTests {
    public static IEnumerable<object[]> CorpusPages() {
        using JsonDocument manifest = JsonDocument.Parse(File.ReadAllText(PdfRendererQualitySupport.CorpusPath("manifest.json")));
        foreach (JsonElement page in manifest.RootElement.GetProperty("pages").EnumerateArray()) {
            yield return new object[] { page.GetProperty("id").GetString()! };
        }
    }

    [Theory]
    [MemberData(nameof(CorpusPages))]
    public void CorpusPageMeetsRecordedFidelityThresholds(string id) {
        using JsonDocument manifest = JsonDocument.Parse(File.ReadAllText(PdfRendererQualitySupport.CorpusPath("manifest.json")));
        JsonElement page = manifest.RootElement.GetProperty("pages").EnumerateArray().Single(entry => entry.GetProperty("id").GetString() == id);
        Assert.True(page.TryGetProperty("minimumPsnr", out JsonElement minimumPsnr), id + " has no PSNR threshold.");
        Assert.True(page.TryGetProperty("minimumSsim", out JsonElement minimumSsim), id + " has no SSIM threshold.");

        OfficeRasterImage actual = PdfRendererQualitySupport.Render(page.GetProperty("pdf").GetString()!);
        OfficeRasterImage expected = PdfRendererQualitySupport.Baseline(page.GetProperty("baseline").GetString()!);

        Assert.Equal((expected.Width, expected.Height), (actual.Width, actual.Height));
        double psnr = PdfRendererQualitySupport.Psnr(actual, expected);
        double ssim = PdfRendererQualitySupport.Ssim(actual, expected);
        Assert.True(psnr >= minimumPsnr.GetDouble(), $"{id}: PSNR {psnr:0.00} dB is below {minimumPsnr.GetDouble():0.00} dB.");
        Assert.True(ssim >= minimumSsim.GetDouble(), $"{id}: SSIM {ssim:0.0000} is below {minimumSsim.GetDouble():0.0000}.");
    }

    [Fact]
    public void Jbig2RegionsDecodeBitExactToPoppler() {
        PdfReadDocument document = PdfReadDocument.Open(File.ReadAllBytes(PdfRendererQualitySupport.CorpusPath("jbig2.pdf")));
        PdfStream image = Jbig2Images(document).Single();
        int width = (int)image.Dictionary.Get<PdfNumber>("Width")!.Value, height = (int)image.Dictionary.Get<PdfNumber>("Height")!.Value;
        byte[] decoded = StreamDecoder.Decode(image, document.Objects);
        OfficeRasterImage expected = PdfRendererQualitySupport.Baseline("jbig2.poppler-decoded.png");
        Assert.Equal((width, height), (expected.Width, expected.Height));

        int stride = (width + 7) / 8, mismatches = 0;
        for (int y = 0; y < height; y++) {
            for (int x = 0; x < width; x++) {
                bool black = ((decoded[y * stride + x / 8] >> (7 - (x & 7))) & 1) == 0;
                if (black != expected.GetPixel(x, y).R < 128) mismatches++;
            }
        }
        Assert.Equal(0, mismatches);
    }

    [Fact]
    public void Jbig2DecoderRejectsCorruptStreamsWithBoundedExceptions() {
        PdfReadDocument document = PdfReadDocument.Open(File.ReadAllBytes(PdfRendererQualitySupport.CorpusPath("jbig2.pdf")));
        PdfStream image = Jbig2Images(document).Single();
        var globalsReference = (PdfReference)image.Dictionary.Get<PdfDictionary>("DecodeParms")!.Items["JBIG2Globals"];
        var globals = (PdfStream)document.Objects[globalsReference.ObjectNumber].Value;
        byte[] data = image.Data, globalData = globals.Data;
        var random = new Random(2026);
        for (int iteration = 0; iteration < 400; iteration++) {
            byte[] mutated = (byte[])data.Clone();
            int edits = 1 + random.Next(6);
            for (int edit = 0; edit < edits; edit++) mutated[random.Next(mutated.Length)] = (byte)random.Next(256);
            int length = iteration % 5 == 0 ? random.Next(mutated.Length) : mutated.Length;
            Array.Resize(ref mutated, length);
            try {
                Jbig2Decoder.Decode(mutated, iteration % 2 == 0 ? globalData : null, 400, 300, 1 << 20, default);
            } catch (Exception exception) when (exception is InvalidDataException or NotSupportedException) {
            }
        }
    }

    [Fact]
    public void UnsupportedJbig2SegmentsReportTheOmittedImage() {
        var jbig2 = new MemoryStream();
        // Page information (type 48), then a halftone region (type 22), which the decoder does not implement.
        WriteBytes(jbig2, new byte[] { 0, 0, 0, 0, 48, 0, 1, 0, 0, 0, 19, 0, 0, 0, 8, 0, 0, 0, 8, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 });
        WriteBytes(jbig2, new byte[] { 0, 0, 0, 1, 22, 0, 1, 0, 0, 0, 0 });
        byte[] pdf = BuildImagePdf(jbig2.ToArray(), "/Filter /JBIG2Decode");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        Assert.Contains(result.CapabilityDiagnostics, static diagnostic =>
            diagnostic.Code == PdfRenderCapabilities.UnsupportedImageId && diagnostic.Subject == "Im1:JBIG2-halftone");
    }

    private static IEnumerable<PdfStream> Jbig2Images(PdfReadDocument document) =>
        document.Objects.Values.Select(entry => entry.Value).OfType<PdfStream>()
            .Where(stream => stream.Dictionary.Get<PdfName>("Subtype")?.Name == "Image" &&
                stream.Dictionary.Get<PdfName>("Filter")?.Name == "JBIG2Decode");

    private static void WriteBytes(Stream stream, byte[] bytes) => stream.Write(bytes, 0, bytes.Length);

    private static byte[] BuildImagePdf(byte[] imageData, string filter) {
        using var pdf = new MemoryStream();
        void Ascii(string text) { byte[] bytes = Encoding.ASCII.GetBytes(text); pdf.Write(bytes, 0, bytes.Length); }
        const string content = "q 80 0 0 80 20 20 cm /Im1 Do Q";
        Ascii("%PDF-1.4\n1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        Ascii("2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 120 120] >>\nendobj\n");
        Ascii("3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /XObject << /Im1 5 0 R >> >> /Contents 4 0 R >>\nendobj\n");
        Ascii($"4 0 obj\n<< /Length {content.Length} >>\nstream\n{content}\nendstream\nendobj\n");
        Ascii($"5 0 obj\n<< /Type /XObject /Subtype /Image /Width 8 /Height 8 /ColorSpace /DeviceGray /BitsPerComponent 1 {filter} /Length {imageData.Length} >>\nstream\n");
        pdf.Write(imageData, 0, imageData.Length);
        Ascii("\nendstream\nendobj\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return pdf.ToArray();
    }
}
