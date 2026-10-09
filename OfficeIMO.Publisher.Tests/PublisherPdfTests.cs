using OfficeIMO.Publisher.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherPdfTests {
    [Fact]
    public void Pdf_keeps_document_pages_and_carries_source_and_rendering_evidence() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("Sample.pub"));
        var result = document.ToPdfDocumentResult();
        byte[] bytes = result.ToBytes();
        PdfReadDocument pdf = PdfReadDocument.Open(bytes);
        Assert.Equal(2, pdf.Pages.Count);
        string text = pdf.ExtractText();
        Assert.Contains("This is some text", text);
        Assert.Contains("Bottom Right", text);
        Assert.Contains("second page", text);
        Assert.Contains(document.ReadReport, result.SourceConversionReports);
        Assert.True(result.HasLoss);
        Assert.Throws<OfficeConversionException>(() => document.ReadReport.RequireNoLoss());
    }

    [Fact]
    public void Pdf_save_keeps_the_caller_stream_open_and_reports_source_losses() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("Simple.pub"));
        using var output = new MemoryStream();
        var result = document.SaveAsPdf(output);
        Assert.True(result.Succeeded);
        Assert.True(output.CanWrite);
        Assert.True(output.Length > 0);
        Assert.True(result.HasLoss);
    }
    [Theory]
    [InlineData("SampleBrochure.pub", 2)]
    [InlineData("SampleNewsletter.pub", 4)]
    public void Unsupported_native_metafiles_do_not_abort_publication_exports(string file, int pages) {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture(file));
        var result = document.ToPdfDocumentResult();
        PdfReadDocument pdf = PdfReadDocument.Open(result.ToBytes());
        Assert.Equal(pages, pdf.Pages.Count);
        Assert.Contains(document.ReadReport.FidelityDiagnostics, item => item.Code == "IMAGE_SOURCE_DECODE_FALLBACK" && item.LossKind == OfficeConversionLossKind.Omission);
        Assert.DoesNotContain(document.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_METAFILE_RASTERIZED");
    }

    [Fact]
    public void Application_metafile_projection_preserves_original_assets_and_reports_its_role() {
        string file = PublisherNativeTests.Fixture("SampleBrochure.pub");
        PublisherDocument baseline = PublisherDocument.Load(file);
        var codec = new TestCodec();
        PublisherDocument document = PublisherDocument.Load(file, new PublisherReadOptions { ImageCodec = codec });
        Assert.True(codec.Calls > 0);
        Assert.Contains(document.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_METAFILE_RASTERIZED" && item.LossKind == OfficeConversionLossKind.Approximation);
        Assert.DoesNotContain(document.ReadReport.FidelityDiagnostics, item => item.Code == "IMAGE_SOURCE_DECODE_FALLBACK");
        Assert.Equal(baseline.Images.Select(image => image.GetBytes()), document.Images.Select(image => image.GetBytes()), ByteArrayComparer.Instance);
        OfficeDrawingImage projected = PublisherNativeTests.Elements(document.Pages[0].Drawing).OfType<OfficeDrawingImage>().First();
        Assert.Equal("image/png", projected.ContentType);
        Assert.True(OfficeRasterImageDecoder.TryDecode(projected.Bytes, out OfficeRasterImage? raster));
        Assert.Equal(OfficeColor.FromRgb(20, 160, 40), raster!.GetPixel(0, 0));
        Assert.Equal(2, PdfReadDocument.Open(document.ToPdfBytes()).Pages.Count);
    }

    [Fact]
    public void Application_codec_output_obeys_pixel_limits_and_cancellation() {
        string file = PublisherNativeTests.Fixture("SampleBrochure.pub");
        Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(file, new PublisherReadOptions { ImageCodec = new TestCodec(), MaximumRasterPixels = 3 }));
        using var source = new CancellationTokenSource();
        Assert.Throws<OperationCanceledException>(() => PublisherDocument.Load(file,
            new PublisherReadOptions { ImageCodec = new TestCodec(source) }, source.Token));
    }

    private sealed class TestCodec : IOfficeRasterImageCodec {
        private readonly CancellationTokenSource? _cancel;
        internal TestCodec(CancellationTokenSource? cancel = null) => _cancel = cancel;
        internal int Calls { get; private set; }
        public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) {
            Calls++;
            encodedBytes[0] = 0; // The codec receives its own copy, never the retained native asset.
            _cancel?.Cancel();
            image = new OfficeRasterImage(2, 2, OfficeColor.FromRgb(20, 160, 40));
            return true;
        }
    }
    private sealed class ByteArrayComparer : IEqualityComparer<byte[]> {
        internal static readonly ByteArrayComparer Instance = new();
        public bool Equals(byte[]? first, byte[]? second) => first != null && second != null && first.SequenceEqual(second);
        public int GetHashCode(byte[] bytes) => bytes.Length;
    }
}
