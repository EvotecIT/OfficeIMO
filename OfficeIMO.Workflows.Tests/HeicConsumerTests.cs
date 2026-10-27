using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.Workflows.Tests;

// These tests replace the process-wide platform codec and must run without other collections.
[CollectionDefinition("HEIC consumers", DisableParallelization = true)]
public sealed class HeicConsumerCollection { }

[Collection("HEIC consumers")]
public sealed class HeicConsumerTests {
    private static string Fixture => Path.Combine(AppContext.BaseDirectory, "Fixtures", "Heic", "basic-8bit-srgb.heic");

    [Theory]
    [InlineData(".docx")]
    [InlineData(".pptx")]
    [InlineData(".xlsx")]
    public void InsertsStorePngPartsInsteadOfHeic(string extension) {
        var codec = new SolidTileCodec();
        using var scope = new CodecScope(codec);
        string output = Path.Combine(scope.Directory, "inserted" + extension);
        byte[] heic = File.ReadAllBytes(Fixture);
        switch (extension) {
            case ".docx":
                using (WordDocument document = WordDocument.Create(output)) {
                    document.AddParagraph().AddImage(Fixture);
                    document.Save();
                }
                break;
            case ".pptx":
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(output)) {
                    using var input = new MemoryStream(heic, writable: false);
                    input.Position = 7;
                    presentation.AddSlide().AddPicture(input, OfficeImageFormat.Heic);
                    presentation.Save();
                }
                break;
            case ".xlsx":
                using (ExcelDocument document = ExcelDocument.Create(output)) {
                    document.AddWorksheet("Images").AddImage(1, 1, heic, "image/heic");
                    document.Save();
                }
                break;
        }

        Assert.True(codec.DecodeCalls > 0);
        using ZipArchive package = ZipFile.OpenRead(output);
        using Stream contentTypes = package.GetEntry("[Content_Types].xml")!.Open();
        XElement[] pngTypes = XDocument.Load(contentTypes).Descendants()
            .Where(element => (string?)element.Attribute("ContentType") == "image/png").ToArray();
        ZipArchiveEntry media = Assert.Single(package.Entries, entry => pngTypes.Any(type =>
            (string?)type.Attribute("PartName") == "/" + entry.FullName ||
            (string?)type.Attribute("Extension") == Path.GetExtension(entry.FullName).TrimStart('.')));
        using var stream = media.Open();
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        byte[] embedded = buffer.ToArray();
        Assert.Equal(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, embedded.Take(8).ToArray());
        Assert.True(OfficeRasterImageDecoder.TryDecode(embedded, out OfficeRasterImage? decoded));
        OfficeImageInfo original = OfficeImageReader.Identify(heic);
        Assert.Equal(original.Width, decoded!.Width);
        Assert.Equal(original.Height, decoded.Height);
    }

    [Theory]
    [InlineData(".heic")]
    [InlineData(".heif")]
    [InlineData(".hif")]
    public async Task WorkflowRoutesAllHeicAliasesToAReopenedPdf(string extension) {
        using var scope = new CodecScope(new SolidTileCodec());
        string source = Path.Combine(scope.Directory, "source" + extension);
        File.Copy(Fixture, source);
        string output = Path.Combine(scope.Directory, "converted.pdf");
        OfficeWorkflowRoute? route = OfficeWorkflowCatalog.Find(extension, ".pdf", executableOnly: true);
        Assert.NotNull(route);
        OfficeConversionCapability capability = Assert.Single(OfficeConversionCapabilityCatalog.All, item => item.Id == route.Id);
        Assert.Equal("OfficeIMO.Pdf", capability.PackageId);
        Assert.False(capability.BrowserAvailable);
        OfficeWorkflowResult result = await new OfficeWorkflowRunner().RunAsync(new OfficeWorkflowRequest {
            Operation = OfficeWorkflowOperation.Convert, InputPath = source, OutputPath = output,
            ConversionRouteId = route.Id,
            ConflictPolicy = OfficeWorkflowConflictPolicy.Fail
        });

        Assert.True(result.Succeeded, result.Summary);
        Assert.Equal(1, PdfDocument.Load(output).Inspect().PageCount);
        Assert.Contains(result.Diagnostics, item => item.Code == "OutputReopened");
    }

    [Theory]
    [InlineData(OfficeRasterCodecUnavailableReason.NoPlatformCodec)]
    [InlineData(OfficeRasterCodecUnavailableReason.WindowsHeifExtensionMissing)]
    [InlineData(OfficeRasterCodecUnavailableReason.WindowsHevcExtensionMissing)]
    public async Task ConversionAndAssemblyKeepCodecUnavailableReasonsAndPublishNothing(OfficeRasterCodecUnavailableReason reason) {
        using var scope = new CodecScope(reason == OfficeRasterCodecUnavailableReason.NoPlatformCodec ? null : new UnavailableCodec(reason));
        string source = Path.Combine(scope.Directory, "source.hif");
        File.Copy(Fixture, source);
        var runner = new OfficeWorkflowRunner();
        string converted = Path.Combine(scope.Directory, "converted.pdf");
        OfficeWorkflowResult conversion = await runner.RunAsync(new OfficeWorkflowRequest {
            Operation = OfficeWorkflowOperation.Convert, InputPath = source, OutputPath = converted,
            ConversionRouteId = "heic-pdf",
            ConflictPolicy = OfficeWorkflowConflictPolicy.Fail
        });
        string assembled = Path.Combine(scope.Directory, "assembled.pdf");
        PdfAssemblyResult assembly = await runner.AssemblePdfAsync(new PdfAssemblyRequest {
            Sources = new[] { source }, OutputPath = assembled, ConflictPolicy = OfficeWorkflowConflictPolicy.Fail
        });

        Assert.False(conversion.Succeeded);
        Assert.False(assembly.Succeeded);
        Assert.Equal(OfficeWorkflowFailureKind.UnsupportedInput, conversion.FailureKind);
        Assert.Equal(OfficeWorkflowFailureKind.UnsupportedInput, assembly.FailureKind);
        foreach (IReadOnlyList<OfficeWorkflowDiagnostic> diagnostics in new[] { conversion.Diagnostics, assembly.Diagnostics }) {
            OfficeWorkflowDiagnostic diagnostic = Assert.Single(diagnostics, item => item.Code == "image-codec-unavailable");
            Assert.Equal(reason.ToString(), diagnostic.Details["reason"]);
        }
        Assert.False(File.Exists(converted));
        Assert.False(File.Exists(assembled));
        Assert.Empty(System.IO.Directory.GetFiles(scope.Directory, ".*.tmp"));
    }

    [Fact]
    public void TypedCodecUnavailableFailuresKeepTheirClassificationAtEveryWorkflowStage() {
        using var scope = new CodecScope(null);
        OfficeRasterDecodeException failure = Assert.Throws<OfficeRasterDecodeException>(() =>
            OfficeImageReader.NormalizeForEmbedding(File.ReadAllBytes(Fixture), null, out _));
        foreach (OfficeWorkflowRunner.WorkflowFailureStage stage in Enum.GetValues<OfficeWorkflowRunner.WorkflowFailureStage>())
            Assert.Equal(OfficeWorkflowFailureKind.UnsupportedInput, OfficeWorkflowRunner.ClassifyFailure(failure, stage));
    }

    [Fact]
    public void PreparedPdfImagesDoNotBypassTheCurrentHeicCodecAvailability() {
        using var scope = new CodecScope(new SolidTileCodec());
        byte[] heic = File.ReadAllBytes(Fixture);
        byte[] pdf = PdfDocument.CreateFromImages(new[] { new PdfImageDocumentSource(heic, "photo.heic") }).ToBytes();
        Assert.Equal(1, PdfDocument.Load(pdf).Inspect().PageCount);

        OfficeRasterImageCodecs.Platform = null;
        OfficeRasterDecodeException failure = Assert.Throws<OfficeRasterDecodeException>(() => {
            PdfDocument.CreateFromImages(new[] { new PdfImageDocumentSource(heic, "photo.heic") }).ToBytes();
        });
        Assert.Equal(OfficeRasterCodecUnavailableReason.NoPlatformCodec, failure.DecodeInfo.CodecUnavailableReason);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PreparedSvgImagesDoNotBypassEmbeddedHeicCodecAvailability(bool copySource) {
        using var scope = new CodecScope(new SolidTileCodec());
        string dataUri = "data:image/heic;base64," + Convert.ToBase64String(File.ReadAllBytes(Fixture));
        byte[] svg = System.Text.Encoding.UTF8.GetBytes(
            $"<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"160\" height=\"120\"><image width=\"160\" height=\"120\" href=\"{dataUri}\"/></svg>");
        byte[] pdf = PdfDocument.CreateFromImages(new[] { new PdfImageDocumentSource(svg, "photo.svg") }).ToBytes();
        Assert.Equal(1, PdfDocument.Load(pdf).Inspect().PageCount);
        OfficeRasterImageCodecs.Platform = null;
        OfficeRasterDecodeException failure = Assert.Throws<OfficeRasterDecodeException>(() => {
            PdfDocument.CreateFromImages(new[] { new PdfImageDocumentSource(copySource ? (byte[])svg.Clone() : svg, "photo.svg") }).ToBytes();
        });
        Assert.Equal(OfficeRasterCodecUnavailableReason.NoPlatformCodec, failure.DecodeInfo.CodecUnavailableReason);
    }

    private sealed class SolidTileCodec : IOfficeRasterImageCodec {
        internal int DecodeCalls { get; private set; }
        public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) {
            DecodeCalls++;
            OfficeImageInfo info = OfficeImageReader.Identify(encodedBytes);
            image = new OfficeRasterImage(info.Width, info.Height, OfficeColor.FromRgb(24, 100, 176));
            return true;
        }
    }

    private sealed class UnavailableCodec(OfficeRasterCodecUnavailableReason reason) : IOfficeRasterImageCodec, IOfficeRasterImageCodecAvailability {
        public OfficeRasterCodecAvailability GetAvailability() => new(reason, "The requested system component is unavailable.");
        public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) =>
            throw new InvalidOperationException("Unavailable codecs must not be invoked.");
    }

    private sealed class CodecScope : IDisposable {
        private readonly IOfficeRasterImageCodec? _previous;
        internal CodecScope(IOfficeRasterImageCodec? codec) {
            // Complete startup registration before substituting the external codec boundary.
            _ = new OfficeWorkflowRunner();
            _previous = OfficeRasterImageCodecs.Platform;
            OfficeRasterImageCodecs.Platform = codec;
            Directory = Path.Combine(Path.GetTempPath(), "officeimo-heic-consumers-" + Guid.NewGuid().ToString("N"));
            System.IO.Directory.CreateDirectory(Directory);
        }
        internal string Directory { get; }
        public void Dispose() {
            OfficeRasterImageCodecs.Platform = _previous;
            System.IO.Directory.Delete(Directory, recursive: true);
        }
    }
}
