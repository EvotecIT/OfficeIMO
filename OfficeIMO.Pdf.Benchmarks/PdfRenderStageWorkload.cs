using System.Security.Cryptography;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks;

/// <summary>
/// Prepared public rendering boundaries for one fixed business PDF. PowerForge owns timing and allocation measurement.
/// Separate stage measurements have different prepared inputs and are not additive pipeline timings.
/// </summary>
public sealed partial class PdfRenderStageWorkload {
    private readonly byte[] _source;
    private readonly OfficeRasterEncodingOptions _encodingOptions = new() {
        WriteResolutionMetadata = false,
        Png = new OfficePngEncodeOptions { WritePhysicalResolution = false }
    };
    private readonly List<OfficeImageExportDiagnostic> _rasterDiagnostics = new();
    private readonly OfficeDrawingRasterRenderOptions _rasterOptions;
    private readonly PdfPageRenderOptions _controlOptions;
    private PdfReadDocument? _reader;
    private OfficeDrawing[]? _drawings;
    private OfficeRasterImage[]? _rasters;
    private byte[][]? _pngs;
    private IReadOnlyList<PdfPageRenderResult>? _control;
    private string? _preparedOperation;
    private bool _executed;

    /// <summary>Generates the immutable twenty-page embedded-Carlito business source outside measurement.</summary>
    public PdfRenderStageWorkload() {
        _source = PdfRuntimeCorpus.Create(Pages, PdfObjectSerializationMode.Buffered).ToBytes();
        SourceSha256 = Convert.ToHexString(SHA256.HashData(_source));
        _rasterOptions = new OfficeDrawingRasterRenderOptions {
            Scale = Scale,
            Background = OfficeColor.White,
            ThrowOnImageDecodeFailure = true,
            MaximumRasterPixels = (long)PixelWidth * PixelHeight,
            DiagnosticSink = _rasterDiagnostics
        };
        _controlOptions = new PdfPageRenderOptions {
            Format = PdfPageRenderFormat.Png,
            Scale = Scale,
            Background = OfficeColor.White,
            MaxPages = PreviewPages,
            MaxPixelsPerPage = (long)PixelWidth * PixelHeight,
            ContinueOnError = false
        };
    }

    /// <summary>Complete source page count, including pages outside the preview selection.</summary>
    public int Pages => 20;
    /// <summary>One-based pages one through four are processed by each preview operation.</summary>
    public int PreviewPages => 4;
    /// <summary>Fixed preview scale relative to the 600 by 800 point PDF page.</summary>
    public double Scale => .64D;
    /// <summary>Expected raster width for every preview page.</summary>
    public int PixelWidth => 384;
    /// <summary>Expected raster height for every preview page.</summary>
    public int PixelHeight => 512;
    /// <summary>Existing corpus identity, shared with complete public runtime evidence.</summary>
    public string CorpusProfile => PdfRuntimeCorpus.BusinessProfile;
    /// <summary>SHA-256 of the exact immutable input PDF bytes.</summary>
    public string SourceSha256 { get; }
    /// <summary>Exact source byte count.</summary>
    public long InputBytes => _source.LongLength;
    /// <summary>Whether the exact source passed independent Poppler text and page-geometry validation.</summary>
    public bool HasIndependentSource { get; private set; }
    /// <summary>Complete source pages checked outside measurement.</summary>
    public int SourceVerifiedPages { get; private set; }
    /// <summary>Preview pages whose complete PNG bytes and pixels matched the public control.</summary>
    public int VerifiedPages { get; private set; }
    /// <summary>Complete preview PNG bytes, including downstream completion performed outside measurement.</summary>
    public long OutputBytes { get; private set; }
    /// <summary>RGBA bytes returned by rasterization, or prepared as the encoding input.</summary>
    public long RasterBytes { get; private set; }
    /// <summary>Aggregate root drawing elements in the four validated page scenes.</summary>
    public int ProjectionElements { get; private set; }
    /// <summary>Aggregate root text elements in the validated scenes.</summary>
    public int ProjectionTextElements { get; private set; }
    /// <summary>Aggregate root image elements in the validated scenes.</summary>
    public int ProjectionImageElements { get; private set; }
    /// <summary>Aggregate root shape elements in the validated scenes.</summary>
    public int ProjectionShapeElements { get; private set; }
    /// <summary>Capability diagnostics returned by the complete public control; validation requires zero.</summary>
    public int CapabilityDiagnostics { get; private set; }
    /// <summary>Fidelity diagnostics observed by the stage rasterizer; validation requires zero.</summary>
    public int RasterDiagnostics { get; private set; }
    /// <summary>Pages compared byte-for-byte with the public control.</summary>
    public int MatchedControlPages { get; private set; }
    /// <summary>Total decoded RGBA bytes compared with the public control outside measurement.</summary>
    public long ComparedPixelBytes { get; private set; }
    /// <summary>Exact retained input shape at the beginning of the measured operation.</summary>
    public string PreparedInputShape => _preparedOperation switch {
        "SourceOpen" or "RenderPng" => "Immutable PDF bytes; no retained reader, drawing, raster, or prior result.",
        "Projection" => "Immutable PDF bytes and one freshly opened reader; no page was projected during preparation.",
        "Rasterization" => "Immutable PDF bytes and four newly projected drawings with embedded fonts; no retained source reader or raster.",
        "PngEncoding" => "Immutable PDF bytes and four newly rendered RGBA rasters; no retained source reader or drawing.",
        _ => throw new InvalidOperationException("No rendering stage was prepared.")
    };
    /// <summary>Result retained by the measured boundary until validation or the next preparation.</summary>
    public string MeasuredOutputShape => _preparedOperation switch {
        "SourceOpen" => "One reader over all twenty source pages; text extraction and projection are outside measurement.",
        "Projection" => "Four drawings, including their resolved embedded fonts, from the prepared reader.",
        "Rasterization" => "Four RGBA rasters; prepared drawings remain retained.",
        "PngEncoding" => "Four complete PNG byte arrays; prepared rasters remain retained.",
        "RenderPng" => "Four complete public page-render results, including capability diagnostics and encoded bytes.",
        _ => throw new InvalidOperationException("No rendering stage was prepared.")
    };

    /// <summary>Releases previous prepared inputs and results outside measurement; source qualification is retained.</summary>
    public void ReleaseResults() {
        _reader = null;
        _drawings = null;
        _rasters = null;
        _pngs = null;
        _control = null;
        _preparedOperation = null;
        _executed = false;
        _rasterDiagnostics.Clear();
        SourceVerifiedPages = VerifiedPages = ProjectionElements = ProjectionTextElements = 0;
        ProjectionImageElements = ProjectionShapeElements = CapabilityDiagnostics = RasterDiagnostics = MatchedControlPages = 0;
        OutputBytes = RasterBytes = ComparedPixelBytes = 0L;
    }

    /// <summary>
    /// Prepares only the preceding public-boundary inputs for one measured call, without reusing a prior reader or result.
    /// Process-wide JIT, font, and codec caches are not reset; these are prepared-stage measurements, not process-cold claims.
    /// </summary>
    public void Prepare(string operation) {
        ReleaseResults();
        switch (operation) {
            case "SourceOpen":
            case "RenderPng":
                break;
            case "Projection":
                _reader = PdfReadDocument.Open(_source);
                break;
            case "Rasterization":
                _drawings = Project(PdfReadDocument.Open(_source));
                break;
            case "PngEncoding":
                _rasters = Rasterize(Project(PdfReadDocument.Open(_source)));
                break;
            default:
                throw new ArgumentException("Unknown PDF render-stage operation.", nameof(operation));
        }
        _preparedOperation = operation;
    }

    /// <summary>Executes one public boundary, or the complete public four-page PNG control, without measurement timers.</summary>
    public void Execute(string operation) {
        if (_preparedOperation is null || _preparedOperation != operation || _executed) throw new InvalidOperationException("Prepare the same stage before each measured call.");
        switch (operation) {
            case "SourceOpen":
                _reader = PdfReadDocument.Open(_source);
                break;
            case "Projection":
                _drawings = Project(_reader!);
                break;
            case "Rasterization":
                _rasters = Rasterize(_drawings!);
                break;
            case "PngEncoding":
                _pngs = Encode(_rasters!);
                break;
            case "RenderPng":
                _control = RenderControl();
                break;
        }
        _executed = true;
    }

    /// <summary>Writes the exact input bytes for independent validation outside measurement.</summary>
    public void ExportSource(string path) => File.WriteAllBytes(path, _source);

    /// <summary>Reuses the business corpus's independent source contract and checks the exact source identity.</summary>
    public void ValidateIndependentSource(string xmlPath) {
        var contract = new PdfRuntimeWorkload(Pages);
        if (contract.SourceSha256 != SourceSha256) throw new InvalidDataException("Independent source contract uses different PDF bytes.");
        contract.ValidateIndependentSource(xmlPath);
        HasIndependentSource = true;
    }

    private OfficeDrawing[] Project(PdfReadDocument reader) {
        var drawings = new OfficeDrawing[PreviewPages];
        for (int index = 0; index < drawings.Length; index++) drawings[index] = reader.Pages[index].ToDrawing();
        return drawings;
    }

    private OfficeRasterImage[] Rasterize(OfficeDrawing[] drawings) {
        var rasters = new OfficeRasterImage[PreviewPages];
        for (int index = 0; index < rasters.Length; index++) rasters[index] = OfficeDrawingRasterRenderer.Render(drawings[index], _rasterOptions);
        return rasters;
    }

    private byte[][] Encode(OfficeRasterImage[] rasters) {
        var pngs = new byte[PreviewPages][];
        for (int index = 0; index < pngs.Length; index++) pngs[index] = OfficeRasterImageEncoder.Encode(rasters[index], OfficeImageExportFormat.Png, _encodingOptions);
        return pngs;
    }

    private IReadOnlyList<PdfPageRenderResult> RenderControl() =>
        PdfDocument.Load(_source).Render.Pages("1-4", _controlOptions);
}
