using System.Security.Cryptography;
using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks;

/// <summary>
/// Untimed preparation and validation around public PDF operations. PowerForge owns all measurements.
/// </summary>
public sealed partial class PdfRuntimeWorkload {
    private readonly byte[] _source;
    private readonly PdfPageRegion _region;
    private readonly bool _textOnly;
    private readonly int _expectedAffectedSpans;
    private PdfDocumentInfo[]? _inspections;
    private IReadOnlyList<PdfPageRenderResult>? _renders;
    private byte[]? _output;
    private PdfTextEditResult? _edit;
    private PdfSerializationReport? _serialization;
    private string? _operation;
    private IndependentPage[]? _independentSource;

    /// <summary>Prepares one immutable PDF and a single ordinary text region on its middle page.</summary>
    public PdfRuntimeWorkload(int pages) : this(pages, PdfRuntimeCorpus.BusinessProfile) { }

    /// <summary>Prepares the named business or ordinary text editing input profile.</summary>
    public PdfRuntimeWorkload(int pages, string corpusProfile) {
        if (pages < 1 || pages > 500) throw new ArgumentOutOfRangeException(nameof(pages));
        _textOnly = corpusProfile switch {
            PdfRuntimeCorpus.BusinessProfile => false,
            PdfRuntimeCorpus.EditProfile => true,
            _ => throw new ArgumentException("Unknown PDF runtime corpus profile.", nameof(corpusProfile))
        };
        CorpusProfile = corpusProfile;
        Pages = pages;
        TargetPage = (pages + 1) / 2;
        _source = PdfRuntimeCorpus.Create(pages, PdfObjectSerializationMode.Buffered, _textOnly).ToBytes();
        SourceSha256 = Convert.ToHexString(SHA256.HashData(_source));
        PdfTextMatch match = PdfDocument.Load(_source).Text.Find(PdfRuntimeCorpus.Target(TargetPage),
            new PdfTextSearchOptions { MatchCase = true }).Single();
        _region = new PdfPageRegion(match.PageNumber, match.X, match.Y, match.Width, match.Height);
        _expectedAffectedSpans = PdfDocument.Load(_source).Text.Inspect(_region).Spans.Count;
    }

    /// <summary>Input page count, including pages not touched by a region edit.</summary>
    public int Pages { get; }
    /// <summary>One-based page whose region is edited.</summary>
    public int TargetPage { get; }
    /// <summary>Exact source profile; only the text-only profile qualifies ordinary region editing.</summary>
    public string CorpusProfile { get; }
    /// <summary>SHA-256 of the exact immutable input bytes.</summary>
    public string SourceSha256 { get; }
    /// <summary>Exact source byte count.</summary>
    public long InputBytes => _source.LongLength;
    /// <summary>Saved or rendered byte count, populated by validation.</summary>
    public long OutputBytes { get; private set; }
    /// <summary>Number of pages successfully checked after the operation.</summary>
    public int VerifiedPages { get; private set; }
    /// <summary>Source text spans reported affected by a region mutation, including spaces.</summary>
    public int AffectedTextSpans => _edit?.AffectedCount ?? 0;
    /// <summary>Writer-reported page payload retention, zero for other operations.</summary>
    public long PeakRetainedPageContentBytes => _serialization?.PeakRetainedPageContentBytes ?? 0L;
    /// <summary>Writer-reported completed object retention, zero for other operations.</summary>
    public long PeakRetainedObjectBytes => _serialization?.PeakRetainedObjectBytes ?? 0L;
    /// <summary>Writer-reported largest serialized object, zero for other operations.</summary>
    public long LargestSerializedObjectBytes => _serialization?.LargestSerializedObjectBytes ?? 0L;
    /// <summary>Whether this operation produced saved PDF bytes.</summary>
    public bool HasPdfOutput => _output is not null;
    /// <summary>Whether Poppler has qualified the prepared source and its geometry.</summary>
    public bool HasIndependentSource => _independentSource is not null;

    /// <summary>Releases previous results outside the measured operation.</summary>
    public void ReleaseResults() {
        _inspections = null;
        _renders = null;
        _output = null;
        _edit = null;
        _serialization = null;
        _operation = null;
        OutputBytes = 0L;
        VerifiedPages = 0;
    }

    /// <summary>Executes one complete caller operation; source preparation and readback are separate.</summary>
    public void Execute(string operation) {
        _operation = operation;
        switch (operation) {
            case "GenerateBuffered":
                Generate(PdfObjectSerializationMode.Buffered);
                break;
            case "GenerateForward":
                Generate(PdfObjectSerializationMode.ForwardOnly);
                break;
            case "InspectCold":
                _inspections = new PdfDocumentInfo[4];
                for (int index = 0; index < _inspections.Length; index++) {
                    _inspections[index] = PdfDocument.Load(_source).Inspect();
                }
                break;
            case "InspectReused":
                PdfDocument document = PdfDocument.Load(_source);
                _inspections = new PdfDocumentInfo[4];
                for (int index = 0; index < _inspections.Length; index++) {
                    _inspections[index] = document.Inspect();
                }
                break;
            case "RenderSvg":
                Render(global::OfficeIMO.Drawing.OfficeImageExportFormat.Svg, Math.Min(Pages, 12));
                break;
            case "RenderPng":
                Render(global::OfficeIMO.Drawing.OfficeImageExportFormat.Png, Math.Min(Pages, 4));
                break;
            case "ReplaceRegion":
                _edit = PdfDocument.Load(_source).Text.Replace(_region, PdfRuntimeCorpus.Replacement(TargetPage));
                _output = _edit.Document.ToBytes();
                break;
            case "MoveRegion":
                _edit = PdfDocument.Load(_source).Text.Move(_region, 8D, 0D);
                _output = _edit.Document.ToBytes();
                break;
            default:
                throw new ArgumentException("Unknown PDF runtime operation.", nameof(operation));
        }
    }

    private void Generate(PdfObjectSerializationMode mode) {
        PdfDocument document = PdfRuntimeCorpus.Create(Pages, mode, _textOnly);
        using var stream = new MemoryStream();
        PdfSaveResult result = document.SaveWithEvidence(stream);
        _serialization = result.Serialization ?? throw new InvalidDataException("Generation returned no serialization report.");
        _output = stream.ToArray();
    }

    private void Render(global::OfficeIMO.Drawing.OfficeImageExportFormat format, int count) {
        _renders = PdfDocument.Load(_source).Render.Pages("1-" + count,
            new PdfPageRenderOptions {
                Format = format,
                MaxPages = count,
                ThumbnailMaxDimension = format == global::OfficeIMO.Drawing.OfficeImageExportFormat.Png ? 512 : null,
                MaxPixelsPerPage = format == global::OfficeIMO.Drawing.OfficeImageExportFormat.Png ? 512L * 512L : 16_000_000L,
                ContinueOnError = false
            });
    }

    /// <summary>Writes the prepared source for independent extraction outside timing.</summary>
    public void ExportSource(string path) => File.WriteAllBytes(path, _source);

}
