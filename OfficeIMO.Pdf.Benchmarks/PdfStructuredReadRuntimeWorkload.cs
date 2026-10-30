using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks;

/// <summary>Fresh public structured reads over slightly drifting repeated headers.</summary>
public sealed class PdfStructuredReadRuntimeWorkload {
    private const string Header = "Running operations";
    private readonly byte[] _source;
    private readonly PdfLoadOptions _loadOptions = new() { Limits = PdfReadLimits.Default };
    private readonly PdfLoadOptions _readOptions = new() {
        Profile = PdfReadProfile.Structured, Pipeline = PdfUnderstandingPipelineOptions.Structured()
    };
    private PdfDocumentReadResult? _result;

    /// <summary>Prepares a fixed source outside measurement; page geometry uses top-left canvas coordinates.</summary>
    public PdfStructuredReadRuntimeWorkload(int pages) {
        if (pages < 2 || pages > 500) throw new ArgumentOutOfRangeException(nameof(pages));
        Pages = pages;
        _source = PdfDocument.Create(compose => {
            for (int page = 1; page <= pages; page++) {
                int number = page;
                compose.Page(layout => layout.Size(600D, 800D).Canvas(canvas => canvas
                    .Text(Header, global::OfficeIMO.OfficeLength.Points(36D + .02D * number), global::OfficeIMO.OfficeLength.Points(36D), global::OfficeIMO.OfficeLength.Points(300D), global::OfficeIMO.OfficeLength.Points(24D), fontSize: global::OfficeIMO.OfficeLength.Points(10D + .001D * number))
                    .Text(Body(number), global::OfficeIMO.OfficeLength.Points(36D), global::OfficeIMO.OfficeLength.Points(240D), global::OfficeIMO.OfficeLength.Points(420D), global::OfficeIMO.OfficeLength.Points(24D), fontSize: global::OfficeIMO.OfficeLength.Points(12D))));
            }
        }, new PdfCreateOptions { CompressContentStreams = true, FileVersion = PdfFileVersion.Pdf17 })
            .Meta(title: "OfficeIMO drifting header read corpus").ToBytes();
        SourceSha256 = Convert.ToHexString(SHA256.HashData(_source));
    }

    /// <summary>Number of source pages read by each fresh operation.</summary>
    public int Pages { get; }
    /// <summary>SHA-256 of the exact generated source PDF.</summary>
    public string SourceSha256 { get; }
    /// <summary>Exact immutable source byte count.</summary>
    public long InputBytes => _source.LongLength;
    /// <summary>Successfully validated pages, populated outside measurement.</summary>
    public int VerifiedPages { get; private set; }
    /// <summary>One validated repeated header per page.</summary>
    public int VerifiedHeaders { get; private set; }
    /// <summary>Complete body markers validated in document order.</summary>
    public int VerifiedBodyMarkers { get; private set; }
    /// <summary>Digest of the validated ordered semantic header/body sequence.</summary>
    public string SemanticSha256 { get; private set; } = string.Empty;
    /// <summary>Whether the source has passed independent Poppler extraction.</summary>
    public bool HasIndependentSource { get; private set; }
    /// <summary>Exact configured canonical parser limits, retained as provenance.</summary>
    public string ParserLimitsJson => JsonSerializer.Serialize(_loadOptions.Limits);
    /// <summary>Maximum pages admitted by the canonical semantic pipeline.</summary>
    public int MaximumReadPages => _readOptions.Pipeline!.MaxPages;
    /// <summary>Canonical per-page semantic work budget.</summary>
    public long MaximumPageWorkUnits => _readOptions.Pipeline!.MaxWorkUnitsPerPage;
    /// <summary>Canonical document enrichment work budget.</summary>
    public long MaximumDocumentWorkUnits => _readOptions.Pipeline!.MaxDocumentWorkUnits;

    /// <summary>Releases the previous read result before the next measured call.</summary>
    public void ReleaseResults() {
        _result = null;
        VerifiedPages = VerifiedHeaders = VerifiedBodyMarkers = 0;
        SemanticSha256 = string.Empty;
    }

    /// <summary>Loads the same bytes afresh and invokes the complete public structured read.</summary>
    public void Execute(CancellationToken cancellationToken = default) =>
        _result = PdfDocument.Load(_source, _loadOptions).Read(_readOptions, cancellationToken);

    /// <summary>Requires one header and exactly the expected ordered body on every page.</summary>
    public void Validate(string directory) {
        PdfDocumentReadResult result = _result ?? throw new InvalidOperationException("No structured read was executed.");
        if (SourceSha256 != Convert.ToHexString(SHA256.HashData(_source)) || result.Profile != PdfReadProfile.Structured ||
            result.SourcePageCount != Pages || result.Pages.Count != Pages) throw new InvalidDataException("Structured source/profile/page contract changed.");
        var sequence = new List<string>();
        for (int index = 0; index < Pages; index++) {
            PdfLogicalPage page = result.Pages[index];
            if (page.PageNumber != index + 1 || page.Width != 600D || page.Height != 800D ||
                page.Headers.Count != 1 || page.Headers[0].Text != Header || page.Footers.Count != 0) {
                throw new InvalidDataException("Repeated header classification or page order changed.");
            }
            string[] body = page.TextBlocks.Where(block => block.Kind is not (PdfLogicalElementKind.Header or PdfLogicalElementKind.Footer))
                .Select(block => block.Text).ToArray();
            if (body.Length != 1 || body[0] != Body(index + 1)) throw new InvalidDataException("Complete ordered body content changed.");
            sequence.Add(page.Headers[0].Text);
            sequence.Add(body[0]);
        }
        VerifiedPages = VerifiedHeaders = VerifiedBodyMarkers = Pages;
        SemanticSha256 = Convert.ToHexString(SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(string.Join('\n', sequence))));
        Directory.CreateDirectory(directory);
        File.WriteAllText(Path.Combine(directory, "structured-evidence.json"), JsonSerializer.Serialize(new {
            Contract = "pdf-structured-drift-v1", SourceSha256, InputBytes, Pages,
            VerifiedPages, VerifiedHeaders, VerifiedBodyMarkers, SemanticSha256,
            MaximumReadPages, MaximumPageWorkUnits, MaximumDocumentWorkUnits,
            HeaderX = "36 + .02 * one-based page", HeaderFontSize = "10 + .001 * one-based page", HeaderTop = 36,
            BodyTop = 240, OrderedContent = sequence
        }, new JsonSerializerOptions { WriteIndented = true }));
    }

    /// <summary>Writes the exact prepared source for independent validation outside measurement.</summary>
    public void ExportSource(string path) => File.WriteAllBytes(path, _source);

    /// <summary>Requires the complete header/body sequence and page geometry in independent extraction.</summary>
    public void ValidateIndependentSource(string xmlPath) {
        using var reader = XmlReader.Create(xmlPath, new XmlReaderSettings { DtdProcessing = DtdProcessing.Ignore, XmlResolver = null });
        XElement[] pages = XDocument.Load(reader).Descendants().Where(element => element.Name.LocalName == "page").ToArray();
        if (pages.Length != Pages) throw new InvalidDataException("Independent drifting-header page count changed.");
        for (int index = 0; index < Pages; index++) {
            XElement page = pages[index];
            string text = string.Join(' ', page.Descendants().Where(element => element.Name.LocalName == "word").Select(word => word.Value));
            if (double.Parse((string?)page.Attribute("width") ?? "NaN", CultureInfo.InvariantCulture) != 600D ||
                double.Parse((string?)page.Attribute("height") ?? "NaN", CultureInfo.InvariantCulture) != 800D ||
                text != Header + " " + Body(index + 1)) throw new InvalidDataException("Independent header/body content or order changed.");
        }
        HasIndependentSource = true;
    }

    private static string Body(int page) => "Body marker" + page.ToString("D4", CultureInfo.InvariantCulture);
}
