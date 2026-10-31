using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks;

/// <summary>Distinct geometry contracts for the same independent ledger input and ordered cell content.</summary>
public enum PdfClosedLedgerExpectedGeometry {
    /// <summary>The established text-derived column boundaries and baseline bounds.</summary>
    DetectedProportions,
    /// <summary>The independently authored closed ruling stroke centres and enclosing grid bounds.</summary>
    ClosedRuling
}

/// <summary>Fresh public structured reads over the exact independent ReportLab ledger fixtures.</summary>
public sealed class PdfClosedLedgerRuntimeWorkload {
    private const string OracleSha256 = "138662292a8b97dd5920d991424d6b0d61f0149d0811f7bf458fb3eff199dbf7";
    private const string RulingEvidence = global::OfficeIMO.OfficeDiagnosticCodes.TableClosedRulingGrid;
    private const double GeometryTolerancePoints = 0.000001D;
    private static readonly double[] DetectedBoundaries = { 50D, 216.338D, 380.338D, 480.338D, 562D };
    private readonly byte[] _source;
    private readonly byte[] _manifest;
    private readonly string[][] _expectedRows;
    private readonly string[] _readingOrder;
    private readonly double[] _authoredBoundaries;
    private readonly double[] _rowBoundaries;
    private readonly double _baselineOffset;
    private readonly PdfLoadOptions _loadOptions = new() { Limits = PdfReadLimits.Default };
    private readonly PdfLoadOptions _readOptions = new() {
        Profile = PdfReadProfile.Structured, Pipeline = PdfUnderstandingPipelineOptions.Structured()
    };
    private PdfDocumentReadResult? _result;

    /// <summary>Loads and fingerprints a supported fixture and its checked-in independent manifest outside measurement.</summary>
    public PdfClosedLedgerRuntimeWorkload(string sourcePath, PdfClosedLedgerExpectedGeometry expectedGeometry) {
        ArgumentException.ThrowIfNullOrWhiteSpace(sourcePath);
        if (!Enum.IsDefined(expectedGeometry)) throw new ArgumentOutOfRangeException(nameof(expectedGeometry));
        string fullPath = Path.GetFullPath(sourcePath);
        FixtureId = Path.GetFileName(fullPath) switch {
            "ruled-ledger-0-native.pdf" => "ruled-ledger-0",
            "ruled-ledger-90-native.pdf" => "ruled-ledger-90",
            _ => throw new ArgumentException("Select a checked-in ruled-ledger-0 or ruled-ledger-90 native PDF.", nameof(sourcePath))
        };
        ExpectedGeometry = expectedGeometry;
        _manifest = File.ReadAllBytes(Path.Combine(Path.GetDirectoryName(fullPath)!, "manifest.json"));
        ManifestSha256 = Fingerprint(_manifest);
        if (!string.Equals(ManifestSha256, OracleSha256, StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("The independent ledger manifest differs from the qualified checked-in manifest.");
        using JsonDocument manifest = JsonDocument.Parse(_manifest);
        JsonElement label = manifest.RootElement.GetProperty("cases").EnumerateArray()
            .Single(item => item.GetProperty("id").GetString() == FixtureId);
        _source = File.ReadAllBytes(fullPath);
        SourceSha256 = Fingerprint(_source);
        if (!string.Equals(SourceSha256, label.GetProperty("files").GetProperty("native.pdf").GetString(), StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("The ledger PDF does not match its independent producer manifest.");
        Rotation = label.GetProperty("clockwiseDegrees").GetInt32();
        _expectedRows = label.GetProperty("table").EnumerateArray()
            .Select(row => row.EnumerateArray().Select(cell => cell.GetString()!).ToArray()).ToArray();
        _readingOrder = label.GetProperty("readingOrder").EnumerateArray().Select(row => row.GetString()!).ToArray();
        JsonElement geometry = label.GetProperty("geometry");
        _authoredBoundaries = Numbers(geometry, "columnBoundariesPoints");
        _rowBoundaries = Numbers(geometry, "rowBoundariesPoints");
        _baselineOffset = geometry.GetProperty("textBaselineOffsetPoints").GetDouble();
    }

    /// <summary>Independent producer case selected by the exact fixture filename.</summary>
    public string FixtureId { get; }
    /// <summary>Inherited PDF page rotation; the native content coordinates are unchanged.</summary>
    public int Rotation { get; }
    /// <summary>Geometry contract selected explicitly for this lane.</summary>
    public PdfClosedLedgerExpectedGeometry ExpectedGeometry { get; }
    /// <summary>SHA-256 of the immutable source PDF, verified against the producer manifest.</summary>
    public string SourceSha256 { get; }
    /// <summary>SHA-256 of the exact independent producer manifest used for validation.</summary>
    public string ManifestSha256 { get; }
    /// <summary>Exact source byte count loaded by every measured operation.</summary>
    public long InputBytes => _source.LongLength;
    /// <summary>Expected source page count for either checked-in fixture.</summary>
    public int PageCount => 1;
    /// <summary>Page count qualified outside measurement.</summary>
    public int VerifiedPages { get; private set; }
    /// <summary>Table count qualified outside measurement.</summary>
    public int VerifiedTables { get; private set; }
    /// <summary>Complete ordered cell count qualified outside measurement.</summary>
    public int VerifiedCells { get; private set; }
    /// <summary>Digest of ordered cell text only; geometry is deliberately excluded.</summary>
    public string SemanticSha256 { get; private set; } = string.Empty;
    /// <summary>Digest of the exact exported page/table output, including observed geometry.</summary>
    public string OutputSha256 { get; private set; } = string.Empty;
    /// <summary>Byte count of the compact exported page/table output.</summary>
    public long OutputBytes { get; private set; }
    /// <summary>Whether independent Poppler extraction qualified every source word in manifest order.</summary>
    public bool HasIndependentSource { get; private set; }
    /// <summary>Canonical parser resource limits retained as run provenance.</summary>
    public string ParserLimitsJson => JsonSerializer.Serialize(_loadOptions.Limits);
    /// <summary>Canonical semantic page limit.</summary>
    public int MaximumReadPages => _readOptions.Pipeline!.MaxPages;
    /// <summary>Canonical per-page semantic work budget.</summary>
    public long MaximumPageWorkUnits => _readOptions.Pipeline!.MaxWorkUnitsPerPage;
    /// <summary>Canonical document enrichment work budget.</summary>
    public long MaximumDocumentWorkUnits => _readOptions.Pipeline!.MaxDocumentWorkUnits;

    /// <summary>Releases prior output and validation counters outside measurement.</summary>
    public void ReleaseResults() {
        _result = null;
        VerifiedPages = VerifiedTables = VerifiedCells = 0;
        SemanticSha256 = OutputSha256 = string.Empty;
        OutputBytes = 0;
    }

    /// <summary>Loads the exact source bytes afresh and runs the complete public structured read with canonical limits.</summary>
    public void Execute() => _result = PdfDocument.Load(_source, _loadOptions).Read(_readOptions, CancellationToken.None);

    /// <summary>Qualifies complete ordered cells and the selected geometry, then exports exact input/output evidence outside measurement.</summary>
    public void Validate(string directory) {
        PdfDocumentReadResult result = _result ?? throw new InvalidOperationException("No closed-ledger read was executed.");
        if (Fingerprint(_source) != SourceSha256 || result.Profile != PdfReadProfile.Structured ||
            result.SourcePageCount != PageCount || result.Pages.Count != PageCount || result.Tables.Count != 1)
            throw new InvalidDataException("Closed-ledger source/profile/page/table contract changed.");
        PdfLogicalPage page = result.Pages[0];
        if (page.PageNumber != 1 || page.Width != 612D || page.Height != 792D || page.RotationDegrees != Rotation || page.Tables.Count != 1)
            throw new InvalidDataException("Closed-ledger page geometry or table placement changed.");
        PdfLogicalTable table = page.Tables[0];
        if (table.PageNumber != 1 || table.SourceKind != PdfLogicalContentSourceKind.Native ||
            table.CoordinateSpace != PdfTableCoordinateSpace.PdfUserSpace || table.Rows.Count != 4 || table.Columns.Count != 4 || table.Cells.Count != 16)
            throw new InvalidDataException("Closed-ledger native 4x4 table contract changed.");
        for (int row = 0; row < 4; row++) {
            if (!table.Rows[row].SequenceEqual(_expectedRows[row], StringComparer.Ordinal) ||
                string.Join(' ', table.Rows[row]) != _readingOrder[row])
                throw new InvalidDataException("Complete closed-ledger row text or order changed.");
            for (int column = 0; column < 4; column++) {
                PdfLogicalTableCell cell = table.Cells[row * 4 + column];
                if (cell.PageNumber != 1 || cell.RowIndex != row || cell.ColumnIndex != column ||
                    cell.Text != _expectedRows[row][column] || cell.Column is null)
                    throw new InvalidDataException("Closed-ledger indexed cell text or order changed.");
                RequireCoordinate(cell.Column.From, table.Columns[column].From);
                RequireCoordinate(cell.Column.To, table.Columns[column].To);
            }
        }
        bool closedRuling = ExpectedGeometry == PdfClosedLedgerExpectedGeometry.ClosedRuling;
        double[] boundaries = closedRuling ? _authoredBoundaries : DetectedBoundaries;
        double top = closedRuling ? _rowBoundaries[0] : _rowBoundaries[0] - _baselineOffset;
        double bottom = closedRuling ? _rowBoundaries[^1] : _rowBoundaries[^2] - _baselineOffset;
        RequireCoordinate(table.YTop, top);
        RequireCoordinate(table.YBottom, bottom);
        for (int column = 0; column < 4; column++) {
            RequireCoordinate(table.Columns[column].From, boundaries[column]);
            RequireCoordinate(table.Columns[column].To, boundaries[column + 1]);
        }
        if (table.Evidence.Any(item => item.Code == RulingEvidence) != closedRuling)
            throw new InvalidDataException("Closed-ruling admission evidence disagrees with the selected geometry mode.");
        VerifiedPages = VerifiedTables = 1;
        VerifiedCells = 16;
        SemanticSha256 = Fingerprint(JsonSerializer.SerializeToUtf8Bytes(table.Rows));
        byte[] output = JsonSerializer.SerializeToUtf8Bytes(new {
            page.PageNumber, page.Width, page.Height, page.RotationDegrees,
            Table = new {
                table.PageNumber, table.DetectionKind, SourceKind = table.SourceKind.ToString(), CoordinateSpace = table.CoordinateSpace.ToString(),
                Bounds = new { Left = table.Columns[0].From, Bottom = table.YBottom, Right = table.Columns[^1].To, Top = table.YTop },
                table.VisualBounds,
                Columns = table.Columns.Select(column => new { column.From, column.To, Width = column.To - column.From, column.VisualBounds }),
                table.Rows,
                Cells = table.Cells.Select(cell => new { cell.PageNumber, cell.RowIndex, cell.ColumnIndex, cell.Text }),
                EvidenceCodes = table.Evidence.Select(item => item.Code)
            }
        });
        OutputBytes = output.LongLength;
        OutputSha256 = Fingerprint(output);
        Directory.CreateDirectory(directory);
        File.WriteAllBytes(Path.Combine(directory, "source.pdf"), _source);
        File.WriteAllBytes(Path.Combine(directory, "oracle-manifest.json"), _manifest);
        File.WriteAllBytes(Path.Combine(directory, "table-output.json"), output);
        File.WriteAllText(Path.Combine(directory, "closed-ledger-evidence.json"), JsonSerializer.Serialize(new {
            Contract = "pdf-closed-ledger-structured-read-v1", FixtureId, Rotation, SourceSha256, ManifestSha256, InputBytes, PageCount,
            ExpectedGeometry = ExpectedGeometry.ToString(), VerifiedPages, VerifiedTables, VerifiedCells, SemanticSha256, OutputSha256, OutputBytes,
            HasIndependentSource, IndependentSourceContract = "Complete manifest-ordered page words, ordered by Poppler word coordinates normalized for /Rotate 0 or 90; independent grid geometry is not asserted.",
            GeometryTolerancePoints, ExpectedColumnBoundariesPoints = boundaries, ExpectedTopPoints = top, ExpectedBottomPoints = bottom,
            AuthoredColumnBoundariesPoints = _authoredBoundaries, AuthoredRowBoundariesPoints = _rowBoundaries,
            GeometryComparison = "Distinct detected-proportion and authored closed-ruling contracts; output geometry equality is not asserted.",
            ParserLimits = _loadOptions.Limits, MaximumReadPages, MaximumPageWorkUnits, MaximumDocumentWorkUnits,
            Cancellation = "CancellationToken.None", ReadProfile = "Structured"
        }));
    }

    /// <summary>Exports the exact fingerprinted source for an independent extractor outside measurement.</summary>
    public void ExportSource(string path) => File.WriteAllBytes(path, _source);

    /// <summary>Requires all source words, including duplicates and numeric punctuation, in independent manifest order.</summary>
    public void ValidateIndependentSource(string xmlPath) {
        using var reader = XmlReader.Create(xmlPath, new XmlReaderSettings { DtdProcessing = DtdProcessing.Ignore, XmlResolver = null });
        XElement[] pages = XDocument.Load(reader).Descendants().Where(element => element.Name.LocalName == "page").ToArray();
        if (pages.Length != PageCount) throw new InvalidDataException("Independent closed-ledger page count changed.");
        // Poppler's XML flow order changes on /Rotate 90. Recover the fixture's native
        // row/column progression from independent word boxes rather than accepting a bag of words.
        var words = pages[0].Descendants().Where(element => element.Name.LocalName == "word")
            .Select(word => new {
                Text = word.Value,
                Row = Rotation == 90 ? -WordCoordinate(word, "xMin") : WordCoordinate(word, "yMin"),
                Column = Rotation == 90 ? WordCoordinate(word, "yMin") : WordCoordinate(word, "xMin")
            }).OrderBy(word => word.Row).ThenBy(word => word.Column);
        string actual = string.Join(' ', words.Select(word => word.Text));
        if (actual != string.Join(' ', _readingOrder)) throw new InvalidDataException("Independent closed-ledger complete word order changed.");
        HasIndependentSource = true;
    }

    private static string Fingerprint(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes));
    private static double[] Numbers(JsonElement value, string name) => value.GetProperty(name).EnumerateArray().Select(item => item.GetDouble()).ToArray();
    private static double WordCoordinate(XElement word, string attribute) {
        if (!double.TryParse((string?)word.Attribute(attribute), NumberStyles.Float, CultureInfo.InvariantCulture, out double value) || !double.IsFinite(value))
            throw new InvalidDataException("Independent ledger word coordinates are missing or invalid.");
        return value;
    }
    private static void RequireCoordinate(double actual, double expected) {
        if (!double.IsFinite(actual) || Math.Abs(actual - expected) > GeometryTolerancePoints)
            throw new InvalidDataException($"Closed-ledger geometry changed: expected {expected:R}, observed {actual:R} points.");
    }
}
