using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeIMO.ConversionConsistency;

internal sealed record ConsistencySuite {
    public int SchemaVersion { get; init; } = 1;
    public required string FontPath { get; init; }
    public string FontFamily { get; init; } = "Consistency Sans";
    public int Dpi { get; init; } = 96;
    public required List<ConsistencyCase> Cases { get; init; }
}

internal sealed record ConsistencyCase {
    public required string Id { get; init; }
    public required string Format { get; init; }
    public required string Source { get; init; }
    public required string Evidence { get; init; }
    public string? ReferencePdf { get; init; }
    public bool RequireSearchablePdf { get; init; } = true;
    public bool ComparePdfPixels { get; init; } = true;
    public List<string> Limitations { get; init; } = new();
    public required List<PageExpectation> Pages { get; init; }
    public List<string> AllowedDiagnostics { get; init; } = new();
    public PixelTolerance VisualTolerance { get; init; } = new();
}

internal sealed record PageExpectation {
    public required int Width { get; init; }
    public required int Height { get; init; }
    public int? PdfWidth { get; init; }
    public int? PdfHeight { get; init; }
    public required List<string> Text { get; init; }
    public List<RegionExpectation> Regions { get; init; } = new();
    public List<TextPositionExpectation> TextPositions { get; init; } = new();
    public List<TextPositionExpectation> SvgTextPositions { get; init; } = new();
}

internal sealed record TextPositionExpectation(string Text, double X, double Baseline, double Tolerance = 0.25D);

internal sealed record RegionExpectation {
    public required string Id { get; init; }
    public required int X { get; init; }
    public required int Y { get; init; }
    public required int Width { get; init; }
    public required int Height { get; init; }
    public required string Color { get; init; }
    public int ColorTolerance { get; init; } = 48;
    public required int MinimumPixels { get; init; }
    public int MaximumPixels { get; init; } = int.MaxValue;
}

internal sealed record PixelTolerance {
    public int Channel { get; init; } = 32;
    public double DifferentRatio { get; init; } = 0.02D;
    public double MeanAbsoluteError { get; init; } = 3D;
}

internal sealed record EvidenceBundle(int SchemaVersion, string Commit, string SourceDiffSha256,
    string FontSha256, string FontFamily, int Dpi, string ShapingProvider, string Background,
    List<CaseBundle> Cases, List<SourceFileHash> UntrackedSourceFiles);
internal sealed record SourceFileHash(string Path, string Sha256);
internal sealed record CaseBundle(ConsistencyCase Contract, string SourceSha256, string PdfPath,
    string PdfSha256, string PdfRoute, List<ImageArtifact> Images, List<string> Diagnostics) {
    public List<string> DiagnosticDetails { get; init; } = new();
}
internal sealed record ImageArtifact(int Page, string Format, string Path, string Sha256, int Width, int Height);
internal sealed record GateReport(int SchemaVersion, string Commit, string FontSha256,
    string ExternalRasterizer, string BrowserVersion, bool Passed, List<CaseReport> Cases);
internal sealed record CaseReport(string Id, bool Passed, List<string> Errors, List<PageReport> Pages,
    string PdfRoute, bool FullVisualCoverage, bool SearchablePdfVerified, List<string> Limitations);
internal sealed record PageReport(int Page, bool Passed, List<string> Errors, List<ComparisonReport> Comparisons);
internal sealed record ComparisonReport(string Route, bool Passed, int DifferentPixels, int TotalPixels, double? MeanAbsoluteError, string DiffPath);

internal static class GateJson {
    internal static readonly JsonSerializerOptions Options = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        WriteIndented = true,
        UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow
    };
    internal static T Read<T>(string path) => JsonSerializer.Deserialize<T>(File.ReadAllText(path), Options)
        ?? throw new InvalidDataException("Empty JSON document: " + path);
    internal static void Write<T>(string path, T value) => File.WriteAllText(path, JsonSerializer.Serialize(value, Options) + "\n");
}
