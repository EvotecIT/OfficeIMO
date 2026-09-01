using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

/// <summary>Runs reusable PDF output and heterogeneous intake workflows.</summary>
public interface IOfficeOutputWorkflowRunner {
    /// <summary>Exports selected PDF pages into a validated image folder.</summary>
    Task<PdfPageImageExportResult> ExportPdfPagesAsync(
        PdfPageImageExportRequest request,
        IProgress<OfficeWorkflowProgress>? progress = null,
        CancellationToken cancellationToken = default);

    /// <summary>Normalizes supported inputs and assembles them into one validated PDF.</summary>
    Task<PdfAssemblyResult> AssemblePdfAsync(
        PdfAssemblyRequest request,
        IProgress<OfficeWorkflowProgress>? progress = null,
        CancellationToken cancellationToken = default);
}

/// <summary>Request to export selected PDF pages into one newly published folder.</summary>
public sealed class PdfPageImageExportRequest {
    /// <summary>Caller-provided request identifier.</summary>
    public string Id { get; set; } = Guid.NewGuid().ToString("N");

    /// <summary>Source PDF file.</summary>
    public required string InputPath { get; set; }

    /// <summary>Requested output folder. The complete folder is staged and then published.</summary>
    public required string OutputDirectory { get; set; }

    /// <summary>Document-relative selection such as <c>1-3,last</c>; all pages when omitted.</summary>
    public string? Pages { get; set; }

    /// <summary>Output format.</summary>
    public OfficeImageExportFormat Format { get; set; } = OfficeImageExportFormat.Png;

    /// <summary>Target raster density. SVG records the requested logical scale without rasterizing.</summary>
    public double TargetDpi { get; set; } = 144D;

    /// <summary>Optional maximum output width or height in pixels.</summary>
    public int? MaximumDimension { get; set; }

    /// <summary>Maximum selected page count.</summary>
    public int MaximumPages { get; set; } = 100;

    /// <summary>Optional source password.</summary>
    public string? PdfPassword { get; set; }

    /// <summary>How an existing output folder is handled.</summary>
    public OfficeWorkflowConflictPolicy ConflictPolicy { get; set; } = OfficeWorkflowConflictPolicy.Rename;

    /// <summary>Shared input and aggregate output limits.</summary>
    public OfficeWorkflowLimits Limits { get; set; } = new();
}

/// <summary>One committed page-image artifact.</summary>
public sealed class PdfPageImageFile {
    internal PdfPageImageFile(
        int pageNumber,
        string path,
        OfficeImageExportFormat format,
        int width,
        int height,
        long sizeBytes) {
        PageNumber = pageNumber;
        Path = path;
        Format = format;
        Width = width;
        Height = height;
        SizeBytes = sizeBytes;
    }

    /// <summary>One-based source page number.</summary>
    public int PageNumber { get; }
    /// <summary>Published image path.</summary>
    public string Path { get; }
    /// <summary>Encoded image format.</summary>
    public OfficeImageExportFormat Format { get; }
    /// <summary>Encoded width.</summary>
    public int Width { get; }
    /// <summary>Encoded height.</summary>
    public int Height { get; }
    /// <summary>Encoded byte count.</summary>
    public long SizeBytes { get; }
}

/// <summary>Terminal result of one PDF page-image export.</summary>
public sealed class PdfPageImageExportResult {
    internal PdfPageImageExportResult(
        string requestId,
        OfficeWorkflowStatus status,
        OfficeWorkflowFailureKind failureKind,
        string? outputDirectory,
        long inputBytes,
        long outputBytes,
        TimeSpan duration,
        string summary,
        IReadOnlyList<PdfPageImageFile> files,
        IReadOnlyList<OfficeWorkflowDiagnostic> diagnostics) {
        RequestId = requestId;
        Status = status;
        FailureKind = failureKind;
        OutputDirectory = outputDirectory;
        InputBytes = inputBytes;
        OutputBytes = outputBytes;
        Duration = duration;
        Summary = summary;
        Files = files.ToArray();
        Diagnostics = diagnostics.ToArray();
    }

    /// <summary>Caller-provided request identifier.</summary>
    public string RequestId { get; }
    /// <summary>Terminal state.</summary>
    public OfficeWorkflowStatus Status { get; }
    /// <summary>Stable failure category, or <see cref="OfficeWorkflowFailureKind.None"/> when no failure occurred.</summary>
    public OfficeWorkflowFailureKind FailureKind { get; }
    /// <summary>Published output folder.</summary>
    public string? OutputDirectory { get; }
    /// <summary>Source PDF size.</summary>
    public long InputBytes { get; }
    /// <summary>Aggregate encoded image size.</summary>
    public long OutputBytes { get; }
    /// <summary>Total duration.</summary>
    public TimeSpan Duration { get; }
    /// <summary>User-facing outcome.</summary>
    public string Summary { get; }
    /// <summary>Published page images.</summary>
    public IReadOnlyList<PdfPageImageFile> Files { get; }
    /// <summary>Structured diagnostics.</summary>
    public IReadOnlyList<OfficeWorkflowDiagnostic> Diagnostics { get; }
    /// <summary>Whether export completed successfully.</summary>
    public bool Succeeded => Status == OfficeWorkflowStatus.Completed;
}

/// <summary>Safety and expansion options for heterogeneous PDF assembly.</summary>
public sealed class PdfAssemblyOptions {
    /// <summary>Includes nested files when a source is a folder.</summary>
    public bool IncludeSubdirectories { get; set; } = true;

    /// <summary>Ignores unsupported files discovered inside folders and archives.</summary>
    public bool IgnoreDiscoveredUnsupportedFiles { get; set; } = true;

    /// <summary>Maximum normalized source count.</summary>
    public int MaximumSourceCount { get; set; } = 250;

    /// <summary>Maximum aggregate folder files and ZIP entries inspected before filtering.</summary>
    public int MaximumDiscoveredEntries { get; set; } = 10_000;

    /// <summary>Maximum number of entries inspected in one ZIP source.</summary>
    public int MaximumArchiveEntries { get; set; } = 1_000;

    /// <summary>Maximum uncompressed byte count for one ZIP entry.</summary>
    public long MaximumArchiveEntryBytes { get; set; } = 256L * 1024L * 1024L;

    /// <summary>Maximum aggregate uncompressed byte count across all ZIP sources.</summary>
    public long MaximumArchiveBytes { get; set; } = 512L * 1024L * 1024L;

    /// <summary>Maximum accepted uncompressed-to-compressed ratio for one non-empty ZIP entry.</summary>
    public double MaximumArchiveCompressionRatio { get; set; } = 500D;

    /// <summary>Image-to-page settings.</summary>
    public PdfImageDocumentOptions ImageOptions { get; set; } = new();

    internal PdfAssemblyOptions CloneAndValidate() {
        if (MaximumSourceCount < 1) throw new ArgumentOutOfRangeException(nameof(MaximumSourceCount));
        if (MaximumDiscoveredEntries < 1) throw new ArgumentOutOfRangeException(nameof(MaximumDiscoveredEntries));
        if (MaximumArchiveEntries < 1) throw new ArgumentOutOfRangeException(nameof(MaximumArchiveEntries));
        if (MaximumArchiveEntryBytes < 1) throw new ArgumentOutOfRangeException(nameof(MaximumArchiveEntryBytes));
        if (MaximumArchiveBytes < 1) throw new ArgumentOutOfRangeException(nameof(MaximumArchiveBytes));
        if (MaximumArchiveCompressionRatio < 1D || double.IsNaN(MaximumArchiveCompressionRatio) || double.IsInfinity(MaximumArchiveCompressionRatio)) {
            throw new ArgumentOutOfRangeException(nameof(MaximumArchiveCompressionRatio));
        }
        if (ImageOptions == null) throw new ArgumentException("Image options cannot be null.", nameof(ImageOptions));
        return new PdfAssemblyOptions {
            IncludeSubdirectories = IncludeSubdirectories,
            IgnoreDiscoveredUnsupportedFiles = IgnoreDiscoveredUnsupportedFiles,
            MaximumSourceCount = MaximumSourceCount,
            MaximumDiscoveredEntries = MaximumDiscoveredEntries,
            MaximumArchiveEntries = MaximumArchiveEntries,
            MaximumArchiveEntryBytes = MaximumArchiveEntryBytes,
            MaximumArchiveBytes = MaximumArchiveBytes,
            MaximumArchiveCompressionRatio = MaximumArchiveCompressionRatio,
            ImageOptions = new PdfImageDocumentOptions {
                FixedPageSize = ImageOptions.FixedPageSize,
                FallbackPageSize = ImageOptions.FallbackPageSize,
                AutoOrientPage = ImageOptions.AutoOrientPage,
                Margin = ImageOptions.Margin,
                Fit = ImageOptions.Fit,
                MaximumPageDimension = ImageOptions.MaximumPageDimension
            }
        };
    }
}

/// <summary>Request to assemble PDFs, images, Office documents, folders, and ZIPs into one PDF.</summary>
public sealed class PdfAssemblyRequest {
    /// <summary>Caller-provided request identifier.</summary>
    public string Id { get; set; } = Guid.NewGuid().ToString("N");

    /// <summary>Ordered files, folders, or ZIP archives.</summary>
    public required IReadOnlyList<string> Sources { get; set; }

    /// <summary>Requested output PDF.</summary>
    public required string OutputPath { get; set; }

    /// <summary>How an existing output path is handled.</summary>
    public OfficeWorkflowConflictPolicy ConflictPolicy { get; set; } = OfficeWorkflowConflictPolicy.Rename;

    /// <summary>Output intent used while normalizing Office inputs.</summary>
    public OfficeWorkflowOutputProfile OutputProfile { get; set; } = OfficeWorkflowOutputProfile.Faithful;

    /// <summary>Optional password used for source PDFs.</summary>
    public string? PdfPassword { get; set; }

    /// <summary>Expansion and intake controls.</summary>
    public PdfAssemblyOptions Options { get; set; } = new();

    /// <summary>Shared input and final output limits.</summary>
    public OfficeWorkflowLimits Limits { get; set; } = new();
}

/// <summary>Terminal result of one heterogeneous PDF assembly.</summary>
public sealed class PdfAssemblyResult {
    internal PdfAssemblyResult(
        string requestId,
        OfficeWorkflowStatus status,
        OfficeWorkflowFailureKind failureKind,
        string? outputPath,
        int sourceCount,
        int pageCount,
        long inputBytes,
        long outputBytes,
        TimeSpan duration,
        string summary,
        IReadOnlyList<OfficeWorkflowDiagnostic> diagnostics) {
        RequestId = requestId;
        Status = status;
        FailureKind = failureKind;
        OutputPath = outputPath;
        SourceCount = sourceCount;
        PageCount = pageCount;
        InputBytes = inputBytes;
        OutputBytes = outputBytes;
        Duration = duration;
        Summary = summary;
        Diagnostics = diagnostics.ToArray();
    }

    /// <summary>Caller-provided request identifier.</summary>
    public string RequestId { get; }
    /// <summary>Terminal state.</summary>
    public OfficeWorkflowStatus Status { get; }
    /// <summary>Stable failure category, or <see cref="OfficeWorkflowFailureKind.None"/> when no failure occurred.</summary>
    public OfficeWorkflowFailureKind FailureKind { get; }
    /// <summary>Published PDF path.</summary>
    public string? OutputPath { get; }
    /// <summary>Normalized source count.</summary>
    public int SourceCount { get; }
    /// <summary>Output page count.</summary>
    public int PageCount { get; }
    /// <summary>Aggregate normalized source bytes.</summary>
    public long InputBytes { get; }
    /// <summary>Published PDF byte count.</summary>
    public long OutputBytes { get; }
    /// <summary>Total duration.</summary>
    public TimeSpan Duration { get; }
    /// <summary>User-facing outcome.</summary>
    public string Summary { get; }
    /// <summary>Structured diagnostics.</summary>
    public IReadOnlyList<OfficeWorkflowDiagnostic> Diagnostics { get; }
    /// <summary>Whether assembly completed successfully.</summary>
    public bool Succeeded => Status == OfficeWorkflowStatus.Completed;
}
