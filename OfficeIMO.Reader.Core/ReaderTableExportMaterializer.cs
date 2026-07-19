using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;

namespace OfficeIMO.Reader;

/// <summary>
/// Table export payload formats.
/// </summary>
public enum ReaderTableExportFormat {
    /// <summary>Comma-separated values.</summary>
    Csv,

    /// <summary>GitHub-style Markdown table.</summary>
    Markdown,

    /// <summary>JSON table sidecar.</summary>
    Json
}

/// <summary>
/// Options controlling how table export bundles are materialized.
/// </summary>
public sealed class ReaderTableExportMaterializationOptions {
    /// <summary>
    /// Creates the destination directory when it does not exist. Defaults to true.
    /// </summary>
    public bool CreateDirectory { get; set; } = true;

    /// <summary>
    /// Overwrites existing files with the same deterministic filename. Defaults to true.
    /// </summary>
    public bool Overwrite { get; set; } = true;

    /// <summary>
    /// Writes CSV payloads. Defaults to true.
    /// </summary>
    public bool IncludeCsv { get; set; } = true;

    /// <summary>
    /// Writes Markdown payloads. Defaults to true.
    /// </summary>
    public bool IncludeMarkdown { get; set; } = true;

    /// <summary>
    /// Writes JSON payloads. Defaults to true.
    /// </summary>
    public bool IncludeJson { get; set; } = true;

    /// <summary>
    /// Optional export predicate used to materialize only selected table bundles.
    /// </summary>
    public Func<ReaderTableExportBundle, bool>? Predicate { get; set; }
}

/// <summary>
/// Result for a single table sidecar materialization attempt.
/// </summary>
public sealed class ReaderTableMaterializedExport {
    /// <summary>
    /// Export bundle that owns the payload.
    /// </summary>
    public ReaderTableExportBundle Export { get; set; } = new ReaderTableExportBundle();

    /// <summary>
    /// Payload format.
    /// </summary>
    public ReaderTableExportFormat Format { get; set; }

    /// <summary>
    /// Deterministic filename used for this materialization attempt.
    /// </summary>
    public string? FileName { get; set; }

    /// <summary>
    /// Full output path when payloads are written to a directory.
    /// </summary>
    public string? Path { get; set; }

    /// <summary>
    /// True when the payload was written or streamed.
    /// </summary>
    public bool Written { get; set; }

    /// <summary>
    /// Explanation when the payload was skipped.
    /// </summary>
    public string? SkippedReason { get; set; }
}

/// <summary>
/// Helpers for writing or streaming table export sidecars.
/// </summary>
public static class ReaderTableExportMaterializer {
    private static readonly Encoding Utf8NoBom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

    /// <summary>
    /// Writes table export sidecars to <paramref name="directoryPath"/> using each bundle's deterministic filename stem.
    /// </summary>
    /// <param name="exports">Table export bundles.</param>
    /// <param name="directoryPath">Destination directory.</param>
    /// <param name="options">Materialization options.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IReadOnlyList<ReaderTableMaterializedExport> WriteTableExportsToDirectory(
        this IEnumerable<ReaderTableExportBundle> exports,
        string directoryPath,
        ReaderTableExportMaterializationOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (exports == null) throw new ArgumentNullException(nameof(exports));
        if (string.IsNullOrWhiteSpace(directoryPath)) throw new ArgumentException("Directory path cannot be empty.", nameof(directoryPath));

        ReaderTableExportMaterializationOptions effectiveOptions = options ?? new ReaderTableExportMaterializationOptions();
        if (effectiveOptions.CreateDirectory) {
            Directory.CreateDirectory(directoryPath);
        } else if (!Directory.Exists(directoryPath)) {
            throw new DirectoryNotFoundException("Directory '" + directoryPath + "' does not exist.");
        }

        var results = new List<ReaderTableMaterializedExport>();
        foreach (ReaderTableExportBundle export in SelectExports(exports, effectiveOptions)) {
            cancellationToken.ThrowIfCancellationRequested();
            foreach (ReaderTableExportFormat format in SelectFormats(effectiveOptions)) {
                cancellationToken.ThrowIfCancellationRequested();
                string fileName = ResolveFileName(export, format);
                string outputPath = System.IO.Path.Combine(directoryPath, fileName);
                string payload = GetPayload(export, format);

                if (!effectiveOptions.Overwrite && File.Exists(outputPath)) {
                    results.Add(Skipped(export, format, fileName, outputPath, "Destination file already exists."));
                    continue;
                }

                ReaderFileCommit.WriteAllBytes(outputPath, Utf8NoBom.GetBytes(payload));
                results.Add(new ReaderTableMaterializedExport {
                    Export = export,
                    Format = format,
                    FileName = fileName,
                    Path = outputPath,
                    Written = true
                });
            }
        }

        return results;
    }

    /// <summary>
    /// Streams table export sidecars to a caller-owned callback without writing files.
    /// </summary>
    /// <param name="exports">Table export bundles.</param>
    /// <param name="writeExport">Callback that receives each export, format, and readable payload stream.</param>
    /// <param name="options">Materialization options.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IReadOnlyList<ReaderTableMaterializedExport> StreamTableExports(
        this IEnumerable<ReaderTableExportBundle> exports,
        Action<ReaderTableExportBundle, ReaderTableExportFormat, Stream> writeExport,
        ReaderTableExportMaterializationOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (exports == null) throw new ArgumentNullException(nameof(exports));
        if (writeExport == null) throw new ArgumentNullException(nameof(writeExport));

        ReaderTableExportMaterializationOptions effectiveOptions = options ?? new ReaderTableExportMaterializationOptions();
        var results = new List<ReaderTableMaterializedExport>();
        foreach (ReaderTableExportBundle export in SelectExports(exports, effectiveOptions)) {
            cancellationToken.ThrowIfCancellationRequested();
            foreach (ReaderTableExportFormat format in SelectFormats(effectiveOptions)) {
                cancellationToken.ThrowIfCancellationRequested();
                string payload = GetPayload(export, format);
                byte[] bytes = Utf8NoBom.GetBytes(payload);
                using var stream = new MemoryStream(bytes, writable: false);
                writeExport(export, format, stream);
                results.Add(new ReaderTableMaterializedExport {
                    Export = export,
                    Format = format,
                    FileName = ResolveFileName(export, format),
                    Written = true
                });
            }
        }

        return results;
    }

    private static IEnumerable<ReaderTableExportBundle> SelectExports(IEnumerable<ReaderTableExportBundle> exports, ReaderTableExportMaterializationOptions options) {
        foreach (ReaderTableExportBundle? export in exports) {
            if (export == null) {
                continue;
            }

            if (options.Predicate == null || options.Predicate(export)) {
                yield return export;
            }
        }
    }

    private static IEnumerable<ReaderTableExportFormat> SelectFormats(ReaderTableExportMaterializationOptions options) {
        if (options.IncludeCsv) yield return ReaderTableExportFormat.Csv;
        if (options.IncludeMarkdown) yield return ReaderTableExportFormat.Markdown;
        if (options.IncludeJson) yield return ReaderTableExportFormat.Json;
    }

    private static string ResolveFileName(ReaderTableExportBundle export, ReaderTableExportFormat format) {
        string prefix = string.IsNullOrWhiteSpace(export.FileNamePrefix) ? export.Id : export.FileNamePrefix;
        if (string.IsNullOrWhiteSpace(prefix)) {
            prefix = "table-export";
        }

        string fileName = OfficeDocumentAssetNaming.BuildFileName(prefix, GetExtension(format));
        fileName = System.IO.Path.GetFileName(fileName);
        return string.IsNullOrWhiteSpace(fileName) ? OfficeDocumentAssetNaming.BuildFileName("table-export", GetExtension(format)) : fileName;
    }

    private static string GetPayload(ReaderTableExportBundle export, ReaderTableExportFormat format) {
        switch (format) {
            case ReaderTableExportFormat.Csv:
                return export.Csv ?? string.Empty;
            case ReaderTableExportFormat.Markdown:
                return export.Markdown ?? string.Empty;
            case ReaderTableExportFormat.Json:
                return export.Json ?? string.Empty;
            default:
                throw new ArgumentOutOfRangeException(nameof(format), format, "Unsupported table export format.");
        }
    }

    private static string GetExtension(ReaderTableExportFormat format) {
        switch (format) {
            case ReaderTableExportFormat.Csv:
                return ".csv";
            case ReaderTableExportFormat.Markdown:
                return ".md";
            case ReaderTableExportFormat.Json:
                return ".json";
            default:
                throw new ArgumentOutOfRangeException(nameof(format), format, "Unsupported table export format.");
        }
    }

    private static ReaderTableMaterializedExport Skipped(ReaderTableExportBundle export, ReaderTableExportFormat format, string fileName, string? path, string reason) {
        return new ReaderTableMaterializedExport {
            Export = export,
            Format = format,
            FileName = fileName,
            Path = path,
            Written = false,
            SkippedReason = reason
        };
    }
}
