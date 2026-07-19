using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;

namespace OfficeIMO.Reader;

/// <summary>
/// Visual export payload formats.
/// </summary>
public enum ReaderVisualExportFormat {
    /// <summary>Raw source visual payload.</summary>
    Payload,

    /// <summary>JSON visual sidecar.</summary>
    Json
}

/// <summary>
/// Options controlling how visual export bundles are materialized.
/// </summary>
public sealed class ReaderVisualExportMaterializationOptions {
    /// <summary>
    /// Creates the destination directory when it does not exist. Defaults to true.
    /// </summary>
    public bool CreateDirectory { get; set; } = true;

    /// <summary>
    /// Overwrites existing files with the same deterministic filename. Defaults to true.
    /// </summary>
    public bool Overwrite { get; set; } = true;

    /// <summary>
    /// Writes raw source visual payloads. Defaults to true.
    /// </summary>
    public bool IncludePayload { get; set; } = true;

    /// <summary>
    /// Writes JSON visual sidecars. Defaults to true.
    /// </summary>
    public bool IncludeJson { get; set; } = true;

    /// <summary>
    /// Optional export predicate used to materialize only selected visual bundles.
    /// </summary>
    public Func<ReaderVisualExportBundle, bool>? Predicate { get; set; }
}

/// <summary>
/// Result for a single visual sidecar materialization attempt.
/// </summary>
public sealed class ReaderVisualMaterializedExport {
    /// <summary>
    /// Export bundle that owns the payload.
    /// </summary>
    public ReaderVisualExportBundle Export { get; set; } = new ReaderVisualExportBundle();

    /// <summary>
    /// Payload format.
    /// </summary>
    public ReaderVisualExportFormat Format { get; set; }

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
/// Helpers for writing or streaming visual export sidecars.
/// </summary>
public static class ReaderVisualExportMaterializer {
    private static readonly Encoding Utf8NoBom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

    /// <summary>
    /// Writes visual export sidecars to <paramref name="directoryPath"/> using each bundle's deterministic filename stem.
    /// </summary>
    /// <param name="exports">Visual export bundles.</param>
    /// <param name="directoryPath">Destination directory.</param>
    /// <param name="options">Materialization options.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IReadOnlyList<ReaderVisualMaterializedExport> WriteVisualExportsToDirectory(
        this IEnumerable<ReaderVisualExportBundle> exports,
        string directoryPath,
        ReaderVisualExportMaterializationOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (exports == null) throw new ArgumentNullException(nameof(exports));
        if (string.IsNullOrWhiteSpace(directoryPath)) throw new ArgumentException("Directory path cannot be empty.", nameof(directoryPath));

        ReaderVisualExportMaterializationOptions effectiveOptions = options ?? new ReaderVisualExportMaterializationOptions();
        if (effectiveOptions.CreateDirectory) {
            Directory.CreateDirectory(directoryPath);
        } else if (!Directory.Exists(directoryPath)) {
            throw new DirectoryNotFoundException("Directory '" + directoryPath + "' does not exist.");
        }

        var results = new List<ReaderVisualMaterializedExport>();
        foreach (ReaderVisualExportBundle export in SelectExports(exports, effectiveOptions)) {
            cancellationToken.ThrowIfCancellationRequested();
            foreach (ReaderVisualExportFormat format in SelectFormats(effectiveOptions)) {
                cancellationToken.ThrowIfCancellationRequested();
                string fileName = ResolveFileName(export, format);
                string outputPath = System.IO.Path.Combine(directoryPath, fileName);
                string payload = GetPayload(export, format);

                if (!effectiveOptions.Overwrite && File.Exists(outputPath)) {
                    results.Add(Skipped(export, format, fileName, outputPath, "Destination file already exists."));
                    continue;
                }

                ReaderFileCommit.WriteAllBytes(outputPath, Utf8NoBom.GetBytes(payload));
                results.Add(new ReaderVisualMaterializedExport {
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
    /// Streams visual export sidecars to a caller-owned callback without writing files.
    /// </summary>
    /// <param name="exports">Visual export bundles.</param>
    /// <param name="writeExport">Callback that receives each export, format, and readable payload stream.</param>
    /// <param name="options">Materialization options.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IReadOnlyList<ReaderVisualMaterializedExport> StreamVisualExports(
        this IEnumerable<ReaderVisualExportBundle> exports,
        Action<ReaderVisualExportBundle, ReaderVisualExportFormat, Stream> writeExport,
        ReaderVisualExportMaterializationOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (exports == null) throw new ArgumentNullException(nameof(exports));
        if (writeExport == null) throw new ArgumentNullException(nameof(writeExport));

        ReaderVisualExportMaterializationOptions effectiveOptions = options ?? new ReaderVisualExportMaterializationOptions();
        var results = new List<ReaderVisualMaterializedExport>();
        foreach (ReaderVisualExportBundle export in SelectExports(exports, effectiveOptions)) {
            cancellationToken.ThrowIfCancellationRequested();
            foreach (ReaderVisualExportFormat format in SelectFormats(effectiveOptions)) {
                cancellationToken.ThrowIfCancellationRequested();
                string payload = GetPayload(export, format);
                byte[] bytes = Utf8NoBom.GetBytes(payload);
                using var stream = new MemoryStream(bytes, writable: false);
                writeExport(export, format, stream);
                results.Add(new ReaderVisualMaterializedExport {
                    Export = export,
                    Format = format,
                    FileName = ResolveFileName(export, format),
                    Written = true
                });
            }
        }

        return results;
    }

    private static IEnumerable<ReaderVisualExportBundle> SelectExports(IEnumerable<ReaderVisualExportBundle> exports, ReaderVisualExportMaterializationOptions options) {
        foreach (ReaderVisualExportBundle? export in exports) {
            if (export == null) {
                continue;
            }

            if (options.Predicate == null || options.Predicate(export)) {
                yield return export;
            }
        }
    }

    private static IEnumerable<ReaderVisualExportFormat> SelectFormats(ReaderVisualExportMaterializationOptions options) {
        if (options.IncludePayload) yield return ReaderVisualExportFormat.Payload;
        if (options.IncludeJson) yield return ReaderVisualExportFormat.Json;
    }

    private static string ResolveFileName(ReaderVisualExportBundle export, ReaderVisualExportFormat format) {
        string prefix = string.IsNullOrWhiteSpace(export.FileNamePrefix) ? export.Id : export.FileNamePrefix;
        if (string.IsNullOrWhiteSpace(prefix)) {
            prefix = "visual-export";
        }

        string extension = GetExtension(export, format);
        if (format == ReaderVisualExportFormat.Json &&
            string.Equals(NormalizeExtension(export.PayloadExtension), ".json", StringComparison.OrdinalIgnoreCase)) {
            prefix += ".metadata";
        }

        string fileName = OfficeDocumentAssetNaming.BuildFileName(prefix, extension);
        fileName = System.IO.Path.GetFileName(fileName);
        return string.IsNullOrWhiteSpace(fileName) ? OfficeDocumentAssetNaming.BuildFileName("visual-export", extension) : fileName;
    }

    private static string GetPayload(ReaderVisualExportBundle export, ReaderVisualExportFormat format) {
        switch (format) {
            case ReaderVisualExportFormat.Payload:
                return export.Payload ?? string.Empty;
            case ReaderVisualExportFormat.Json:
                return export.Json ?? string.Empty;
            default:
                throw new ArgumentOutOfRangeException(nameof(format), format, "Unsupported visual export format.");
        }
    }

    private static string GetExtension(ReaderVisualExportBundle export, ReaderVisualExportFormat format) {
        switch (format) {
            case ReaderVisualExportFormat.Payload:
                return string.IsNullOrWhiteSpace(export.PayloadExtension) ? ".txt" : export.PayloadExtension;
            case ReaderVisualExportFormat.Json:
                return ".json";
            default:
                throw new ArgumentOutOfRangeException(nameof(format), format, "Unsupported visual export format.");
        }
    }

    private static string? NormalizeExtension(string? extension) {
        if (string.IsNullOrWhiteSpace(extension)) {
            return null;
        }

        string trimmed = extension!.Trim();
        if (trimmed.Length == 0) {
            return null;
        }

        return trimmed.StartsWith(".", StringComparison.Ordinal) ? trimmed : "." + trimmed;
    }

    private static ReaderVisualMaterializedExport Skipped(ReaderVisualExportBundle export, ReaderVisualExportFormat format, string fileName, string? path, string reason) {
        return new ReaderVisualMaterializedExport {
            Export = export,
            Format = format,
            FileName = fileName,
            Path = path,
            Written = false,
            SkippedReason = reason
        };
    }
}
