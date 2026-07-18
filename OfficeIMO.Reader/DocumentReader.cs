using OfficeIMO.Excel;
using OfficeIMO.Markdown;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading;

namespace OfficeIMO.Reader;

/// <summary>
/// Unified, read-only document extraction facade intended for AI ingestion.
/// </summary>
/// <remarks>
/// This facade is intentionally dependency-free and deterministic.
/// It normalizes extraction into <see cref="ReaderChunk"/> instances with stable IDs and location metadata.
/// Read operations are thread-safe. Use <see cref="OfficeDocumentReaderBuilder"/> when modular handlers
/// or a processing pipeline are required.
/// </remarks>
internal static partial class DocumentReaderEngine {
    private static readonly string[] DefaultFolderExtensions = {
        ".docx", ".docm", ".doc",
        ".xlsx", ".xlsm", ".xls",
        ".pptx", ".pptm", ".ppt", ".pot", ".pps",
        ".md", ".markdown",
        ".pdf",
        ".eml", ".msg", ".oft", ".mbox", ".mbx", ".tnef",
        ".txt", ".log", ".csv", ".tsv", ".json", ".xml", ".yml", ".yaml"
    };

    private static readonly ReaderHandlerCapability[] BuiltInCapabilities = {
        new ReaderHandlerCapability {
            Id = "officeimo.reader.word",
            DisplayName = "Word Reader",
            Description = "Built-in Word (.docx/.docm/.doc) chunk extractor.",
            Kind = ReaderInputKind.Word,
            Extensions = new[] { ".docx", ".docm", ".doc" },
            IsBuiltIn = true,
            SupportsPath = true,
            SupportsStream = true
        },
        new ReaderHandlerCapability {
            Id = "officeimo.reader.excel",
            DisplayName = "Excel Reader",
            Description = "Built-in Excel (.xlsx/.xlsm/.xls) table and markdown extractor.",
            Kind = ReaderInputKind.Excel,
            Extensions = new[] { ".xlsx", ".xlsm", ".xls" },
            IsBuiltIn = true,
            SupportsPath = true,
            SupportsStream = true
        },
        new ReaderHandlerCapability {
            Id = "officeimo.reader.powerpoint",
            DisplayName = "PowerPoint Reader",
            Description = "Built-in Open XML PowerPoint (.pptx/.pptm) slide extractor.",
            Kind = ReaderInputKind.PowerPoint,
            Extensions = new[] { ".pptx", ".pptm" },
            IsBuiltIn = true,
            SupportsPath = true,
            SupportsStream = true
        },
        new ReaderHandlerCapability {
            Id = "officeimo.reader.powerpoint.binary",
            DisplayName = "Binary PowerPoint Reader",
            Description = "Built-in PowerPoint 97-2003 (.ppt/.pot/.pps) slide extractor.",
            Kind = ReaderInputKind.PowerPoint,
            Extensions = new[] { ".ppt", ".pot", ".pps" },
            IsBuiltIn = true,
            SupportsPath = true,
            SupportsStream = true,
            DefaultMaxInputBytes =
                LegacyPptImportOptions.DefaultMaxInputBytes
        },
        new ReaderHandlerCapability {
            Id = "officeimo.reader.markdown",
            DisplayName = "Markdown Reader",
            Description = "Built-in Markdown chunk extractor.",
            Kind = ReaderInputKind.Markdown,
            Extensions = new[] { ".md", ".markdown" },
            IsBuiltIn = true,
            SupportsPath = true,
            SupportsStream = true
        },
        new ReaderHandlerCapability {
            Id = "officeimo.reader.pdf",
            DisplayName = "PDF Reader",
            Description = "Built-in PDF logical page and markdown extractor.",
            Kind = ReaderInputKind.Pdf,
            Extensions = new[] { ".pdf" },
            IsBuiltIn = true,
            SupportsPath = true,
            SupportsStream = true
        },
        new ReaderHandlerCapability {
            Id = "officeimo.reader.email",
            DisplayName = "Email Reader",
            Description = "Built-in EML/MIME, Outlook MSG/OFT/MAPI, TNEF, and mbox extractor.",
            Kind = ReaderInputKind.Email,
            Extensions = new[] { ".eml", ".msg", ".oft", ".mbox", ".mbx", ".tnef" },
            IsBuiltIn = true,
            SupportsPath = true,
            SupportsStream = true,
            DefaultMaxInputBytes = OfficeIMO.Email.EmailReaderOptions.Default.MaxInputBytes
        },
        new ReaderHandlerCapability {
            Id = "officeimo.reader.text",
            DisplayName = "Text Reader",
            Description = "Built-in plain text reader for text-like formats.",
            Kind = ReaderInputKind.Text,
            Extensions = new[] { ".txt", ".log", ".csv", ".tsv", ".json", ".xml", ".yml", ".yaml" },
            IsBuiltIn = true,
            SupportsPath = true,
            SupportsStream = true
        }
    };

    internal static readonly HashSet<string> BuiltInExtensions = BuildBuiltInExtensionSet();
    private static readonly ReaderHandlerRegistrySnapshot BuiltInHandlerRegistry =
        new ReaderHandlerRegistry(BuiltInExtensions).CaptureSnapshot();

    private static string? TryGetExtension(string path) {
        if (path == null) return null;
        try {
            return Path.GetExtension(path);
        } catch (ArgumentException) {
            return null;
        } catch (NotSupportedException) {
            return null;
        }
    }

    private static bool IsLegacyWordExtension(string? path) {
        return string.Equals(NormalizeExtension(TryGetExtension(path ?? string.Empty)), ".doc", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsLegacyExcelExtension(string? path) {
        return string.Equals(NormalizeExtension(TryGetExtension(path ?? string.Empty)), ".xls", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsLegacyPowerPointExtension(string? path) {
        string extension = NormalizeExtension(TryGetExtension(path ?? string.Empty));
        return extension is ".ppt" or ".pot" or ".pps";
    }

    private static bool IsLegacyBinaryOfficeExtension(string? path) {
        return IsLegacyWordExtension(path) || IsLegacyExcelExtension(path)
            || IsLegacyPowerPointExtension(path);
    }

    /// <summary>
    /// Lists the capabilities supported by the built-in static reader.
    /// </summary>
    public static IReadOnlyList<ReaderHandlerCapability> GetCapabilities() {
        return GetCapabilities(BuiltInHandlerRegistry);
    }

    internal static IReadOnlyList<ReaderHandlerCapability> GetCapabilities(ReaderHandlerRegistrySnapshot snapshot) {
        var list = new List<ReaderHandlerCapability>();
        var customCapabilities = new List<ReaderHandlerCapability>();
        Dictionary<string, ExtensionOverrideCoverage>? overriddenBuiltInExtensions = null;

        foreach (ReaderHandlerDescriptor custom in snapshot.Handlers) {
            ReaderHandlerCapability capability = custom.ToCapability();
            customCapabilities.Add(capability);
            for (int extensionIndex = 0; extensionIndex < capability.Extensions.Count; extensionIndex++) {
                string extension = capability.Extensions[extensionIndex];
                if (BuiltInExtensions.Contains(extension)) {
                    overriddenBuiltInExtensions ??= new Dictionary<string, ExtensionOverrideCoverage>(StringComparer.OrdinalIgnoreCase);
                    overriddenBuiltInExtensions.TryGetValue(extension, out ExtensionOverrideCoverage coverage);
                    overriddenBuiltInExtensions[extension] = coverage.Add(capability.SupportsPath, capability.SupportsStream);
                }
            }
        }

        for (int i = 0; i < BuiltInCapabilities.Length; i++) {
            ReaderHandlerCapability? capability = CloneCapabilityWithRemainingSupport(BuiltInCapabilities[i], overriddenBuiltInExtensions);
            if (capability is not null) {
                list.Add(capability);
            }
        }

        list.AddRange(customCapabilities);

        return list
            .OrderBy(static c => c.IsBuiltIn ? 0 : 1)
            .ThenBy(static c => c.Id, StringComparer.Ordinal)
            .ToArray();
    }

    /// <summary>
    /// Builds a machine-readable capability manifest for the built-in static reader.
    /// </summary>
    public static ReaderCapabilityManifest GetCapabilityManifest() {
        return GetCapabilityManifest(BuiltInHandlerRegistry);
    }

    internal static ReaderCapabilityManifest GetCapabilityManifest(ReaderHandlerRegistrySnapshot snapshot) {
        var handlers = GetCapabilities(snapshot)
            .Select(CloneCapability)
            .ToArray();

        return new ReaderCapabilityManifest {
            SchemaId = ReaderCapabilitySchema.Id,
            SchemaVersion = ReaderCapabilitySchema.Version,
            Handlers = handlers
        };
    }

    /// <summary>
    /// Builds a JSON capability manifest payload for the built-in static reader.
    /// </summary>
    public static string GetCapabilityManifestJson(bool indented = false) {
        return ReaderCapabilityManifestJson.Serialize(GetCapabilityManifest(), indented);
    }

    /// <summary>
    /// Extracts structured tables from Markdown text using the same parser path as Markdown reader chunks.
    /// </summary>
    /// <param name="markdown">Markdown text to inspect.</param>
    /// <param name="options">Reader options. <see cref="ReaderOptions.MaxTableRows"/> is honored.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <returns>Tables discovered in source order.</returns>
    public static IReadOnlyList<ReaderTable> ExtractMarkdownTables(string markdown, ReaderOptions? options = null, CancellationToken cancellationToken = default) {
        var opt = NormalizeOptions(options);
        var tables = new List<ReaderTable>();

        foreach (var block in ParseMarkdownBlocksForChunking(markdown ?? string.Empty, opt, cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (block.Tables.Count == 0) {
                continue;
            }

            tables.AddRange(block.Tables);
        }

        return tables.Count == 0 ? Array.Empty<ReaderTable>() : tables.ToArray();
    }

    /// <summary>
    /// Detects the input kind based on file extension.
    /// </summary>
    /// <param name="path">Source file path.</param>
    public static ReaderInputKind DetectKind(string path) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (path.Length == 0) throw new ArgumentException("Path cannot be empty.", nameof(path));

        var extLower = NormalizeExtension(GetLogicalExtension(path));
        if (extLower.Length > 0 && TryResolveCustomHandlerByExtension(extLower, out var custom)) {
            return custom.Kind;
        }
        if (string.Equals(GetLogicalFileName(path), "winmail.dat", StringComparison.OrdinalIgnoreCase)) {
            return ReaderInputKind.Email;
        }
        return DetectBuiltInKind(path);
    }

    private static string GetLogicalFileName(string path) {
        int separator = Math.Max(path.LastIndexOf('/'), path.LastIndexOf('\\'));
        return separator < 0 ? path : path.Substring(separator + 1);
    }

    private static string GetLogicalExtension(string path) {
        string fileName = GetLogicalFileName(path);
        int separator = fileName.LastIndexOf('.');
        return separator <= 0 ? string.Empty : fileName.Substring(separator);
    }

    private static ReaderInputKind DetectBuiltInKind(string path) {
        var extLower = NormalizeExtension(GetLogicalExtension(path));
        if (extLower.Length == 0) return ReaderInputKind.Unknown;
        return extLower switch {
            ".docx" or ".docm" or ".doc" => ReaderInputKind.Word,
            ".xlsx" or ".xlsm" or ".xls" => ReaderInputKind.Excel,
            ".pptx" or ".pptm" or ".ppt" or ".pot" or ".pps" => ReaderInputKind.PowerPoint,
            ".md" or ".markdown" => ReaderInputKind.Markdown,
            ".pdf" => ReaderInputKind.Pdf,
            ".eml" or ".msg" or ".oft" or ".mbox" or ".mbx" or ".tnef" => ReaderInputKind.Email,
            ".txt" or ".log" or ".csv" or ".tsv" or ".json" or ".xml" or ".yml" or ".yaml" => ReaderInputKind.Text,
            _ => ReaderInputKind.Unknown
        };
    }

}
