using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Provenance;

/// <summary>Inspects standards-defined provenance carriers without performing cryptographic validation.</summary>
public static class OfficeProvenanceInspector {
    /// <summary>Inspects an asset file using bounded structural parsers.</summary>
    public static OfficeProvenanceReport InspectFile(string filePath, OfficeProvenanceOptions? options = null) {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("A file path is required.", nameof(filePath));
        string fullPath = Path.GetFullPath(filePath);
        using var stream = File.OpenRead(fullPath);
        return Inspect(stream, fullPath, options);
    }

    /// <summary>Inspects a bounded asset stream and restores its original position when seekable.</summary>
    public static OfficeProvenanceReport Inspect(Stream stream, string? fileName = null, OfficeProvenanceOptions? options = null) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        options ??= new OfficeProvenanceOptions();
        OfficeProvenanceBinary.ValidateLimits(options);
        long originalPosition = stream.CanSeek ? stream.Position : 0L;
        try {
            byte[] data = OfficeProvenanceBinary.ReadBounded(stream, options.MaxAssetBytes);
            return Inspect(data, fileName, options);
        } finally {
            if (stream.CanSeek) stream.Position = originalPosition;
        }
    }

    /// <summary>Inspects bounded encoded asset bytes.</summary>
    public static OfficeProvenanceReport Inspect(byte[] data, string? fileName = null, OfficeProvenanceOptions? options = null) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        options ??= new OfficeProvenanceOptions();
        OfficeProvenanceBinary.ValidateLimits(options);
        if (data.LongLength > options.MaxAssetBytes) {
            throw OfficeProvenanceLimitException.Create($"The asset exceeds the configured limit of {options.MaxAssetBytes} bytes.");
        }
        return InspectCore(data, fileName, options);
    }

    internal static OfficeProvenanceReport InspectCore(byte[] data, string? fileName, OfficeProvenanceOptions options) {
        OfficeProvenanceAssetFormat format = DetectFormat(data, fileName, options);
        return InspectCore(data, options, format);
    }

    internal static OfficeProvenanceReport InspectStructuredText(byte[] data, OfficeProvenanceOptions options) =>
        InspectCore(data, options, OfficeProvenanceAssetFormat.StructuredText);

    private static OfficeProvenanceReport InspectCore(
        byte[] data,
        OfficeProvenanceOptions options,
        OfficeProvenanceAssetFormat format) {
        var context = new OfficeProvenanceContext(format, options);
        switch (format) {
            case OfficeProvenanceAssetFormat.Jpeg:
                OfficeProvenanceJpeg.Inspect(data, options, context);
                break;
            case OfficeProvenanceAssetFormat.Png:
                OfficeProvenancePng.Inspect(data, options, context);
                break;
            case OfficeProvenanceAssetFormat.Webp:
                OfficeProvenanceRiff.Inspect(data, options, context);
                break;
            case OfficeProvenanceAssetFormat.Gif:
                OfficeProvenanceGif.Inspect(data, options, context);
                break;
            case OfficeProvenanceAssetFormat.Tiff:
                OfficeProvenanceTiff.Inspect(data, options, context);
                break;
            case OfficeProvenanceAssetFormat.Svg:
                OfficeProvenanceSvg.Inspect(data, options, context);
                break;
            case OfficeProvenanceAssetFormat.ZipPackage:
                OfficeProvenanceZip.Inspect(data, options, context);
                break;
            case OfficeProvenanceAssetFormat.StructuredText:
            case OfficeProvenanceAssetFormat.UnstructuredText:
                OfficeProvenanceText.Inspect(data, options, context);
                break;
        }
        return context.ToReport();
    }

    internal static OfficeProvenanceAssetFormat DetectFormat(
        byte[] data,
        string? fileName,
        OfficeProvenanceOptions options) {
        if (OfficeProvenanceBinary.HasPrefix(data, 0xFF, 0xD8)) return OfficeProvenanceAssetFormat.Jpeg;
        if (OfficeProvenanceBinary.HasPrefix(data, 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A)) return OfficeProvenanceAssetFormat.Png;
        if (data.Length >= 12 && OfficeProvenanceBinary.MatchesAscii(data, 0, "RIFF") && OfficeProvenanceBinary.MatchesAscii(data, 8, "WEBP")) {
            return OfficeProvenanceAssetFormat.Webp;
        }
        if (data.Length >= 6 && (OfficeProvenanceBinary.MatchesAscii(data, 0, "GIF87a") || OfficeProvenanceBinary.MatchesAscii(data, 0, "GIF89a"))) {
            return OfficeProvenanceAssetFormat.Gif;
        }
        if (data.Length >= 8 && ((data[0] == (byte)'I' && data[1] == (byte)'I') || (data[0] == (byte)'M' && data[1] == (byte)'M'))) {
            ushort version = OfficeProvenanceBinary.ReadUInt16(data, 2, data[0] == (byte)'I');
            if (version == 42 || version == 43) return OfficeProvenanceAssetFormat.Tiff;
        }
        if (OfficeProvenanceZip.HasSignature(data)) return OfficeProvenanceAssetFormat.ZipPackage;
        if (OfficeProvenanceBinary.MatchesAscii(data, 0, "%PDF-")) return OfficeProvenanceAssetFormat.Pdf;
        if (LooksLikeSvg(data, null, options.MaxContainerEntries)) return OfficeProvenanceAssetFormat.Svg;
        if (LooksLikeHtml(data, null, options.MaxContainerEntries)) return OfficeProvenanceAssetFormat.Html;
        if (OfficeProvenanceText.HasUnstructuredWrapperPrefix(data, options.MaxContainerEntries)) return OfficeProvenanceAssetFormat.UnstructuredText;
        if (OfficeProvenanceText.HasStructuredDelimiter(data)) return OfficeProvenanceAssetFormat.StructuredText;
        string extension = Path.GetExtension(fileName ?? string.Empty);
        if (extension.Equals(".svg", StringComparison.OrdinalIgnoreCase)) return OfficeProvenanceAssetFormat.Svg;
        if (extension.Equals(".html", StringComparison.OrdinalIgnoreCase) ||
            extension.Equals(".htm", StringComparison.OrdinalIgnoreCase)) return OfficeProvenanceAssetFormat.Html;
        if (HasStructuredTextExtension(fileName)) return OfficeProvenanceAssetFormat.StructuredText;
        if (extension.Equals(".pdf", StringComparison.OrdinalIgnoreCase)) return OfficeProvenanceAssetFormat.Pdf;
        return OfficeProvenanceAssetFormat.Unknown;
    }

    private static bool LooksLikeHtml(byte[] data, string? fileName, int maximumContainerEntries) {
        string extension = Path.GetExtension(fileName ?? string.Empty);
        if (extension.Equals(".html", StringComparison.OrdinalIgnoreCase) ||
            extension.Equals(".htm", StringComparison.OrdinalIgnoreCase)) return true;
        int offset = 0;
        if (data.Length >= 3 && data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF) offset = 3;
        int commentCount = 0;
        while (true) {
            while (offset < data.Length && data[offset] is 0x09 or 0x0A or 0x0C or 0x0D or 0x20) offset++;
            if (!MatchesAsciiIgnoreCase(data, offset, "<!--")) break;
            if (++commentCount > maximumContainerEntries) {
                throw new InvalidDataException("HTML format detection exceeds the configured container-entry limit.");
            }
            int commentEnd = FindAscii(data, offset + 4, "-->");
            if (commentEnd < 0) return false;
            offset = commentEnd + 3;
        }
        return MatchesHtmlTokenWithBoundary(data, offset, "<!doctype html", allowSelfClosing: false) ||
            MatchesHtmlTokenWithBoundary(data, offset, "<html", allowSelfClosing: true);
    }

    private static bool MatchesHtmlTokenWithBoundary(byte[] data, int offset, string token, bool allowSelfClosing) {
        if (!MatchesAsciiIgnoreCase(data, offset, token)) return false;
        int boundary = offset + token.Length;
        if (boundary >= data.Length) return false;
        byte value = data[boundary];
        return value is 0x09 or 0x0A or 0x0C or 0x0D or 0x20 or (byte)'>' ||
            allowSelfClosing && value == (byte)'/';
    }

    private static int FindAscii(byte[] data, int offset, string expected) {
        for (int index = Math.Max(0, offset); index <= data.Length - expected.Length; index++) {
            if (MatchesAsciiIgnoreCase(data, index, expected)) return index;
        }
        return -1;
    }

    private static bool MatchesAsciiIgnoreCase(byte[] data, int offset, string expected) {
        if (offset < 0 || expected.Length > data.Length - offset) return false;
        for (int index = 0; index < expected.Length; index++) {
            byte actual = data[offset + index];
            byte wanted = (byte)expected[index];
            if (actual >= (byte)'A' && actual <= (byte)'Z') actual = (byte)(actual + 32);
            if (wanted >= (byte)'A' && wanted <= (byte)'Z') wanted = (byte)(wanted + 32);
            if (actual != wanted) return false;
        }
        return true;
    }

    private static bool LooksLikeSvg(byte[] data, string? fileName, int maximumContainerEntries) {
        string extension = Path.GetExtension(fileName ?? string.Empty);
        if (extension.Equals(".svg", StringComparison.OrdinalIgnoreCase)) return true;
        if (!CouldStartXml(data)) return false;
        try {
            var settings = new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = data.LongLength,
                MaxCharactersFromEntities = 0
            };
            using var stream = new MemoryStream(data, writable: false);
            using XmlReader reader = XmlReader.Create(stream, settings);
            int nodeCount = 0;
            while (reader.Read()) {
                int currentCount = reader.NodeType switch {
                    XmlNodeType.Element => 1 + reader.AttributeCount,
                    XmlNodeType.Text or XmlNodeType.CDATA or XmlNodeType.ProcessingInstruction or
                        XmlNodeType.Comment or XmlNodeType.Whitespace or XmlNodeType.SignificantWhitespace => 1,
                    _ => 0
                };
                if (currentCount > 0 && nodeCount > maximumContainerEntries - currentCount) {
                    throw new InvalidDataException("SVG format detection exceeds the configured XML node limit.");
                }
                nodeCount += currentCount;
                if (reader.Depth > 256) throw new InvalidDataException("SVG format detection exceeds the supported XML depth limit.");
                if (reader.NodeType != XmlNodeType.Element) continue;
                return reader.LocalName.Equals("svg", StringComparison.OrdinalIgnoreCase) &&
                    reader.NamespaceURI.Equals("http://www.w3.org/2000/svg", StringComparison.Ordinal);
            }
            return false;
        } catch (XmlException) {
            return false;
        }
    }

    private static bool CouldStartXml(byte[] data) {
        if (data.Length == 0) return false;
        if (data.Length >= 4 &&
            ((data[0] == 0x00 && data[1] == 0x00 && data[2] == 0xFE && data[3] == 0xFF) ||
             (data[0] == 0x00 && data[1] == 0x00 && data[2] == 0x00 && data[3] == (byte)'<'))) {
            return true;
        }
        if (data.Length >= 2 &&
            ((data[0] == 0xFF && data[1] == 0xFE) ||
             (data[0] == 0xFE && data[1] == 0xFF) ||
             (data[0] == 0x00 && data[1] == (byte)'<') ||
             (data[0] == (byte)'<' && data[1] == 0x00))) {
            return true;
        }

        int offset = data.Length >= 3 && data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF ? 3 : 0;
        while (offset < data.Length && data[offset] is 0x09 or 0x0A or 0x0C or 0x0D or 0x20) offset++;
        return offset < data.Length && data[offset] == (byte)'<';
    }

    private static bool HasStructuredTextExtension(string? fileName) {
        string extension = Path.GetExtension(fileName ?? string.Empty).ToLowerInvariant();
        return extension is ".md" or ".markdown" or ".txt" or ".yaml" or ".yml" or ".toml" or ".ini" or ".json" or ".xml" or
            ".adoc" or ".asciidoc" or ".tex" or ".py" or ".rb" or ".sh" or ".ps1" or ".js" or ".mjs" or ".cjs" or ".ts" or
            ".css" or ".sql" or ".cs" or ".vb" or ".java" or ".c" or ".h" or ".cpp" or ".hpp" or ".go" or ".rs" or ".lua" or
            ".bat" or ".cmd";
    }
}

/// <summary>Removes selected standards-defined provenance carriers while preserving unrelated data.</summary>
public static class OfficeProvenanceRemover {
    /// <summary>Removes selected provenance from encoded asset bytes.</summary>
    public static OfficeProvenanceRemovalResult Remove(
        byte[] data,
        string? fileName = null,
        OfficeProvenanceRemovalOptions? options = null) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        options ??= new OfficeProvenanceRemovalOptions();
        OfficeProvenanceBinary.ValidateRemovalOptions(options);
        if (data.LongLength > options.Limits.MaxAssetBytes) {
            throw new InvalidDataException($"The asset exceeds the configured limit of {options.Limits.MaxAssetBytes} bytes.");
        }

        OfficeProvenanceOptions inspectionOptions = CreateInspectionOptions(options);
        return RemoveCore(data, fileName, options, inspectionOptions, forcedFormat: null);
    }

    internal static OfficeProvenanceRemovalResult RemoveStructuredText(
        byte[] data,
        string? fileName,
        OfficeProvenanceRemovalOptions options) {
        OfficeProvenanceOptions inspectionOptions = CreateInspectionOptions(options);
        return RemoveCore(data, fileName, options, inspectionOptions, OfficeProvenanceAssetFormat.StructuredText);
    }

    internal static OfficeProvenanceRemovalResult RemoveZipPackage(
        byte[] data,
        string? fileName,
        OfficeProvenanceRemovalOptions options,
        bool removeOpcManifestReferences,
        Func<string, bool>? shouldReplacePackageMetadata = null,
        Func<string, byte[], bool, byte[]>? replacePackageMetadata = null) {
        OfficeProvenanceOptions inspectionOptions = CreateInspectionOptions(options);
        return RemoveCore(
            data,
            fileName,
            options,
            inspectionOptions,
            forcedFormat: null,
            removeOpcManifestReferences: removeOpcManifestReferences,
            shouldReplacePackageMetadata: shouldReplacePackageMetadata,
            replacePackageMetadata: replacePackageMetadata);
    }

    private static OfficeProvenanceRemovalResult RemoveCore(
        byte[] data,
        string? fileName,
        OfficeProvenanceRemovalOptions options,
        OfficeProvenanceOptions inspectionOptions,
        OfficeProvenanceAssetFormat? forcedFormat,
        bool removeOpcManifestReferences = true,
        Func<string, bool>? shouldReplacePackageMetadata = null,
        Func<string, byte[], bool, byte[]>? replacePackageMetadata = null) {
        OfficeProvenanceReport before = forcedFormat.HasValue
            ? OfficeProvenanceInspector.InspectStructuredText(data, inspectionOptions)
            : OfficeProvenanceInspector.InspectCore(data, fileName, inspectionOptions);
        var changes = new List<OfficeProvenanceChange>();
        byte[] output;
        bool reserialized = false;
        switch (before.Format) {
            case OfficeProvenanceAssetFormat.Jpeg:
                output = OfficeProvenanceJpeg.Remove(data, options, changes, out reserialized);
                break;
            case OfficeProvenanceAssetFormat.Png:
                output = OfficeProvenancePng.Remove(data, options, changes, out reserialized);
                break;
            case OfficeProvenanceAssetFormat.Webp:
                output = OfficeProvenanceRiff.Remove(data, options, changes, out reserialized);
                break;
            case OfficeProvenanceAssetFormat.Gif:
                output = OfficeProvenanceGif.Remove(data, options, changes, out reserialized);
                break;
            case OfficeProvenanceAssetFormat.Tiff:
                output = OfficeProvenanceTiff.Remove(data, options, changes, out reserialized);
                break;
            case OfficeProvenanceAssetFormat.Svg:
                output = OfficeProvenanceSvg.Remove(data, options, changes, out reserialized);
                break;
            case OfficeProvenanceAssetFormat.ZipPackage:
                output = OfficeProvenanceZip.Remove(
                    data,
                    options,
                    changes,
                    out reserialized,
                    removeOpcManifestReferences,
                    shouldReplacePackageMetadata,
                    replacePackageMetadata);
                break;
            case OfficeProvenanceAssetFormat.StructuredText:
            case OfficeProvenanceAssetFormat.UnstructuredText:
                output = OfficeProvenanceText.Remove(data, options, changes);
                break;
            default:
                output = (byte[])data.Clone();
                break;
        }

        OfficeProvenanceBinary.EnsureOutputWithinLimit(output.LongLength, options.EffectiveMaxOutputBytes);
        OfficeProvenanceOptions outputInspectionOptions = CreateOutputInspectionOptions(options);
        OfficeProvenanceReport after = forcedFormat.HasValue
            ? OfficeProvenanceInspector.InspectStructuredText(output, outputInspectionOptions)
            : OfficeProvenanceInspector.InspectCore(output, fileName, outputInspectionOptions);
        return OfficeProvenanceRemovalResult.CreateOwned(output, before, after, changes.AsReadOnly(), reserialized);
    }

    private static OfficeProvenanceOptions CreateInspectionOptions(OfficeProvenanceRemovalOptions source) => new OfficeProvenanceOptions {
        MaxAssetBytes = source.Limits.MaxAssetBytes,
        MaxManifestBytes = source.Limits.MaxManifestBytes,
        MaxCarriers = source.Limits.MaxCarriers,
        MaxContainerEntries = source.Limits.MaxContainerEntries,
        MaxExpandedContainerBytes = source.Limits.MaxExpandedContainerBytes,
        ProcessEmbeddedAssets = source.ProcessEmbeddedAssets && source.Limits.ProcessEmbeddedAssets,
        MaxEmbeddedAssets = Math.Min(source.MaxEmbeddedAssets, source.Limits.MaxEmbeddedAssets)
    };

    private static OfficeProvenanceOptions CreateOutputInspectionOptions(OfficeProvenanceRemovalOptions source) => new OfficeProvenanceOptions {
        MaxAssetBytes = source.EffectiveMaxOutputBytes,
        MaxManifestBytes = Math.Min(source.Limits.MaxManifestBytes, source.EffectiveMaxOutputBytes),
        MaxCarriers = source.Limits.MaxCarriers,
        MaxContainerEntries = source.Limits.MaxContainerEntries,
        MaxExpandedContainerBytes = source.Limits.MaxExpandedContainerBytes,
        ProcessEmbeddedAssets = source.ProcessEmbeddedAssets && source.Limits.ProcessEmbeddedAssets,
        MaxEmbeddedAssets = Math.Min(source.MaxEmbeddedAssets, source.Limits.MaxEmbeddedAssets)
    };

    /// <summary>Removes selected provenance from a file and atomically commits the output.</summary>
    public static OfficeProvenanceRemovalResult RemoveFile(
        string inputPath,
        string outputPath,
        OfficeProvenanceRemovalOptions? options = null) {
        if (string.IsNullOrWhiteSpace(inputPath)) throw new ArgumentException("An input path is required.", nameof(inputPath));
        if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("An output path is required.", nameof(outputPath));
        string fullInputPath = Path.GetFullPath(inputPath);
        string fullOutputPath = Path.GetFullPath(outputPath);
        options ??= new OfficeProvenanceRemovalOptions();
        OfficeProvenanceBinary.ValidateRemovalOptions(options);
        byte[] data;
        using (var stream = File.OpenRead(fullInputPath)) data = OfficeProvenanceBinary.ReadBounded(stream, options.Limits.MaxAssetBytes);
        OfficeProvenanceRemovalResult result = Remove(data, fullInputPath, options);
        OfficeFileCommit.WriteAllBytes(fullOutputPath, result.ToArray());
        return result;
    }
}
