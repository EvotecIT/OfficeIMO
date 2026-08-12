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
            throw new InvalidDataException($"The asset exceeds the configured limit of {options.MaxAssetBytes} bytes.");
        }
        return InspectCore(data, fileName, options);
    }

    internal static OfficeProvenanceReport InspectCore(byte[] data, string? fileName, OfficeProvenanceOptions options) {
        OfficeProvenanceAssetFormat format = DetectFormat(data, fileName);
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

    internal static OfficeProvenanceAssetFormat DetectFormat(byte[] data, string? fileName) {
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
        if (HasStructuredTextExtension(fileName)) return OfficeProvenanceAssetFormat.StructuredText;
        if (LooksLikeSvg(data, fileName)) return OfficeProvenanceAssetFormat.Svg;
        if (OfficeProvenanceText.HasUnstructuredWrapperPrefix(data)) return OfficeProvenanceAssetFormat.UnstructuredText;
        if (OfficeProvenanceText.HasStructuredDelimiter(data)) return OfficeProvenanceAssetFormat.StructuredText;
        return OfficeProvenanceAssetFormat.Unknown;
    }

    private static bool LooksLikeSvg(byte[] data, string? fileName) {
        string extension = Path.GetExtension(fileName ?? string.Empty);
        if (extension.Equals(".svg", StringComparison.OrdinalIgnoreCase)) return true;
        if (data.Length == 0) return false;
        try {
            var settings = new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = data.LongLength,
                MaxCharactersFromEntities = 0
            };
            using var stream = new MemoryStream(data, writable: false);
            using XmlReader reader = XmlReader.Create(stream, settings);
            while (reader.Read()) {
                if (reader.NodeType != XmlNodeType.Element) continue;
                return reader.LocalName.Equals("svg", StringComparison.OrdinalIgnoreCase) &&
                    reader.NamespaceURI.Equals("http://www.w3.org/2000/svg", StringComparison.Ordinal);
            }
            return false;
        } catch (XmlException) {
            return false;
        }
    }

    private static bool HasStructuredTextExtension(string? fileName) {
        string extension = Path.GetExtension(fileName ?? string.Empty).ToLowerInvariant();
        return extension is ".md" or ".markdown" or ".txt" or ".yaml" or ".yml" or ".toml" or ".ini" or ".json" or ".xml" or ".adoc" or ".asciidoc" or ".tex";
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
        OfficeProvenanceBinary.ValidateLimits(options.Limits);
        if (options.MaxEmbeddedAssets <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxEmbeddedAssets));
        if (data.LongLength > options.Limits.MaxAssetBytes) {
            throw new InvalidDataException($"The asset exceeds the configured limit of {options.Limits.MaxAssetBytes} bytes.");
        }

        OfficeProvenanceOptions inspectionOptions = CreateInspectionOptions(options);
        OfficeProvenanceReport before = OfficeProvenanceInspector.InspectCore(data, fileName, inspectionOptions);
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
                output = OfficeProvenanceGif.Remove(data, options, changes);
                break;
            case OfficeProvenanceAssetFormat.Tiff:
                output = OfficeProvenanceTiff.Remove(data, options, changes, out reserialized);
                break;
            case OfficeProvenanceAssetFormat.Svg:
                output = OfficeProvenanceSvg.Remove(data, options, changes, out reserialized);
                break;
            case OfficeProvenanceAssetFormat.ZipPackage:
                output = OfficeProvenanceZip.Remove(data, options, changes, out reserialized);
                break;
            case OfficeProvenanceAssetFormat.StructuredText:
            case OfficeProvenanceAssetFormat.UnstructuredText:
                output = OfficeProvenanceText.Remove(data, options, changes);
                break;
            default:
                output = (byte[])data.Clone();
                break;
        }

        OfficeProvenanceReport after = OfficeProvenanceInspector.InspectCore(output, fileName, inspectionOptions);
        return new OfficeProvenanceRemovalResult(output, before, after, changes.AsReadOnly(), reserialized);
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
        OfficeProvenanceBinary.ValidateLimits(options.Limits);
        byte[] data;
        using (var stream = File.OpenRead(fullInputPath)) data = OfficeProvenanceBinary.ReadBounded(stream, options.Limits.MaxAssetBytes);
        OfficeProvenanceRemovalResult result = Remove(data, fullInputPath, options);
        OfficeFileCommit.WriteAllBytes(fullOutputPath, result.ToArray());
        return result;
    }
}
