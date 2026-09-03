using System.IO;
using OfficeIMO.Core.Internal;
using OfficeIMO.Provenance;

namespace OfficeIMO.Markdown;

/// <summary>Inspects and selectively removes standards-defined C2PA text carriers from Markdown.</summary>
public static class MarkdownProvenance {
    private static readonly UTF8Encoding StrictUtf8 = new UTF8Encoding(
        encoderShouldEmitUTF8Identifier: false,
        throwOnInvalidBytes: true);
    /// <summary>Inspects a Markdown string for structured and unstructured C2PA carriers.</summary>
    public static OfficeProvenanceReport Inspect(string markdown, OfficeProvenanceOptions? options = null) {
        if (markdown == null) throw new ArgumentNullException(nameof(markdown));
        options ??= new OfficeProvenanceOptions();
        OfficeProvenanceBinary.ValidateLimits(options);
        options.CancellationToken.ThrowIfCancellationRequested();
        return OfficeProvenanceInspector.InspectStructuredText(
            EncodeUtf8Bounded(markdown, options.MaxAssetBytes),
            options);
    }

    /// <summary>Inspects a bounded Markdown file.</summary>
    public static OfficeProvenanceReport InspectFile(string filePath, OfficeProvenanceOptions? options = null) {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("A file path is required.", nameof(filePath));
        options ??= new OfficeProvenanceOptions();
        OfficeProvenanceBinary.ValidateLimits(options);
        byte[] input;
        using (var stream = File.OpenRead(Path.GetFullPath(filePath))) {
            input = OfficeProvenanceBinary.ReadBounded(stream, options.MaxAssetBytes, options.CancellationToken);
        }
        string markdown = DecodeFileText(input, out _, out _);
        return OfficeProvenanceInspector.InspectStructuredText(
            EncodeDecodedFileTextAsUtf8(markdown),
            options);
    }

    /// <summary>Removes selected C2PA text carriers from Markdown content.</summary>
    public static OfficeProvenanceRemovalResult Remove(
        string markdown,
        OfficeProvenanceRemovalOptions? options = null) {
        if (markdown == null) throw new ArgumentNullException(nameof(markdown));
        options ??= new OfficeProvenanceRemovalOptions();
        OfficeProvenanceBinary.ValidateRemovalOptions(options);
        options.Limits.CancellationToken.ThrowIfCancellationRequested();
        return OfficeProvenanceRemover.RemoveStructuredText(
            EncodeUtf8Bounded(markdown, options.Limits.MaxAssetBytes),
            "document.md",
            options);
    }

    /// <summary>Removes selected C2PA text carriers and atomically writes a Markdown file.</summary>
    public static OfficeProvenanceRemovalResult RemoveFile(
        string inputPath,
        string outputPath,
        OfficeProvenanceRemovalOptions? options = null) {
        if (string.IsNullOrWhiteSpace(inputPath)) throw new ArgumentException("An input path is required.", nameof(inputPath));
        if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("An output path is required.", nameof(outputPath));
        options ??= new OfficeProvenanceRemovalOptions();
        byte[] input;
        using (var stream = File.OpenRead(Path.GetFullPath(inputPath))) {
            input = OfficeProvenanceBinary.ReadBounded(stream, options.Limits.MaxAssetBytes, options.Limits.CancellationToken);
        }
        string markdown = DecodeFileText(input, out Encoding encoding, out bool hadPreamble);
        byte[] utf8Input = EncodeDecodedFileTextAsUtf8(markdown);
        OfficeProvenanceRemovalResult utf8Result = OfficeProvenanceRemover.RemoveStructuredText(
            utf8Input,
            inputPath,
            CloneForTranscodedProcessing(options, Math.Max(1L, utf8Input.LongLength)));
        if (!utf8Result.WasChanged) {
            OfficeProvenanceBinary.EnsureOutputWithinLimit(input.LongLength, options.EffectiveMaxOutputBytes);
            OfficeFileCommit.WriteAllBytes(Path.GetFullPath(outputPath), input);
            return new OfficeProvenanceRemovalResult(
                input, utf8Result.Before, utf8Result.After, utf8Result.Changes, wasReserialized: false);
        }
        string cleaned = Encoding.UTF8.GetString(utf8Result.ToArray());
        byte[] output = EncodeFileText(cleaned, encoding, hadPreamble, options.EffectiveMaxOutputBytes);
        OfficeFileCommit.WriteAllBytes(Path.GetFullPath(outputPath), output);
        IReadOnlyList<OfficeProvenanceChange> physicalChanges = utf8Result.Changes;
        if (encoding.CodePage != Encoding.UTF8.CodePage) {
            physicalChanges = utf8Result.Changes
                .Select(change => new OfficeProvenanceChange(change.Carrier, change.Location, removedBytes: 0))
                .ToArray();
        }
        return new OfficeProvenanceRemovalResult(
            output, utf8Result.Before, utf8Result.After, physicalChanges, wasReserialized: true);
    }

    private static OfficeProvenanceRemovalOptions CloneForTranscodedProcessing(
        OfficeProvenanceRemovalOptions source,
        long maximumTranscodedBytes) {
        var clone = new OfficeProvenanceRemovalOptions {
            RemoveC2paManifests = source.RemoveC2paManifests,
            RemoveExternalC2paReferences = source.RemoveExternalC2paReferences,
            RemoveAiSourceMetadata = source.RemoveAiSourceMetadata,
            RequireStructurallyValidCarrier = source.RequireStructurallyValidCarrier,
            SignatureMutationPolicy = source.SignatureMutationPolicy,
            ProcessEmbeddedAssets = source.ProcessEmbeddedAssets,
            MaxEmbeddedAssets = source.MaxEmbeddedAssets,
            MaxOutputBytes = maximumTranscodedBytes
        };
        clone.Limits.MaxAssetBytes = Math.Max(source.Limits.MaxAssetBytes, maximumTranscodedBytes);
        clone.Limits.MaxManifestBytes = source.Limits.MaxManifestBytes;
        clone.Limits.MaxCarriers = source.Limits.MaxCarriers;
        clone.Limits.MaxContainerEntries = source.Limits.MaxContainerEntries;
        clone.Limits.MaxExpandedContainerBytes = source.Limits.MaxExpandedContainerBytes;
        clone.Limits.CancellationToken = source.Limits.CancellationToken;
        clone.Limits.ProcessEmbeddedAssets = source.Limits.ProcessEmbeddedAssets;
        clone.Limits.MaxEmbeddedAssets = source.Limits.MaxEmbeddedAssets;
        return clone;
    }

    private static string DecodeFileText(byte[] data, out Encoding encoding, out bool hadPreamble) {
        Encoding[] candidates = {
            new UTF32Encoding(bigEndian: true, byteOrderMark: true, throwOnInvalidCharacters: true),
            new UTF32Encoding(bigEndian: false, byteOrderMark: true, throwOnInvalidCharacters: true),
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: true, throwOnInvalidBytes: true),
            new UnicodeEncoding(bigEndian: true, byteOrderMark: true, throwOnInvalidBytes: true),
            new UnicodeEncoding(bigEndian: false, byteOrderMark: true, throwOnInvalidBytes: true)
        };
        encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true);
        int offset = 0;
        foreach (Encoding candidate in candidates) {
            byte[] preamble = candidate.GetPreamble();
            if (!StartsWith(data, preamble)) continue;
            encoding = candidate;
            offset = preamble.Length;
            break;
        }
        hadPreamble = offset != 0;
        try {
            return encoding.GetString(data, offset, data.Length - offset);
        } catch (DecoderFallbackException exception) {
            throw new InvalidDataException("The Markdown document contains invalid encoded text.", exception);
        }
    }

    private static bool StartsWith(byte[] data, byte[] prefix) {
        if (prefix.Length == 0 || data.Length < prefix.Length) return false;
        for (int index = 0; index < prefix.Length; index++) {
            if (data[index] != prefix[index]) return false;
        }
        return true;
    }

    private static byte[] EncodeFileText(string text, Encoding encoding, bool includePreamble, long maximumBytes) {
        byte[] preamble = includePreamble ? encoding.GetPreamble() : Array.Empty<byte>();
        int bodyLength = encoding.GetByteCount(text);
        if (bodyLength > maximumBytes - preamble.Length) {
            throw OfficeProvenanceLimitException.CreateOutput(
                $"The rewritten Markdown document exceeds the configured output limit of {maximumBytes} bytes.");
        }
        byte[] body = encoding.GetBytes(text);
        if (preamble.Length == 0) return body;
        byte[] output = new byte[preamble.Length + body.Length];
        Buffer.BlockCopy(preamble, 0, output, 0, preamble.Length);
        Buffer.BlockCopy(body, 0, output, preamble.Length, body.Length);
        return output;
    }

    private static byte[] EncodeUtf8Bounded(string text, long maximumBytes) {
        try {
            int byteCount = StrictUtf8.GetByteCount(text);
            if (byteCount > maximumBytes) {
                throw new InvalidDataException("The Markdown document exceeds the configured asset limit after decoding.");
            }
            return StrictUtf8.GetBytes(text);
        } catch (EncoderFallbackException exception) {
            throw new InvalidDataException("The Markdown document contains invalid Unicode text.", exception);
        }
    }

    private static byte[] EncodeDecodedFileTextAsUtf8(string text) {
        try {
            return StrictUtf8.GetBytes(text);
        } catch (EncoderFallbackException exception) {
            throw new InvalidDataException("The Markdown document contains invalid Unicode text.", exception);
        }
    }
}
