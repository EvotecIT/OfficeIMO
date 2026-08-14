using System.IO;
using OfficeIMO.Core.Internal;
using OfficeIMO.Provenance;

namespace OfficeIMO.Markdown;

/// <summary>Inspects and selectively removes standards-defined C2PA text carriers from Markdown.</summary>
public static class MarkdownProvenance {
    /// <summary>Inspects a Markdown string for structured and unstructured C2PA carriers.</summary>
    public static OfficeProvenanceReport Inspect(string markdown, OfficeProvenanceOptions? options = null) {
        if (markdown == null) throw new ArgumentNullException(nameof(markdown));
        options ??= new OfficeProvenanceOptions();
        OfficeProvenanceBinary.ValidateLimits(options);
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
            input = OfficeProvenanceBinary.ReadBounded(stream, options.MaxAssetBytes);
        }
        string markdown = DecodeFileText(input, out _, out _);
        return OfficeProvenanceInspector.InspectStructuredText(
            EncodeUtf8Bounded(markdown, options.MaxAssetBytes),
            options);
    }

    /// <summary>Removes selected C2PA text carriers from Markdown content.</summary>
    public static OfficeProvenanceRemovalResult Remove(
        string markdown,
        OfficeProvenanceRemovalOptions? options = null) {
        if (markdown == null) throw new ArgumentNullException(nameof(markdown));
        options ??= new OfficeProvenanceRemovalOptions();
        OfficeProvenanceBinary.ValidateLimits(options.Limits);
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
            input = OfficeProvenanceBinary.ReadBounded(stream, options.Limits.MaxAssetBytes);
        }
        string markdown = DecodeFileText(input, out Encoding encoding, out bool hadPreamble);
        OfficeProvenanceRemovalResult utf8Result = OfficeProvenanceRemover.RemoveStructuredText(
            EncodeUtf8Bounded(markdown, options.Limits.MaxAssetBytes), inputPath, options);
        if (!utf8Result.WasChanged) {
            OfficeFileCommit.WriteAllBytes(Path.GetFullPath(outputPath), input);
            return new OfficeProvenanceRemovalResult(
                input, utf8Result.Before, utf8Result.After, utf8Result.Changes, wasReserialized: false);
        }
        string cleaned = Encoding.UTF8.GetString(utf8Result.ToArray());
        byte[] output = EncodeFileText(cleaned, encoding, hadPreamble, options.Limits.MaxAssetBytes);
        OfficeFileCommit.WriteAllBytes(Path.GetFullPath(outputPath), output);
        return new OfficeProvenanceRemovalResult(
            output, utf8Result.Before, utf8Result.After, utf8Result.Changes, wasReserialized: true);
    }

    private static string DecodeFileText(byte[] data, out Encoding encoding, out bool hadPreamble) {
        using var stream = new MemoryStream(data, writable: false);
        using var reader = new StreamReader(stream, new UTF8Encoding(false), detectEncodingFromByteOrderMarks: true, bufferSize: 1024, leaveOpen: false);
        string text = reader.ReadToEnd();
        encoding = reader.CurrentEncoding;
        byte[] preamble = encoding.GetPreamble();
        hadPreamble = preamble.Length != 0 && data.Length >= preamble.Length &&
            preamble.SequenceEqual(data.Take(preamble.Length));
        return text;
    }

    private static byte[] EncodeFileText(string text, Encoding encoding, bool includePreamble, long maximumBytes) {
        byte[] preamble = includePreamble ? encoding.GetPreamble() : Array.Empty<byte>();
        int bodyLength = encoding.GetByteCount(text);
        if (bodyLength > maximumBytes - preamble.Length) {
            throw new InvalidDataException("The rewritten Markdown document exceeds the configured asset limit.");
        }
        byte[] body = encoding.GetBytes(text);
        if (preamble.Length == 0) return body;
        byte[] output = new byte[preamble.Length + body.Length];
        Buffer.BlockCopy(preamble, 0, output, 0, preamble.Length);
        Buffer.BlockCopy(body, 0, output, preamble.Length, body.Length);
        return output;
    }

    private static byte[] EncodeUtf8Bounded(string text, long maximumBytes) {
        int byteCount = Encoding.UTF8.GetByteCount(text);
        if (byteCount > maximumBytes) {
            throw new InvalidDataException("The Markdown document exceeds the configured asset limit after decoding.");
        }
        return Encoding.UTF8.GetBytes(text);
    }
}
