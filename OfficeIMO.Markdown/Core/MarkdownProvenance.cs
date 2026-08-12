using System.IO;
using OfficeIMO.Core.Internal;
using OfficeIMO.Provenance;

namespace OfficeIMO.Markdown;

/// <summary>Inspects and selectively removes standards-defined C2PA text carriers from Markdown.</summary>
public static class MarkdownProvenance {
    /// <summary>Inspects a Markdown string for structured and unstructured C2PA carriers.</summary>
    public static OfficeProvenanceReport Inspect(string markdown, OfficeProvenanceOptions? options = null) {
        if (markdown == null) throw new ArgumentNullException(nameof(markdown));
        return OfficeProvenanceInspector.Inspect(Encoding.UTF8.GetBytes(markdown), "document.md", options);
    }

    /// <summary>Inspects a bounded Markdown file.</summary>
    public static OfficeProvenanceReport InspectFile(string filePath, OfficeProvenanceOptions? options = null) =>
        OfficeProvenanceInspector.InspectFile(filePath, options);

    /// <summary>Removes selected C2PA text carriers from Markdown content.</summary>
    public static OfficeProvenanceRemovalResult Remove(
        string markdown,
        OfficeProvenanceRemovalOptions? options = null) {
        if (markdown == null) throw new ArgumentNullException(nameof(markdown));
        return OfficeProvenanceRemover.Remove(Encoding.UTF8.GetBytes(markdown), "document.md", options);
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
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(input, inputPath, options);
        OfficeFileCommit.WriteAllBytes(Path.GetFullPath(outputPath), result.ToArray());
        return result;
    }
}
