using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Provenance;
using OfficeIMO.Word;

namespace OfficeIMO.Workflows;

/// <summary>Bounded provenance operations for memory-only hosts. Never reads files or follows external references.</summary>
public static class OfficeProvenanceBufferWorkflow {
    /// <summary>File families qualified for the memory-only workflow. Other families use the local workflow runner.</summary>
    public static IReadOnlyList<string> SupportedExtensions { get; } = Array.AsReadOnly(new[] {
        ".jpg", ".jpeg", ".png", ".webp", ".pdf", ".docx", ".xlsx", ".pptx"
    });

    /// <summary>Inspects the supplied bytes through their format owner. Structural findings are not authenticity verification.</summary>
    public static OfficeProvenanceReport Inspect(byte[] data, string fileName, OfficeProvenanceOptions? options = null) {
        ValidateInput(data, fileName);
        return OfficeProvenanceWorkflowAdapter.ResolveByPath(fileName) switch {
            OfficeProvenanceWorkflowAdapter.ProvenanceOwner.Word => WordDocument.InspectProvenance(data, fileName, options),
            OfficeProvenanceWorkflowAdapter.ProvenanceOwner.Excel => ExcelDocument.InspectProvenance(data, fileName, options),
            OfficeProvenanceWorkflowAdapter.ProvenanceOwner.PowerPoint => PowerPointPresentation.InspectProvenance(data, fileName, options),
            OfficeProvenanceWorkflowAdapter.ProvenanceOwner.Pdf => PdfProvenance.Inspect(data, options),
            _ => InspectImage(data, fileName, options)
        };
    }

    /// <summary>Creates a separate cleaned result and re-inspects its bytes. Signed or ambiguous inputs retain the owner's mutation policy.</summary>
    public static OfficeProvenanceRemovalResult Remove(byte[] data, string fileName, OfficeProvenanceRemovalOptions? options = null) {
        ValidateInput(data, fileName);
        options ??= new OfficeProvenanceRemovalOptions();
        var result = OfficeProvenanceWorkflowAdapter.ResolveByPath(fileName) switch {
            OfficeProvenanceWorkflowAdapter.ProvenanceOwner.Word => WordDocument.RemoveProvenance(data, fileName, options),
            OfficeProvenanceWorkflowAdapter.ProvenanceOwner.Excel => ExcelDocument.RemoveProvenance(data, fileName, options),
            OfficeProvenanceWorkflowAdapter.ProvenanceOwner.PowerPoint => PowerPointPresentation.RemoveProvenance(data, fileName, options),
            OfficeProvenanceWorkflowAdapter.ProvenanceOwner.Pdf => PdfProvenance.Remove(data, options),
            _ => RemoveImage(data, fileName, options)
        };
        var outputLimits = OfficeProvenanceRemover.CreateOutputInspectionOptions(options);
        long remainingBytes = options.Limits.MaxExpandedContainerBytes;
        OfficeWorkflowRunner.ConsumeExpandedProcessingBytes(ref remainingBytes, result.Before.ExpandedInspectionBytes);
        if (!ReferenceEquals(result.Before, result.After))
            OfficeWorkflowRunner.ConsumeExpandedProcessingBytes(ref remainingBytes, result.After.ExpandedInspectionBytes);
        outputLimits.MaxExpandedContainerBytes = Math.Max(1, remainingBytes);
        var reopened = Inspect(result.ToArray(), fileName, outputLimits);
        OfficeWorkflowRunner.ConsumeExpandedProcessingBytes(ref remainingBytes, reopened.ExpandedInspectionBytes);
        OfficeWorkflowRunner.EnsureEquivalent(result.After, reopened);
        return result;
    }

    private static OfficeProvenanceReport InspectImage(byte[] data, string name, OfficeProvenanceOptions? options) {
        var report = OfficeProvenanceInspector.Inspect(data, name, options);
        var expected = Path.GetExtension(name).ToLowerInvariant() switch {
            ".png" => OfficeProvenanceAssetFormat.Png, ".webp" => OfficeProvenanceAssetFormat.Webp,
            _ => OfficeProvenanceAssetFormat.Jpeg
        };
        if (report.Format != expected) throw new InvalidDataException("The file contents do not match the selected image format.");
        return report;
    }
    private static OfficeProvenanceRemovalResult RemoveImage(byte[] data, string name, OfficeProvenanceRemovalOptions options) {
        InspectImage(data, name, options.Limits);
        return OfficeProvenanceRemover.Remove(data, name, options);
    }
    private static void ValidateInput(byte[] data, string name) {
        ArgumentNullException.ThrowIfNull(data);
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        if (!SupportedExtensions.Contains(Path.GetExtension(name), StringComparer.OrdinalIgnoreCase))
            throw new NotSupportedException("This file family is not supported by the memory-only provenance workflow.");
    }
}
