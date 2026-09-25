using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Provenance;
using OfficeIMO.Word;

namespace OfficeIMO.Workflows;

/// <summary>Bounded provenance operations for memory-only hosts. Never reads files or follows external references.</summary>
public static class OfficeProvenanceBufferWorkflow {
    /// <summary>File families qualified for the memory-only workflow. Other families use the local workflow runner.</summary>
    public static IReadOnlyList<string> SupportedExtensions => OfficeProvenanceWorkflowCatalog.MemoryOnlyExtensions;

    /// <summary>Inspects the supplied bytes through their format owner. Structural findings are not authenticity verification.</summary>
    public static OfficeProvenanceReport Inspect(byte[] data, string fileName, OfficeProvenanceOptions? options = null) {
        QualifiedBufferInput input = ValidateInput(data, fileName);
        options = (options ?? new OfficeProvenanceOptions()).ForStandardOpenXmlDocument();
        OfficeProvenanceReport report = input.Capability.Owner switch {
            OfficeProvenanceWorkflowAdapter.ProvenanceOwner.Word => WordDocument.InspectProvenance(data, fileName, options),
            OfficeProvenanceWorkflowAdapter.ProvenanceOwner.Excel => ExcelDocument.InspectProvenance(data, fileName, options),
            OfficeProvenanceWorkflowAdapter.ProvenanceOwner.PowerPoint => PowerPointPresentation.InspectProvenance(data, fileName, options),
            OfficeProvenanceWorkflowAdapter.ProvenanceOwner.Pdf => PdfProvenance.Inspect(data, options),
            _ => OfficeProvenanceInspector.Inspect(data, fileName, options)
        };
        EnsureFormatMatches(input.Format, report);
        return report;
    }

    /// <summary>Creates a separate cleaned result and re-inspects its bytes. Signed or ambiguous inputs retain the owner's mutation policy.</summary>
    public static OfficeProvenanceRemovalResult Remove(byte[] data, string fileName, OfficeProvenanceRemovalOptions? options = null) {
        QualifiedBufferInput input = ValidateInput(data, fileName);
        options = OfficeProvenancePackageMutation.ForStandardOpenXmlDocument(options ?? new OfficeProvenanceRemovalOptions());
        var result = input.Capability.Owner switch {
            OfficeProvenanceWorkflowAdapter.ProvenanceOwner.Word => WordDocument.RemoveProvenance(data, fileName, options),
            OfficeProvenanceWorkflowAdapter.ProvenanceOwner.Excel => ExcelDocument.RemoveProvenance(data, fileName, options),
            OfficeProvenanceWorkflowAdapter.ProvenanceOwner.PowerPoint => PowerPointPresentation.RemoveProvenance(data, fileName, options),
            OfficeProvenanceWorkflowAdapter.ProvenanceOwner.Pdf => PdfProvenance.Remove(data, options),
            _ => RemoveImage(data, fileName, input.Format, options)
        };
        var outputLimits = OfficeProvenanceRemover.CreateOutputInspectionOptions(options);
        long remainingBytes = options.Limits.MaxExpandedContainerBytes;
        OfficeWorkflowRunner.ConsumeExpandedProcessingBytes(ref remainingBytes, result.ExpandedInspectionBytes);
        outputLimits.MaxExpandedContainerBytes = Math.Max(1, remainingBytes);
        var reopened = Inspect(result.ToArray(), fileName, outputLimits);
        OfficeWorkflowRunner.ConsumeExpandedProcessingBytes(ref remainingBytes, reopened.ExpandedInspectionBytes);
        OfficeWorkflowRunner.EnsureEquivalent(result.After, reopened);
        return result;
    }

    private static OfficeProvenanceRemovalResult RemoveImage(
        byte[] data,
        string name,
        OfficeProvenanceWorkflowFormat format,
        OfficeProvenanceRemovalOptions options) {
        EnsureFormatMatches(format, OfficeProvenanceInspector.Inspect(data, name, options.Limits));
        return OfficeProvenanceRemover.Remove(data, name, options);
    }
    private static QualifiedBufferInput ValidateInput(byte[] data, string name) {
        ArgumentNullException.ThrowIfNull(data);
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        OfficeProvenanceWorkflowCapability capability = OfficeProvenanceWorkflowCatalog.FindByPath(name) ??
            throw new NotSupportedException("This OfficeIMO format is not qualified for the memory-only provenance workflow.");
        OfficeProvenanceWorkflowFormat format = OfficeProvenanceWorkflowCatalog.FindMemoryOnlyFormatByPath(name) ??
            throw new NotSupportedException("This OfficeIMO format is not qualified for the memory-only provenance workflow.");
        return new QualifiedBufferInput(capability, format);
    }

    private static void EnsureFormatMatches(OfficeProvenanceWorkflowFormat format, OfficeProvenanceReport report) {
        if (!format.AssetFormats.Contains(report.Format)) {
            throw new InvalidDataException(
                $"The file contents identify as {report.Format}, which does not match the registered {format.Extension} format.");
        }
    }

    private sealed record QualifiedBufferInput(
        OfficeProvenanceWorkflowCapability Capability,
        OfficeProvenanceWorkflowFormat Format);
}
