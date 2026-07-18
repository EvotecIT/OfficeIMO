namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    internal PdfPipelineReport AppendOutputStep(
        string operation,
        PdfArtifactSnapshot? output,
        TimeSpan duration,
        Exception? exception = null) {
        IReadOnlyList<string> diagnostics = exception is null
            ? Array.Empty<string>()
            : PdfOutputDiagnostics.BuildExceptionDiagnostics(exception);
        var step = new PdfPipelineStep(
            PdfPipelineStepKind.Output,
            operation,
            exception is null,
            _pipeline.Output,
            output,
            duration,
            mutationOperation: null,
            executionMode: null,
            diagnostics);
        return _pipeline.Append(step);
    }

    private static bool IsAppendOnly(byte[] input, byte[] output) {
        if (output.Length <= input.Length) {
            return false;
        }

        for (int i = 0; i < input.Length; i++) {
            if (input[i] != output[i]) {
                return false;
            }
        }

        return true;
    }

    private static string NormalizeOperationName(string operationName) {
        if (string.IsNullOrWhiteSpace(operationName)) {
            return "Mutation";
        }

        return operationName;
    }

    private static PdfMutationOperation? ResolveMutationOperation(string operationName) {
        switch (operationName) {
            case "Fill":
            case "AppendRevision":
            case "ImportData":
            case "ImportXfdf":
                return PdfMutationOperation.FillFormFields;
            case "Flatten":
                return PdfMutationOperation.FlattenFormFields;
            case "FillAndFlatten":
                return PdfMutationOperation.FillAndFlattenFormFields;
            case "Extract":
            case "Split":
                return PdfMutationOperation.ExtractPages;
            case "Append":
            case "Prepend":
            case "Insert":
            case "Import":
            case "ImportPages":
            case "MergeWith":
                return PdfMutationOperation.MergeDocuments;
            case "Delete":
            case "Reorder":
            case "Duplicate":
            case "Move":
            case "Rotate":
            case "Crop":
            case "Resize":
            case "SetPageBox":
            case "CropAndTranslate":
            case "DestructiveCrop":
                return PdfMutationOperation.ModifyPageTree;
            case "Text":
            case "TextWatermark":
            case "Image":
            case "ImageWatermark":
            case "OverlayPage":
            case "UnderlayPage":
                return PdfMutationOperation.ModifyPageContent;
            case "FlattenVisualAnnotations":
                return PdfMutationOperation.ModifyAnnotations;
            case "SetMetadata":
            case "UpdateMetadata":
            case "ReplaceMetadata":
            case "AppendMetadataRevision":
                return PdfMutationOperation.UpdateMetadata;
            case "SynchronizeMetadata":
                return PdfMutationOperation.SynchronizeMetadata;
            case "ApplyRedactions":
                return PdfMutationOperation.Redact;
            default:
                return null;
        }
    }
}
