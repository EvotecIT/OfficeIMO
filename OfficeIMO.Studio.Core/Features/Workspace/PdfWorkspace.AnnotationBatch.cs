using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspace {
    internal Task EditAnnotationsAsync(IReadOnlyList<PdfEditorSelection> selections, long revision,
        Func<PdfDocumentAnnotations, IReadOnlyList<int>, PdfAnnotationEditResult> edit, string description,
        CancellationToken token, IProgress<PdfWorkspaceProgress>? progress = null) {
        PdfEditorSelection[] captured = selections.ToArray();
        if (captured.Length == 0 || captured.Any(selection => selection.Kind != PdfEditorSelectionKind.Annotation || selection.ObjectNumber is null))
            throw new ArgumentException("Select indirect annotations before editing them.", nameof(selections));
        return MutateAnnotationBytesAsync(PdfWorkspaceOperationKind.Annotation, description,
            captured.Select(selection => selection.PageNumber).Distinct().ToArray(), bytes => {
                if (Revision != revision) throw new InvalidOperationException("The document changed. Select the annotations again.");
                return edit(LoadDocument(bytes).Annotations, captured.Select(selection => selection.ObjectNumber!.Value).ToArray());
            }, token, progress);
    }
}
