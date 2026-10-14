using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspace {
    /// <summary>Publishes a complete drawing as one revision; a failed or cancelled addition leaves no partial markup.</summary>
    internal Task AddMarkupAsync(IReadOnlyList<(PdfEditorTool Tool, PdfEditorGesture Gesture, PdfEditorProperties Properties)> additions,
        CancellationToken token) {
        var captured = additions.ToArray();
        if (captured.Length == 0) throw new ArgumentException("Add at least one annotation.", nameof(additions));
        if (captured.Any(item => item.Tool is PdfEditorTool.Select or PdfEditorTool.Redact or PdfEditorTool.AddText or PdfEditorTool.AddImage))
            throw new ArgumentException("Only markup belongs in a drawing.", nameof(additions));
        AnnotationMapping? mapping = null;
        return MutateBytesAsync(PdfWorkspaceOperationKind.Annotation, "Added markup",
            captured.Select(item => item.Gesture.PageNumber).Distinct().ToArray(), bytes => {
                var numbers = _annotationIdentities.Keys.ToDictionary(number => number, number => number);
                foreach (var item in captured) {
                    token.ThrowIfCancellationRequested();
                    var result = PdfEditorCommandExecutor.ApplyAnnotation(LoadDocument(bytes),
                        PdfEditorCommandFactory.Create(LoadDocument(bytes), item.Tool, item.Gesture, item.Properties));
                    if (result.AnnotationObjectNumberMap is { } remapped) {
                        numbers = numbers.Where(pair => remapped.ContainsKey(pair.Value))
                            .ToDictionary(pair => pair.Key, pair => remapped[pair.Value]);
                    } else if (result.Applied && result.MutationPlan.ExecutionMode != PdfMutationExecutionMode.AppendOnly) {
                        numbers.Clear();
                    }
                    bytes = result.Bytes;
                }
                mapping = new AnnotationMapping(numbers, false);
                return bytes;
            }, token, null, getAnnotationMapping: () => mapping);
    }

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
