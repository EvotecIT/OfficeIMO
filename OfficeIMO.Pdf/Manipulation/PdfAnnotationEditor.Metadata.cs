using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfAnnotationEditor {
    // Editing needs annotation identity and properties even when copying page content is forbidden.
    // The mutation planner retains that metadata privately and applies encryption/certification policy.
    internal static IReadOnlyList<PdfAnnotation> GetEditingMetadata(byte[] pdf, PdfLoadOptions? readOptions, CancellationToken cancellationToken = default) {
        var annotations = GetAnnotationMutationDocumentInfo(RequireAnnotationMutation(pdf, readOptions, cancellationToken: cancellationToken)).Annotations;
        var editing = new List<PdfAnnotation>();
        foreach (PdfAnnotation annotation in annotations) {
            cancellationToken.ThrowIfCancellationRequested();
            if (annotation.Subtype != "Widget") editing.Add(annotation);
        }
        cancellationToken.ThrowIfCancellationRequested();
        return editing.ToArray();
    }
}
