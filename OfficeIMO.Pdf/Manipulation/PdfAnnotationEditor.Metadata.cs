namespace OfficeIMO.Pdf;

internal static partial class PdfAnnotationEditor {
    // Editing needs annotation identity and properties even when copying page content is forbidden.
    // The mutation planner retains that metadata privately and applies encryption/certification policy.
    internal static IReadOnlyList<PdfAnnotation> GetEditingMetadata(byte[] pdf, PdfLoadOptions? readOptions) =>
        GetAnnotationMutationDocumentInfo(RequireAnnotationMutation(pdf, readOptions)).Annotations
            .Where(annotation => annotation.Subtype != "Widget").ToArray();
}
