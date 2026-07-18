namespace OfficeIMO.Pdf;

internal static partial class PdfAnnotationEditor {
    /// <summary>Flattens selected supported visual annotations through a proven full rewrite.</summary>
    public static PdfAnnotationEditResult FlattenAnnotations(byte[] pdf, PdfAnnotationFlattenOptions? options = null) =>
        FlattenAnnotations(pdf, options, readOptions: null);

    /// <summary>Flattens selected supported visual annotations using explicit read limits or credentials.</summary>
    public static PdfAnnotationEditResult FlattenAnnotations(byte[] pdf, PdfAnnotationFlattenOptions? options, PdfReadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        PdfMutationPlan plan = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.ModifyAnnotations, readOptions);
        int before = CountSelectedAnnotations(PdfInspector.Inspect(pdf, readOptions), options);
        byte[] output = PdfAnnotationFlattener.FlattenVisualAnnotations(pdf, options, readOptions);
        int after = CountSelectedAnnotations(PdfInspector.Inspect(output), options);
        int affected = Math.Max(0, before - after);
        return CreateFullRewriteResult(pdf, output, affected, plan, annotationsChanged: affected > 0);
    }

    private static int CountSelectedAnnotations(PdfDocumentInfo info, PdfAnnotationFlattenOptions? options) {
        IEnumerable<PdfAnnotation> values = info.Annotations;
        if (options?.ObjectNumber != null) values = values.Where(annotation => annotation.ObjectNumber == options.ObjectNumber);
        if (options?.PageNumber != null) values = values.Where(annotation => annotation.PageNumber == options.PageNumber);
        if (options?.Subtype != null) values = values.Where(annotation => string.Equals(annotation.Subtype, options.Subtype, StringComparison.OrdinalIgnoreCase));
        return values.Count();
    }
}
