namespace OfficeIMO.Pdf;

internal static partial class PdfAnnotationEditor {
    /// <summary>Flattens selected supported visual annotations through a proven full rewrite.</summary>
    public static PdfAnnotationEditResult FlattenAnnotations(byte[] pdf, PdfAnnotationFlattenOptions? options = null) =>
        FlattenAnnotations(pdf, options, readOptions: null);

    /// <summary>Flattens selected supported visual annotations using explicit read limits or credentials.</summary>
    public static PdfAnnotationEditResult FlattenAnnotations(byte[] pdf, PdfAnnotationFlattenOptions? options, PdfLoadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        PdfMutationPlan plan = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.ModifyAnnotations, readOptions);
        byte[] output = PdfAnnotationFlattener.FlattenVisualAnnotations(
            pdf,
            options,
            readOptions,
            out PdfGeneratedOutputGrowth generatedGrowth,
            out IReadOnlyDictionary<int, int> objectNumberMap,
            out int affected);
        PdfLoadOptions outputReadOptions = PdfLoadOptions.ForGeneratedOutput(readOptions, pdf, output, generatedGrowth);
        return CreateFullRewriteResult(pdf, output, affected, plan, annotationsChanged: affected > 0, readOptions: readOptions, rewrittenReadOptions: outputReadOptions, objectNumberMap: objectNumberMap);
    }

}
