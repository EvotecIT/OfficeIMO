namespace OfficeIMO.Pdf;

/// <summary>Options for flattening form fields and supported visual annotations in one operation.</summary>
public sealed class PdfInteractiveContentFlattenOptions {
    /// <summary>Appearance policy for AcroForm fields.</summary>
    public PdfFormFillerOptions? FormOptions { get; set; }
    /// <summary>Optional selector for visual annotations. Null selects every supported visual annotation.</summary>
    public PdfAnnotationFlattenOptions? AnnotationOptions { get; set; }
}

/// <summary>Result of flattening form fields and supported visual annotations.</summary>
public sealed class PdfInteractiveContentFlattenResult {
    private readonly byte[] _bytes;
    internal PdfInteractiveContentFlattenResult(byte[] bytes, int flattenedFormFieldCount, int flattenedAnnotationCount, PdfLoadOptions? readOptions) {
        _bytes = bytes.ToArray();
        FlattenedFormFieldCount = flattenedFormFieldCount;
        FlattenedAnnotationCount = flattenedAnnotationCount;
        ReadOptions = PdfLoadOptions.WithMinimumInputBytes(readOptions, _bytes.LongLength);
    }

    /// <summary>Rewritten PDF bytes.</summary>
    public byte[] Bytes => _bytes.ToArray();
    /// <summary>Number of readable form fields removed by flattening.</summary>
    public int FlattenedFormFieldCount { get; }
    /// <summary>Number of supported visual annotations removed by flattening.</summary>
    public int FlattenedAnnotationCount { get; }
    /// <summary>Whether any interactive object was flattened.</summary>
    public bool Applied => FlattenedFormFieldCount > 0 || FlattenedAnnotationCount > 0;
    /// <summary>Opens the flattened artifact.</summary>
    public PdfDocument ToDocument() => PdfDocument.Load(_bytes, ReadOptions);
    private PdfLoadOptions ReadOptions { get; }
}

internal static class PdfInteractiveContentFlattener {
    internal static PdfInteractiveContentFlattenResult Flatten(byte[] pdf, PdfInteractiveContentFlattenOptions? options, PdfLoadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        PdfInteractiveContentFlattenOptions effectiveOptions = options ?? new PdfInteractiveContentFlattenOptions();
        PdfDocumentInfo before = PdfInspector.Inspect(pdf, readOptions);
        int formFieldCount = before.FormFields.Count;
        byte[] output = formFieldCount == 0 ? pdf.ToArray() : PdfFormFiller.FlattenFields(pdf, effectiveOptions.FormOptions, readOptions);
        PdfLoadOptions currentReadOptions = PdfLoadOptions.WithMinimumInputBytes(readOptions, output.LongLength);
        PdfDocumentInfo afterForms = PdfInspector.Inspect(output, currentReadOptions);
        int flattenedFormFields = Math.Max(0, formFieldCount - afterForms.FormFields.Count);
        int selectedAnnotations = CountSupportedAnnotations(afterForms, effectiveOptions.AnnotationOptions);
        int flattenedAnnotations = 0;
        if (selectedAnnotations > 0) {
            PdfAnnotationEditResult annotationResult = PdfAnnotationEditor.FlattenAnnotations(output, effectiveOptions.AnnotationOptions, currentReadOptions);
            output = annotationResult.Bytes;
            flattenedAnnotations = annotationResult.AffectedAnnotationCount;
            currentReadOptions = annotationResult.OutputReadOptions;
        }
        return new PdfInteractiveContentFlattenResult(output, flattenedFormFields, flattenedAnnotations, currentReadOptions);
    }

    private static int CountSupportedAnnotations(PdfDocumentInfo info, PdfAnnotationFlattenOptions? options) => info.Annotations.Count(annotation => {
        if (annotation.Subtype is null || !PdfAnnotationFlattener.IsSupportedVisualAnnotation(annotation.Subtype)) return false;
        if (options?.ObjectNumber is int objectNumber && annotation.ObjectNumber != objectNumber) return false;
        if (options?.PageNumber is int pageNumber && annotation.PageNumber != pageNumber) return false;
        return options?.Subtype is null || string.Equals(annotation.Subtype, options.Subtype, StringComparison.OrdinalIgnoreCase);
    });
}
