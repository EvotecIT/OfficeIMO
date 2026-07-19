namespace OfficeIMO.Pdf;

/// <summary>Existing-document annotation creation, update, removal, and flattening operations.</summary>
public sealed class PdfDocumentAnnotations {
    private readonly PdfDocument _document;
    internal PdfDocumentAnnotations(PdfDocument document) { _document = document; }
    /// <summary>Adds an annotation to an existing page.</summary>
    public PdfAnnotationEditResult Add(PdfAnnotationCreateOptions options) => PdfAnnotationEditor.AddAnnotation(_document.GetBytesForOperation(), options, _document.ReadOptions).WithReadOptions(_document.ReadOptions);
    /// <summary>Updates one indirect annotation.</summary>
    public PdfAnnotationEditResult Update(int objectNumber, PdfAnnotationUpdateOptions options) => PdfAnnotationEditor.UpdateAnnotation(_document.GetBytesForOperation(), objectNumber, options, _document.ReadOptions).WithReadOptions(_document.ReadOptions);
    /// <summary>Removes matching annotations.</summary>
    public PdfAnnotationEditResult Remove(PdfAnnotationRemovalOptions? options = null) => PdfAnnotationEditor.RemoveAnnotations(_document.GetBytesForOperation(), options, _document.ReadOptions).WithReadOptions(_document.ReadOptions);
    /// <summary>Flattens selected supported visual annotations.</summary>
    public PdfAnnotationEditResult Flatten(PdfAnnotationFlattenOptions? options = null) => PdfAnnotationEditor.FlattenAnnotations(_document.GetBytesForOperation(), options, _document.ReadOptions).WithReadOptions(_document.ReadOptions);
}
