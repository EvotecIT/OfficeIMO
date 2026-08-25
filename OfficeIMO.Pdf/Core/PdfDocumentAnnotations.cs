namespace OfficeIMO.Pdf;

/// <summary>Existing-document annotation creation, update, removal, and flattening operations.</summary>
public sealed class PdfDocumentAnnotations {
    private readonly PdfDocument _document;
    internal PdfDocumentAnnotations(PdfDocument document) { _document = document; }
    /// <summary>Adds an annotation to an existing page.</summary>
    public PdfAnnotationEditResult Add(PdfAnnotationCreateOptions options) => PdfAnnotationEditor.AddAnnotation(_document.GetBytesForOperation(), options, _document.ReadOptions);
    /// <summary>Updates one indirect annotation.</summary>
    public PdfAnnotationEditResult Update(int objectNumber, PdfAnnotationUpdateOptions options) => PdfAnnotationEditor.UpdateAnnotation(_document.GetBytesForOperation(), objectNumber, options, _document.ReadOptions);
    /// <summary>Adds a reply to one indirect annotation.</summary>
    public PdfAnnotationEditResult AddReply(int parentObjectNumber, string contents, PdfAnnotationReplyOptions? options = null) => PdfAnnotationReviewEditor.AddReply(_document.GetBytesForOperation(), parentObjectNumber, contents, options, _document.ReadOptions);
    /// <summary>Sets the standard review state on one indirect annotation.</summary>
    public PdfAnnotationEditResult SetReviewState(int objectNumber, PdfAnnotationReviewState state, PdfMutationExecutionPreference executionPreference = PdfMutationExecutionPreference.Automatic, bool allowResidualDataInAppendOnly = false) => PdfAnnotationReviewEditor.SetState(_document.GetBytesForOperation(), objectNumber, state, executionPreference, allowResidualDataInAppendOnly, _document.ReadOptions);
    /// <summary>Reads annotation reply threads and review states.</summary>
    public PdfAnnotationReviewCatalog GetReviewCatalog() => PdfAnnotationReviewCatalog.Read(_document.GetBytesForOperation(), _document.ReadOptions);
    /// <summary>Removes matching annotations.</summary>
    public PdfAnnotationEditResult Remove(PdfAnnotationRemovalOptions? options = null) => PdfAnnotationEditor.RemoveAnnotations(_document.GetBytesForOperation(), options, _document.ReadOptions);
    /// <summary>Flattens selected supported visual annotations.</summary>
    public PdfAnnotationEditResult Flatten(PdfAnnotationFlattenOptions? options = null) => PdfAnnotationEditor.FlattenAnnotations(_document.GetBytesForOperation(), options, _document.ReadOptions);
}
