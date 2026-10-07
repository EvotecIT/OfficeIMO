namespace OfficeIMO.Pdf;

/// <summary>Existing-document annotation creation, update, removal, and flattening operations.</summary>
public sealed class PdfDocumentAnnotations {
    private readonly PdfDocument _document;
    internal PdfDocumentAnnotations(PdfDocument document) { _document = document; }
    /// <summary>Adds an annotation to an existing page.</summary>
    public PdfAnnotationEditResult Add(PdfAnnotationCreateOptions options) => PdfAnnotationEditor.AddAnnotation(_document.GetBytesForOperation(), options, _document.ReadOptions);
    /// <summary>Adds a text or image-backed Stamp annotation without changing page content.</summary>
    public PdfAnnotationEditResult AddStamp(PdfStampAnnotationOptions options) => PdfAnnotationEditor.AddStampAnnotation(_document.GetBytesForOperation(), options, _document.ReadOptions);
    /// <summary>Updates one indirect annotation.</summary>
    public PdfAnnotationEditResult Update(int objectNumber, PdfAnnotationUpdateOptions options) => PdfAnnotationEditor.UpdateAnnotation(_document.GetBytesForOperation(), objectNumber, options, _document.ReadOptions);
    /// <summary>Moves one annotation and all supported subtype geometry by a user-space offset.</summary>
    public PdfAnnotationEditResult Move(int objectNumber, double deltaX, double deltaY) => PdfAnnotationEditor.MoveAnnotation(_document.GetBytesForOperation(), objectNumber, deltaX, deltaY, _document.ReadOptions);
    /// <summary>Moves selected annotations, including their standard group members, in one mutation.</summary>
    public PdfAnnotationEditResult MoveMany(IReadOnlyList<int> objectNumbers, double deltaX, double deltaY) => PdfAnnotationEditor.EditBatch(_document.GetBytesForOperation(), objectNumbers, PdfAnnotationEditor.BatchOperation.Move, _document.ReadOptions, deltaX, deltaY);
    /// <summary>Copies selected annotations and their standard groups on the same pages with a user-space offset.</summary>
    /// <remarks>Copies retain styling and appearances, receive unique names, and do not copy conversational replies.</remarks>
    public PdfAnnotationEditResult CopyMany(IReadOnlyList<int> objectNumbers, double deltaX = 10D, double deltaY = 10D) => PdfAnnotationEditor.EditBatch(_document.GetBytesForOperation(), objectNumbers, PdfAnnotationEditor.BatchOperation.Copy, _document.ReadOptions, deltaX, deltaY);
    /// <summary>Groups same-page editable markup annotations using standard /IRT and /RT Group relationships.</summary>
    /// <remarks>Existing groups are expanded. Conversational replies cannot be converted into group members.</remarks>
    public PdfAnnotationEditResult Group(IReadOnlyList<int> objectNumbers) => PdfAnnotationEditor.EditBatch(_document.GetBytesForOperation(), objectNumbers, PdfAnnotationEditor.BatchOperation.Group, _document.ReadOptions);
    /// <summary>Removes standard grouping relationships from selected groups without removing annotations or replies.</summary>
    public PdfAnnotationEditResult Ungroup(IReadOnlyList<int> objectNumbers) => PdfAnnotationEditor.EditBatch(_document.GetBytesForOperation(), objectNumbers, PdfAnnotationEditor.BatchOperation.Ungroup, _document.ReadOptions);
    /// <summary>Changes selected annotations' page painting order while preserving their relative order.</summary>
    public PdfAnnotationEditResult Arrange(IReadOnlyList<int> objectNumbers, PdfAnnotationOrderChange change) => PdfAnnotationEditor.EditBatch(_document.GetBytesForOperation(), objectNumbers, PdfAnnotationEditor.BatchOperation.Arrange, _document.ReadOptions, orderChange: change);
    /// <summary>Removes selected annotations and groups, retaining replies to removed annotations as standalone comments.</summary>
    /// <remarks>Append-only removal retains removed data in older revisions and requires explicit residual-data permission.</remarks>
    public PdfAnnotationEditResult RemoveMany(IReadOnlyList<int> objectNumbers, bool allowResidualDataInAppendOnly = false) => PdfAnnotationEditor.EditBatch(_document.GetBytesForOperation(), objectNumbers, PdfAnnotationEditor.BatchOperation.Remove, _document.ReadOptions, allowResidualDataInAppendOnly: allowResidualDataInAppendOnly);
    /// <summary>Resizes one annotation and proportionally transforms all supported subtype geometry.</summary>
    public PdfAnnotationEditResult Resize(int objectNumber, PdfPageRectangle rectangle) => PdfAnnotationEditor.ResizeAnnotation(_document.GetBytesForOperation(), objectNumber, rectangle, _document.ReadOptions);
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
