using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Existing-document annotation creation, update, removal, and flattening operations.</summary>
public sealed class PdfDocumentAnnotations {
    private readonly PdfDocument _document;
    internal PdfDocumentAnnotations(PdfDocument document) { _document = document; }
    /// <summary>Reads annotation metadata needed to select and edit annotations under the document's annotation mutation policy.</summary>
    /// <remarks>Requires a permitted annotation mutation, including certification restrictions. Returns no page text, images, forms, or other logical document content; general content extraction remains subject to its own permissions.</remarks>
    public IReadOnlyList<PdfAnnotation> GetForEditing() => GetForEditing(CancellationToken.None);
    /// <summary>Reads annotation editing metadata with cooperative cancellation.</summary>
    public IReadOnlyList<PdfAnnotation> GetForEditing(CancellationToken cancellationToken) => PdfAnnotationEditor.GetEditingMetadata(_document.GetBytesForOperation(cancellationToken), _document.ReadOptions, cancellationToken);
    /// <summary>Builds annotation-only visual hit regions for an authorized annotation editing workflow.</summary>
    /// <remarks>Uses the same annotation mutation policy as <see cref="GetForEditing()"/>. Regions contain annotation identity, subtype, and bounds, without extracted text, link targets, images, or form fields.</remarks>
    public PdfPageInteractionMap GetEditingInteractions(int pageNumber) => GetEditingInteractions(pageNumber, CancellationToken.None);
    /// <summary>Builds annotation-only interaction regions with cooperative cancellation.</summary>
    public PdfPageInteractionMap GetEditingInteractions(int pageNumber, CancellationToken cancellationToken) {
        var annotations = GetForEditing(cancellationToken);
        var layouts = _document.GetPageLayouts(_document.ReadOptions, cancellationToken);
        if (pageNumber < 1 || pageNumber > layouts.Count) throw new ArgumentOutOfRangeException(nameof(pageNumber));
        var interactions = PdfPageInteractionMap.CreateAnnotationEditing(layouts[pageNumber - 1], annotations);
        cancellationToken.ThrowIfCancellationRequested();
        return interactions;
    }
    /// <summary>Adds an annotation to an existing page.</summary>
    public PdfAnnotationEditResult Add(PdfAnnotationCreateOptions options) {
        Guard.NotNull(options, nameof(options));
        return PdfAnnotationEditor.AddAnnotation(_document.GetBytesForOperation(options.CancellationToken), options, _document.ReadOptions);
    }
    /// <summary>Adds a text or image-backed Stamp annotation without changing page content.</summary>
    public PdfAnnotationEditResult AddStamp(PdfStampAnnotationOptions options) => PdfAnnotationEditor.AddStampAnnotation(_document.GetBytesForOperation(options?.CancellationToken ?? default), options, _document.ReadOptions);
    /// <summary>Updates one indirect annotation.</summary>
    public PdfAnnotationEditResult Update(int objectNumber, PdfAnnotationUpdateOptions options) {
        Guard.NotNull(options, nameof(options));
        return PdfAnnotationEditor.UpdateAnnotation(_document.GetBytesForOperation(options.CancellationToken), objectNumber, options, _document.ReadOptions);
    }
    /// <summary>Moves one annotation and all supported subtype geometry by a user-space offset.</summary>
    public PdfAnnotationEditResult Move(int objectNumber, double deltaX, double deltaY) => Move(objectNumber, deltaX, deltaY, CancellationToken.None);
    /// <summary>Performs this annotation edit with cooperative cancellation.</summary>
    public PdfAnnotationEditResult Move(int objectNumber, double deltaX, double deltaY, CancellationToken cancellationToken) => PdfAnnotationEditor.MoveAnnotation(_document.GetBytesForOperation(cancellationToken), objectNumber, deltaX, deltaY, _document.ReadOptions, cancellationToken: cancellationToken);
    /// <summary>Moves selected annotations, including their standard group members, in one mutation.</summary>
    public PdfAnnotationEditResult MoveMany(IReadOnlyList<int> objectNumbers, double deltaX, double deltaY) => MoveMany(objectNumbers, deltaX, deltaY, CancellationToken.None);
    /// <summary>Performs this annotation edit with cooperative cancellation.</summary>
    public PdfAnnotationEditResult MoveMany(IReadOnlyList<int> objectNumbers, double deltaX, double deltaY, CancellationToken cancellationToken) => PdfAnnotationEditor.EditBatch(_document.GetBytesForOperation(cancellationToken), objectNumbers, PdfAnnotationEditor.BatchOperation.Move, _document.ReadOptions, deltaX, deltaY, cancellationToken: cancellationToken);
    /// <summary>Moves selected annotations by a top-left visual offset, mapping each page's crop, rotation, and user-unit scale in one mutation.</summary>
    public PdfAnnotationEditResult MoveManyVisual(IReadOnlyList<int> objectNumbers, double deltaX, double deltaY) => MoveManyVisual(objectNumbers, deltaX, deltaY, CancellationToken.None);
    /// <summary>Performs this annotation edit with cooperative cancellation.</summary>
    public PdfAnnotationEditResult MoveManyVisual(IReadOnlyList<int> objectNumbers, double deltaX, double deltaY, CancellationToken cancellationToken) => PdfAnnotationEditor.EditBatch(_document.GetBytesForOperation(cancellationToken), objectNumbers, PdfAnnotationEditor.BatchOperation.Move, _document.ReadOptions, deltaX, deltaY, visualOffset: true, cancellationToken: cancellationToken);
    /// <summary>Copies selected annotations and their standard groups on the same pages with a user-space offset.</summary>
    /// <remarks>Copies retain styling and appearances, receive unique names, and do not copy conversational replies.</remarks>
    public PdfAnnotationEditResult CopyMany(IReadOnlyList<int> objectNumbers, double deltaX = 10D, double deltaY = 10D) => CopyMany(objectNumbers, deltaX, deltaY, CancellationToken.None);
    /// <summary>Performs this annotation edit with cooperative cancellation.</summary>
    public PdfAnnotationEditResult CopyMany(IReadOnlyList<int> objectNumbers, double deltaX, double deltaY, CancellationToken cancellationToken) => PdfAnnotationEditor.EditBatch(_document.GetBytesForOperation(cancellationToken), objectNumbers, PdfAnnotationEditor.BatchOperation.Copy, _document.ReadOptions, deltaX, deltaY, cancellationToken: cancellationToken);
    /// <summary>Copies selected annotations on their existing pages by a top-left visual offset, mapping each page's crop, rotation, and user-unit scale in one mutation.</summary>
    /// <remarks>Copies retain styling and appearances, receive unique names, and do not copy conversational replies.</remarks>
    public PdfAnnotationEditResult CopyManyVisual(IReadOnlyList<int> objectNumbers, double deltaX = 10D, double deltaY = 10D) => CopyManyVisual(objectNumbers, deltaX, deltaY, CancellationToken.None);
    /// <summary>Performs this annotation edit with cooperative cancellation.</summary>
    public PdfAnnotationEditResult CopyManyVisual(IReadOnlyList<int> objectNumbers, double deltaX, double deltaY, CancellationToken cancellationToken) => PdfAnnotationEditor.EditBatch(_document.GetBytesForOperation(cancellationToken), objectNumbers, PdfAnnotationEditor.BatchOperation.Copy, _document.ReadOptions, deltaX, deltaY, visualOffset: true, cancellationToken: cancellationToken);
    /// <summary>Groups same-page editable markup annotations using standard /IRT and /RT Group relationships.</summary>
    /// <remarks>Existing groups are expanded. Conversational replies cannot be converted into group members.</remarks>
    public PdfAnnotationEditResult Group(IReadOnlyList<int> objectNumbers) => Group(objectNumbers, CancellationToken.None);
    /// <summary>Performs this annotation edit with cooperative cancellation.</summary>
    public PdfAnnotationEditResult Group(IReadOnlyList<int> objectNumbers, CancellationToken cancellationToken) => PdfAnnotationEditor.EditBatch(_document.GetBytesForOperation(cancellationToken), objectNumbers, PdfAnnotationEditor.BatchOperation.Group, _document.ReadOptions, cancellationToken: cancellationToken);
    /// <summary>Removes standard grouping relationships from selected groups without removing annotations or replies.</summary>
    public PdfAnnotationEditResult Ungroup(IReadOnlyList<int> objectNumbers) => Ungroup(objectNumbers, CancellationToken.None);
    /// <summary>Performs this annotation edit with cooperative cancellation.</summary>
    public PdfAnnotationEditResult Ungroup(IReadOnlyList<int> objectNumbers, CancellationToken cancellationToken) => PdfAnnotationEditor.EditBatch(_document.GetBytesForOperation(cancellationToken), objectNumbers, PdfAnnotationEditor.BatchOperation.Ungroup, _document.ReadOptions, cancellationToken: cancellationToken);
    /// <summary>Changes selected annotations' page painting order while preserving their relative order.</summary>
    public PdfAnnotationEditResult Arrange(IReadOnlyList<int> objectNumbers, PdfAnnotationOrderChange change) => Arrange(objectNumbers, change, CancellationToken.None);
    /// <summary>Performs this annotation edit with cooperative cancellation.</summary>
    public PdfAnnotationEditResult Arrange(IReadOnlyList<int> objectNumbers, PdfAnnotationOrderChange change, CancellationToken cancellationToken) => PdfAnnotationEditor.EditBatch(_document.GetBytesForOperation(cancellationToken), objectNumbers, PdfAnnotationEditor.BatchOperation.Arrange, _document.ReadOptions, orderChange: change, cancellationToken: cancellationToken);
    /// <summary>Removes selected annotations and groups, retaining replies to removed annotations as standalone comments.</summary>
    /// <remarks>Append-only removal retains removed data in older revisions and requires explicit residual-data permission.</remarks>
    public PdfAnnotationEditResult RemoveMany(IReadOnlyList<int> objectNumbers, bool allowResidualDataInAppendOnly = false) => RemoveMany(objectNumbers, allowResidualDataInAppendOnly, CancellationToken.None);
    /// <summary>Performs this annotation edit with cooperative cancellation.</summary>
    public PdfAnnotationEditResult RemoveMany(IReadOnlyList<int> objectNumbers, bool allowResidualDataInAppendOnly, CancellationToken cancellationToken) => PdfAnnotationEditor.EditBatch(_document.GetBytesForOperation(cancellationToken), objectNumbers, PdfAnnotationEditor.BatchOperation.Remove, _document.ReadOptions, allowResidualDataInAppendOnly: allowResidualDataInAppendOnly, cancellationToken: cancellationToken);
    /// <summary>Resizes one annotation and proportionally transforms all supported subtype geometry.</summary>
    public PdfAnnotationEditResult Resize(int objectNumber, PdfPageRectangle rectangle) => Resize(objectNumber, rectangle, CancellationToken.None);
    /// <summary>Performs this annotation edit with cooperative cancellation.</summary>
    public PdfAnnotationEditResult Resize(int objectNumber, PdfPageRectangle rectangle, CancellationToken cancellationToken) => PdfAnnotationEditor.ResizeAnnotation(_document.GetBytesForOperation(cancellationToken), objectNumber, rectangle, _document.ReadOptions, cancellationToken: cancellationToken);
    /// <summary>Adds a reply to one indirect annotation.</summary>
    public PdfAnnotationEditResult AddReply(int parentObjectNumber, string contents, PdfAnnotationReplyOptions? options = null) => PdfAnnotationReviewEditor.AddReply(_document.GetBytesForOperation(options?.CancellationToken ?? default), parentObjectNumber, contents, options, _document.ReadOptions);
    /// <summary>Sets the standard review state on one indirect annotation.</summary>
    public PdfAnnotationEditResult SetReviewState(int objectNumber, PdfAnnotationReviewState state, PdfMutationExecutionPreference executionPreference = PdfMutationExecutionPreference.Automatic, bool allowResidualDataInAppendOnly = false) => SetReviewState(objectNumber, state, executionPreference, allowResidualDataInAppendOnly, CancellationToken.None);
    /// <summary>Sets an annotation review state with cooperative cancellation.</summary>
    public PdfAnnotationEditResult SetReviewState(int objectNumber, PdfAnnotationReviewState state, PdfMutationExecutionPreference executionPreference, bool allowResidualDataInAppendOnly, CancellationToken cancellationToken) => PdfAnnotationReviewEditor.SetState(_document.GetBytesForOperation(cancellationToken), objectNumber, state, executionPreference, allowResidualDataInAppendOnly, _document.ReadOptions, cancellationToken);
    /// <summary>Reads annotation reply threads and review states.</summary>
    public PdfAnnotationReviewCatalog GetReviewCatalog() => GetReviewCatalog(CancellationToken.None);
    /// <summary>Reads annotation reply threads and states with cooperative cancellation.</summary>
    public PdfAnnotationReviewCatalog GetReviewCatalog(CancellationToken cancellationToken) => PdfAnnotationReviewCatalog.Read(_document.GetBytesForOperation(cancellationToken), _document.ReadOptions, new PdfAnnotationReviewCatalogOptions { CancellationToken = cancellationToken });
    /// <summary>Removes matching annotations.</summary>
    public PdfAnnotationEditResult Remove(PdfAnnotationRemovalOptions? options = null) => PdfAnnotationEditor.RemoveAnnotations(_document.GetBytesForOperation(options?.CancellationToken ?? default), options, _document.ReadOptions);
    /// <summary>Flattens selected supported visual annotations.</summary>
    public PdfAnnotationEditResult Flatten(PdfAnnotationFlattenOptions? options = null) => PdfAnnotationEditor.FlattenAnnotations(_document.GetBytesForOperation(options?.CancellationToken ?? default), options, _document.ReadOptions);
}
