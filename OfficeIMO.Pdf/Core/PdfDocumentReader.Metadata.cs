namespace OfficeIMO.Pdf;

/// <summary>
/// Fluent document metadata readback operations for a <see cref="PdfDocument"/>.
/// </summary>
public sealed partial class PdfDocumentReader {
    /// <summary>
    /// Reads catalog XMP metadata when a readable metadata stream is present.
    /// </summary>
    public PdfXmpMetadataInfo? XmpMetadata(PdfReadOptions? readOptions = null) {
        return DocumentInfo(readOptions).XmpMetadata;
    }

    /// <summary>
    /// Attempts to read catalog XMP metadata, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfXmpMetadataInfo> TryXmpMetadata(PdfReadOptions? options = null) {
        return _document.TryOperation("Read XMP metadata", PdfPreflightCapability.ReadLogicalObjects, () => XmpMetadata(options) ?? throw new InvalidOperationException("No readable XMP metadata was found."), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads catalog output-intent metadata.
    /// </summary>
    public IReadOnlyList<PdfOutputIntentInfo> OutputIntents(PdfReadOptions? readOptions = null) {
        return DocumentInfo(readOptions).OutputIntents;
    }

    /// <summary>
    /// Attempts to read catalog output-intent metadata, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfOutputIntentInfo>> TryOutputIntents(PdfReadOptions? options = null) {
        return _document.TryOperation("Read output intents", PdfPreflightCapability.ReadLogicalObjects, () => OutputIntents(options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads output intents with a matching /S subtype.
    /// </summary>
    public IReadOnlyList<PdfOutputIntentInfo> OutputIntentsBySubtype(string subtype, PdfReadOptions? readOptions = null) {
        return DocumentInfo(readOptions).GetOutputIntentsBySubtype(subtype);
    }

    /// <summary>
    /// Attempts to read output intents with a matching /S subtype, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfOutputIntentInfo>> TryOutputIntentsBySubtype(string subtype, PdfReadOptions? options = null) {
        return _document.TryOperation("Read output intents", PdfPreflightCapability.ReadLogicalObjects, () => OutputIntentsBySubtype(subtype, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads output intents with a matching /OutputConditionIdentifier.
    /// </summary>
    public IReadOnlyList<PdfOutputIntentInfo> OutputIntentsByOutputConditionIdentifier(string outputConditionIdentifier, PdfReadOptions? readOptions = null) {
        return DocumentInfo(readOptions).GetOutputIntentsByOutputConditionIdentifier(outputConditionIdentifier);
    }

    /// <summary>
    /// Attempts to read output intents with a matching /OutputConditionIdentifier, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfOutputIntentInfo>> TryOutputIntentsByOutputConditionIdentifier(string outputConditionIdentifier, PdfReadOptions? options = null) {
        return _document.TryOperation("Read output intents", PdfPreflightCapability.ReadLogicalObjects, () => OutputIntentsByOutputConditionIdentifier(outputConditionIdentifier, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads tagged PDF structure metadata when it is present.
    /// </summary>
    public PdfTaggedContentInfo? TaggedContent(PdfReadOptions? readOptions = null) {
        return DocumentInfo(readOptions).TaggedContent;
    }

    /// <summary>
    /// Attempts to read tagged PDF structure metadata, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfTaggedContentInfo> TryTaggedContent(PdfReadOptions? options = null) {
        return _document.TryOperation("Read tagged content", PdfPreflightCapability.ReadLogicalObjects, () => TaggedContent(options) ?? throw new InvalidOperationException("No readable tagged content metadata was found."), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads optional-content/layer metadata when it is present.
    /// </summary>
    public PdfOptionalContentProperties? OptionalContent(PdfReadOptions? readOptions = null) {
        return DocumentInfo(readOptions).OptionalContent;
    }

    /// <summary>
    /// Attempts to read optional-content/layer metadata, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfOptionalContentProperties> TryOptionalContent(PdfReadOptions? options = null) {
        return _document.TryOperation("Read optional content", PdfPreflightCapability.ReadLogicalObjects, () => OptionalContent(options) ?? throw new InvalidOperationException("No readable optional-content metadata was found."), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads optional-content groups discovered from catalog /OCProperties.
    /// </summary>
    public IReadOnlyList<PdfOptionalContentGroup> OptionalContentGroups(PdfReadOptions? readOptions = null) {
        return DocumentInfo(readOptions).OptionalContentGroups;
    }

    /// <summary>
    /// Attempts to read optional-content groups, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfOptionalContentGroup>> TryOptionalContentGroups(PdfReadOptions? options = null) {
        return _document.TryOperation("Read optional content groups", PdfPreflightCapability.ReadLogicalObjects, () => OptionalContentGroups(options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads optional-content groups with a matching layer display name.
    /// </summary>
    public IReadOnlyList<PdfOptionalContentGroup> OptionalContentGroupsByName(string name, PdfReadOptions? readOptions = null) {
        return DocumentInfo(readOptions).GetOptionalContentGroupsByName(name);
    }

    /// <summary>
    /// Attempts to read optional-content groups with a matching layer display name, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfOptionalContentGroup>> TryOptionalContentGroupsByName(string name, PdfReadOptions? options = null) {
        return _document.TryOperation("Read optional content groups", PdfPreflightCapability.ReadLogicalObjects, () => OptionalContentGroupsByName(name, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads attachment metadata without extracting attachment payload bytes.
    /// </summary>
    public IReadOnlyList<PdfAttachmentInfo> AttachmentMetadata(PdfReadOptions? readOptions = null) {
        return DocumentInfo(readOptions).Attachments;
    }

    /// <summary>
    /// Attempts to read attachment metadata without extracting attachment payload bytes, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfAttachmentInfo>> TryAttachmentMetadata(PdfReadOptions? options = null) {
        return _document.TryOperation("Read attachment metadata", PdfPreflightCapability.ReadLogicalObjects, () => AttachmentMetadata(options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads attachment metadata with a matching name-tree key or associated-file fallback name.
    /// </summary>
    public IReadOnlyList<PdfAttachmentInfo> AttachmentMetadataByName(string name, PdfReadOptions? readOptions = null) {
        return DocumentInfo(readOptions).GetAttachmentsByName(name);
    }

    /// <summary>
    /// Attempts to read attachment metadata with a matching name-tree key or associated-file fallback name.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfAttachmentInfo>> TryAttachmentMetadataByName(string name, PdfReadOptions? options = null) {
        return _document.TryOperation("Read attachment metadata", PdfPreflightCapability.ReadLogicalObjects, () => AttachmentMetadataByName(name, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads attachment metadata with a matching file specification file name.
    /// </summary>
    public IReadOnlyList<PdfAttachmentInfo> AttachmentMetadataByFileName(string fileName, PdfReadOptions? readOptions = null) {
        return DocumentInfo(readOptions).GetAttachmentsByFileName(fileName);
    }

    /// <summary>
    /// Attempts to read attachment metadata with a matching file specification file name.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfAttachmentInfo>> TryAttachmentMetadataByFileName(string fileName, PdfReadOptions? options = null) {
        return _document.TryOperation("Read attachment metadata", PdfPreflightCapability.ReadLogicalObjects, () => AttachmentMetadataByFileName(fileName, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads attachment metadata from a matching catalog source.
    /// </summary>
    public IReadOnlyList<PdfAttachmentInfo> AttachmentMetadataBySource(string source, PdfReadOptions? readOptions = null) {
        return DocumentInfo(readOptions).GetAttachmentsBySource(source);
    }

    /// <summary>
    /// Attempts to read attachment metadata from a matching catalog source.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfAttachmentInfo>> TryAttachmentMetadataBySource(string source, PdfReadOptions? options = null) {
        return _document.TryOperation("Read attachment metadata", PdfPreflightCapability.ReadLogicalObjects, () => AttachmentMetadataBySource(source, options), ResolveReadOptions(options));
    }

    /// <summary>
    /// Reads attachment metadata with a matching associated-file relationship.
    /// </summary>
    public IReadOnlyList<PdfAttachmentInfo> AttachmentMetadataByRelationship(PdfAssociatedFileRelationship relationship, PdfReadOptions? readOptions = null) {
        return DocumentInfo(readOptions).GetAttachmentsByRelationship(relationship);
    }

    /// <summary>
    /// Attempts to read attachment metadata with a matching associated-file relationship.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfAttachmentInfo>> TryAttachmentMetadataByRelationship(PdfAssociatedFileRelationship relationship, PdfReadOptions? options = null) {
        return _document.TryOperation("Read attachment metadata", PdfPreflightCapability.ReadLogicalObjects, () => AttachmentMetadataByRelationship(relationship, options), ResolveReadOptions(options));
    }
}
