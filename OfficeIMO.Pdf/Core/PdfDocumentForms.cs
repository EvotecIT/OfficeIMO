namespace OfficeIMO.Pdf;

/// <summary>
/// Fluent simple AcroForm operations for a <see cref="PdfDocument"/>.
/// </summary>
public sealed class PdfDocumentForms {
    private readonly PdfDocument _document;

    internal PdfDocumentForms(PdfDocument document) {
        _document = document;
    }

    /// <summary>
    /// Creates a new PDF with simple form fields filled.
    /// </summary>
    public PdfDocument Fill(IReadOnlyDictionary<string, string> fieldValues) {
        return PdfDocument.FromBytes(PdfFormFiller.FillFields(_document.Snapshot(), fieldValues));
    }

    /// <summary>
    /// Creates a new PDF with simple form fields filled.
    /// </summary>
    public PdfDocument Fill(IReadOnlyDictionary<string, string> fieldValues, PdfFormFillerOptions formOptions) {
        return PdfDocument.FromBytes(PdfFormFiller.FillFields(_document.Snapshot(), fieldValues, formOptions));
    }

    /// <summary>
    /// Attempts to create a new PDF with simple form fields filled, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryFill(IReadOnlyDictionary<string, string> fieldValues, PdfReadOptions? options = null) {
        return _document.TryOperation("Fill form fields", PdfPreflightCapability.FillSimpleFormFields, () => Fill(fieldValues), options);
    }

    /// <summary>
    /// Attempts to create a new PDF with simple form fields filled, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryFill(IReadOnlyDictionary<string, string> fieldValues, PdfFormFillerOptions formOptions, PdfReadOptions? readOptions) {
        return _document.TryOperation("Fill form fields", PdfPreflightCapability.FillSimpleFormFields, () => Fill(fieldValues, formOptions), readOptions);
    }

    /// <summary>
    /// Creates a new PDF with simple form fields filled, including multi-value fields.
    /// </summary>
    public PdfDocument Fill(IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        return PdfDocument.FromBytes(PdfFormFiller.FillFields(_document.Snapshot(), fieldValues));
    }

    /// <summary>
    /// Creates a new PDF with simple form fields filled, including multi-value fields.
    /// </summary>
    public PdfDocument Fill(IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfFormFillerOptions formOptions) {
        return PdfDocument.FromBytes(PdfFormFiller.FillFields(_document.Snapshot(), fieldValues, formOptions));
    }

    /// <summary>
    /// Attempts to create a new PDF with simple form fields filled, including multi-value fields, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryFill(IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfReadOptions? options = null) {
        return _document.TryOperation("Fill form fields", PdfPreflightCapability.FillSimpleFormFields, () => Fill(fieldValues), options);
    }

    /// <summary>
    /// Attempts to create a new PDF with simple form fields filled, including multi-value fields, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryFill(IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfFormFillerOptions formOptions, PdfReadOptions? readOptions) {
        return _document.TryOperation("Fill form fields", PdfPreflightCapability.FillSimpleFormFields, () => Fill(fieldValues, formOptions), readOptions);
    }

    /// <summary>
    /// Appends a simple AcroForm field-value revision without rewriting the existing PDF bytes.
    /// </summary>
    public PdfDocument AppendRevision(IReadOnlyDictionary<string, string> fieldValues, bool keepNeedAppearances = true) {
        return PdfDocument.FromBytes(PdfIncrementalUpdater.UpdateFormFields(_document.Snapshot(), fieldValues, keepNeedAppearances));
    }

    /// <summary>
    /// Appends a simple AcroForm field-value revision without rewriting the existing PDF bytes.
    /// </summary>
    public PdfDocument AppendRevision(IReadOnlyDictionary<string, string> fieldValues, PdfIncrementalFormFieldUpdateOptions? formOptions) {
        return PdfDocument.FromBytes(PdfIncrementalUpdater.UpdateFormFields(_document.Snapshot(), fieldValues, formOptions));
    }

    /// <summary>
    /// Attempts to append a simple AcroForm field-value revision, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryAppendRevision(IReadOnlyDictionary<string, string> fieldValues, bool keepNeedAppearances = true, PdfReadOptions? options = null) {
        return _document.TryOperation("Append form field revision", PdfPreflightCapability.AppendFormFieldRevision, () => AppendRevision(fieldValues, keepNeedAppearances), options);
    }

    /// <summary>
    /// Attempts to append a simple AcroForm field-value revision, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryAppendRevision(IReadOnlyDictionary<string, string> fieldValues, PdfIncrementalFormFieldUpdateOptions? formOptions, PdfReadOptions? readOptions) {
        return _document.TryOperation("Append form field revision", PdfPreflightCapability.AppendFormFieldRevision, () => AppendRevision(fieldValues, formOptions), readOptions);
    }

    /// <summary>
    /// Appends a simple AcroForm field-value revision, including multi-value fields, without rewriting the existing PDF bytes.
    /// </summary>
    public PdfDocument AppendRevision(IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, bool keepNeedAppearances = true) {
        return PdfDocument.FromBytes(PdfIncrementalUpdater.UpdateFormFields(_document.Snapshot(), fieldValues, keepNeedAppearances));
    }

    /// <summary>
    /// Appends a simple AcroForm field-value revision, including multi-value fields, without rewriting the existing PDF bytes.
    /// </summary>
    public PdfDocument AppendRevision(IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfIncrementalFormFieldUpdateOptions? formOptions) {
        return PdfDocument.FromBytes(PdfIncrementalUpdater.UpdateFormFields(_document.Snapshot(), fieldValues, formOptions));
    }

    /// <summary>
    /// Attempts to append a simple AcroForm field-value revision, including multi-value fields, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryAppendRevision(IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, bool keepNeedAppearances = true, PdfReadOptions? options = null) {
        return _document.TryOperation("Append form field revision", PdfPreflightCapability.AppendFormFieldRevision, () => AppendRevision(fieldValues, keepNeedAppearances), options);
    }

    /// <summary>
    /// Attempts to append a simple AcroForm field-value revision, including multi-value fields, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryAppendRevision(IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfIncrementalFormFieldUpdateOptions? formOptions, PdfReadOptions? readOptions) {
        return _document.TryOperation("Append form field revision", PdfPreflightCapability.AppendFormFieldRevision, () => AppendRevision(fieldValues, formOptions), readOptions);
    }

    /// <summary>
    /// Creates a new PDF with simple form fields flattened.
    /// </summary>
    public PdfDocument Flatten() {
        return PdfDocument.FromBytes(PdfFormFiller.FlattenFields(_document.Snapshot()));
    }

    /// <summary>
    /// Creates a new PDF with simple form fields flattened.
    /// </summary>
    public PdfDocument Flatten(PdfFormFillerOptions formOptions) {
        return PdfDocument.FromBytes(PdfFormFiller.FlattenFields(_document.Snapshot(), formOptions));
    }

    /// <summary>
    /// Attempts to create a new PDF with simple form fields flattened, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryFlatten(PdfReadOptions? options = null) {
        return _document.TryOperation("Flatten form fields", PdfPreflightCapability.FlattenSimpleFormFields, Flatten, options);
    }

    /// <summary>
    /// Attempts to create a new PDF with simple form fields flattened, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryFlatten(PdfFormFillerOptions formOptions, PdfReadOptions? readOptions) {
        return _document.TryOperation("Flatten form fields", PdfPreflightCapability.FlattenSimpleFormFields, () => Flatten(formOptions), readOptions);
    }

    /// <summary>
    /// Creates a new PDF with simple form fields filled and flattened.
    /// </summary>
    public PdfDocument FillAndFlatten(IReadOnlyDictionary<string, string> fieldValues) {
        return PdfDocument.FromBytes(PdfFormFiller.FillAndFlattenFields(_document.Snapshot(), fieldValues));
    }

    /// <summary>
    /// Creates a new PDF with simple form fields filled and flattened.
    /// </summary>
    public PdfDocument FillAndFlatten(IReadOnlyDictionary<string, string> fieldValues, PdfFormFillerOptions formOptions) {
        return PdfDocument.FromBytes(PdfFormFiller.FillAndFlattenFields(_document.Snapshot(), fieldValues, formOptions));
    }

    /// <summary>
    /// Attempts to create a new PDF with simple form fields filled and flattened, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryFillAndFlatten(IReadOnlyDictionary<string, string> fieldValues, PdfReadOptions? options = null) {
        return _document.TryOperation("Fill and flatten form fields", PdfPreflightCapability.FillAndFlattenSimpleFormFields, () => FillAndFlatten(fieldValues), options);
    }

    /// <summary>
    /// Attempts to create a new PDF with simple form fields filled and flattened, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryFillAndFlatten(IReadOnlyDictionary<string, string> fieldValues, PdfFormFillerOptions formOptions, PdfReadOptions? readOptions) {
        return _document.TryOperation("Fill and flatten form fields", PdfPreflightCapability.FillAndFlattenSimpleFormFields, () => FillAndFlatten(fieldValues, formOptions), readOptions);
    }

    /// <summary>
    /// Creates a new PDF with simple form fields filled and flattened, including multi-value fields.
    /// </summary>
    public PdfDocument FillAndFlatten(IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        return PdfDocument.FromBytes(PdfFormFiller.FillAndFlattenFields(_document.Snapshot(), fieldValues));
    }

    /// <summary>
    /// Creates a new PDF with simple form fields filled and flattened, including multi-value fields.
    /// </summary>
    public PdfDocument FillAndFlatten(IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfFormFillerOptions formOptions) {
        return PdfDocument.FromBytes(PdfFormFiller.FillAndFlattenFields(_document.Snapshot(), fieldValues, formOptions));
    }

    /// <summary>
    /// Attempts to create a new PDF with simple form fields filled and flattened, including multi-value fields, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryFillAndFlatten(IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfReadOptions? options = null) {
        return _document.TryOperation("Fill and flatten form fields", PdfPreflightCapability.FillAndFlattenSimpleFormFields, () => FillAndFlatten(fieldValues), options);
    }

    /// <summary>
    /// Attempts to create a new PDF with simple form fields filled and flattened, including multi-value fields, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryFillAndFlatten(IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfFormFillerOptions formOptions, PdfReadOptions? readOptions) {
        return _document.TryOperation("Fill and flatten form fields", PdfPreflightCapability.FillAndFlattenSimpleFormFields, () => FillAndFlatten(fieldValues, formOptions), readOptions);
    }
}
