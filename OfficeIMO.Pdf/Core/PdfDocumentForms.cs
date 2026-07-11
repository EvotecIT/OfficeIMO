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
        Guard.NotNull(fieldValues, nameof(fieldValues));
        return _document.TryMutationOperation(
            "Fill form fields",
            PdfPreflightCapability.FillSimpleFormFields,
            PdfMutationOperation.FillFormFields,
            mode => mode == PdfMutationExecutionMode.AppendOnly
                ? AppendRevision(fieldValues, CreateIncrementalOptions(formOptions: null))
                : Fill(fieldValues),
            fieldValues.Keys,
            options);
    }

    /// <summary>
    /// Attempts to create a new PDF with simple form fields filled, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryFill(IReadOnlyDictionary<string, string> fieldValues, PdfFormFillerOptions formOptions, PdfReadOptions? readOptions) {
        Guard.NotNull(fieldValues, nameof(fieldValues));
        Guard.NotNull(formOptions, nameof(formOptions));
        return _document.TryMutationOperation(
            "Fill form fields",
            PdfPreflightCapability.FillSimpleFormFields,
            PdfMutationOperation.FillFormFields,
            mode => mode == PdfMutationExecutionMode.AppendOnly
                ? AppendRevision(fieldValues, CreateIncrementalOptions(formOptions))
                : Fill(fieldValues, formOptions),
            fieldValues.Keys,
            readOptions);
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
        Guard.NotNull(fieldValues, nameof(fieldValues));
        return _document.TryMutationOperation(
            "Fill form fields",
            PdfPreflightCapability.FillSimpleFormFields,
            PdfMutationOperation.FillFormFields,
            mode => mode == PdfMutationExecutionMode.AppendOnly
                ? AppendRevision(fieldValues, CreateIncrementalOptions(formOptions: null))
                : Fill(fieldValues),
            fieldValues.Keys,
            options);
    }

    /// <summary>
    /// Attempts to create a new PDF with simple form fields filled, including multi-value fields, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryFill(IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfFormFillerOptions formOptions, PdfReadOptions? readOptions) {
        Guard.NotNull(fieldValues, nameof(fieldValues));
        Guard.NotNull(formOptions, nameof(formOptions));
        return _document.TryMutationOperation(
            "Fill form fields",
            PdfPreflightCapability.FillSimpleFormFields,
            PdfMutationOperation.FillFormFields,
            mode => mode == PdfMutationExecutionMode.AppendOnly
                ? AppendRevision(fieldValues, CreateIncrementalOptions(formOptions))
                : Fill(fieldValues, formOptions),
            fieldValues.Keys,
            readOptions);
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
        Guard.NotNull(fieldValues, nameof(fieldValues));
        return _document.TryMutationOperation(
            "Append form field revision",
            PdfPreflightCapability.AppendFormFieldRevision,
            PdfMutationOperation.FillFormFields,
            _ => AppendRevision(fieldValues, keepNeedAppearances),
            fieldValues.Keys,
            options,
            PdfMutationExecutionPreference.RequireAppendOnly);
    }

    /// <summary>
    /// Attempts to append a simple AcroForm field-value revision, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryAppendRevision(IReadOnlyDictionary<string, string> fieldValues, PdfIncrementalFormFieldUpdateOptions? formOptions, PdfReadOptions? readOptions) {
        Guard.NotNull(fieldValues, nameof(fieldValues));
        return _document.TryMutationOperation(
            "Append form field revision",
            PdfPreflightCapability.AppendFormFieldRevision,
            PdfMutationOperation.FillFormFields,
            _ => AppendRevision(fieldValues, formOptions),
            fieldValues.Keys,
            readOptions,
            PdfMutationExecutionPreference.RequireAppendOnly);
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
        Guard.NotNull(fieldValues, nameof(fieldValues));
        return _document.TryMutationOperation(
            "Append form field revision",
            PdfPreflightCapability.AppendFormFieldRevision,
            PdfMutationOperation.FillFormFields,
            _ => AppendRevision(fieldValues, keepNeedAppearances),
            fieldValues.Keys,
            options,
            PdfMutationExecutionPreference.RequireAppendOnly);
    }

    /// <summary>
    /// Attempts to append a simple AcroForm field-value revision, including multi-value fields, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryAppendRevision(IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfIncrementalFormFieldUpdateOptions? formOptions, PdfReadOptions? readOptions) {
        Guard.NotNull(fieldValues, nameof(fieldValues));
        return _document.TryMutationOperation(
            "Append form field revision",
            PdfPreflightCapability.AppendFormFieldRevision,
            PdfMutationOperation.FillFormFields,
            _ => AppendRevision(fieldValues, formOptions),
            fieldValues.Keys,
            readOptions,
            PdfMutationExecutionPreference.RequireAppendOnly);
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
        return _document.TryMutationOperation("Flatten form fields", PdfPreflightCapability.FlattenSimpleFormFields, PdfMutationOperation.FlattenFormFields, Flatten, options);
    }

    /// <summary>
    /// Attempts to create a new PDF with simple form fields flattened, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryFlatten(PdfFormFillerOptions formOptions, PdfReadOptions? readOptions) {
        return _document.TryMutationOperation("Flatten form fields", PdfPreflightCapability.FlattenSimpleFormFields, PdfMutationOperation.FlattenFormFields, () => Flatten(formOptions), readOptions);
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
        return _document.TryMutationOperation("Fill and flatten form fields", PdfPreflightCapability.FillAndFlattenSimpleFormFields, PdfMutationOperation.FillAndFlattenFormFields, () => FillAndFlatten(fieldValues), options);
    }

    /// <summary>
    /// Attempts to create a new PDF with simple form fields filled and flattened, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryFillAndFlatten(IReadOnlyDictionary<string, string> fieldValues, PdfFormFillerOptions formOptions, PdfReadOptions? readOptions) {
        return _document.TryMutationOperation("Fill and flatten form fields", PdfPreflightCapability.FillAndFlattenSimpleFormFields, PdfMutationOperation.FillAndFlattenFormFields, () => FillAndFlatten(fieldValues, formOptions), readOptions);
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
        return _document.TryMutationOperation("Fill and flatten form fields", PdfPreflightCapability.FillAndFlattenSimpleFormFields, PdfMutationOperation.FillAndFlattenFormFields, () => FillAndFlatten(fieldValues), options);
    }

    /// <summary>
    /// Attempts to create a new PDF with simple form fields filled and flattened, including multi-value fields, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryFillAndFlatten(IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfFormFillerOptions formOptions, PdfReadOptions? readOptions) {
        return _document.TryMutationOperation("Fill and flatten form fields", PdfPreflightCapability.FillAndFlattenSimpleFormFields, PdfMutationOperation.FillAndFlattenFormFields, () => FillAndFlatten(fieldValues, formOptions), readOptions);
    }

    private static PdfIncrementalFormFieldUpdateOptions CreateIncrementalOptions(PdfFormFillerOptions? formOptions) {
        if (formOptions?.HasAppearanceFontFamily == true || formOptions?.HasAppearanceFontFallbacks == true) {
            throw new NotSupportedException("Append-only form updates cannot yet embed custom appearance fonts. Use the default appearance policy or a PDF that permits full rewrite.");
        }

        return new PdfIncrementalFormFieldUpdateOptions {
            KeepNeedAppearances = formOptions?.KeepNeedAppearances ?? false,
            GenerateAppearanceStreams = true
        };
    }
}
