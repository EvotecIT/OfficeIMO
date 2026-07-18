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
        return _document.ApplyMutation(input => PdfFormFiller.FillFields(input, fieldValues));
    }

    /// <summary>
    /// Creates a new PDF with simple form fields filled.
    /// </summary>
    public PdfDocument Fill(IReadOnlyDictionary<string, string> fieldValues, PdfFormFillerOptions formOptions) {
        return _document.ApplyMutation(input => PdfFormFiller.FillFields(input, fieldValues, formOptions));
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
                ? AppendRevisionWithReadOptions(fieldValues, CreateIncrementalOptions(formOptions: null), options ?? _document.ReadOptions)
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
                ? AppendRevisionWithReadOptions(fieldValues, CreateIncrementalOptions(formOptions), readOptions ?? _document.ReadOptions)
                : Fill(fieldValues, formOptions),
            fieldValues.Keys,
            readOptions);
    }

    /// <summary>
    /// Creates a new PDF with simple form fields filled, including multi-value fields.
    /// </summary>
    public PdfDocument Fill(IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        return _document.ApplyMutation(input => PdfFormFiller.FillFields(input, fieldValues));
    }

    /// <summary>
    /// Creates a new PDF with simple form fields filled, including multi-value fields.
    /// </summary>
    public PdfDocument Fill(IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfFormFillerOptions formOptions) {
        return _document.ApplyMutation(input => PdfFormFiller.FillFields(input, fieldValues, formOptions));
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
                ? AppendRevisionWithReadOptions(fieldValues, CreateIncrementalOptions(formOptions: null), options ?? _document.ReadOptions)
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
                ? AppendRevisionWithReadOptions(fieldValues, CreateIncrementalOptions(formOptions), readOptions ?? _document.ReadOptions)
                : Fill(fieldValues, formOptions),
            fieldValues.Keys,
            readOptions);
    }

    /// <summary>
    /// Appends a simple AcroForm field-value revision without rewriting the existing PDF bytes.
    /// </summary>
    public PdfDocument AppendRevision(IReadOnlyDictionary<string, string> fieldValues, bool keepNeedAppearances = true) {
        return _document.ApplyMutation(input => PdfIncrementalUpdater.UpdateFormFields(input, fieldValues, keepNeedAppearances));
    }

    /// <summary>
    /// Appends a simple AcroForm field-value revision without rewriting the existing PDF bytes.
    /// </summary>
    public PdfDocument AppendRevision(IReadOnlyDictionary<string, string> fieldValues, PdfIncrementalFormFieldUpdateOptions? formOptions) {
        return _document.ApplyMutation(input => PdfIncrementalUpdater.UpdateFormFields(input, fieldValues, formOptions));
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
            _ => AppendRevisionWithReadOptions(fieldValues, CreateIncrementalOptions(keepNeedAppearances), options ?? _document.ReadOptions),
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
            _ => AppendRevisionWithReadOptions(fieldValues, formOptions, readOptions ?? _document.ReadOptions),
            fieldValues.Keys,
            readOptions,
            PdfMutationExecutionPreference.RequireAppendOnly);
    }

    /// <summary>
    /// Appends a simple AcroForm field-value revision, including multi-value fields, without rewriting the existing PDF bytes.
    /// </summary>
    public PdfDocument AppendRevision(IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, bool keepNeedAppearances = true) {
        return _document.ApplyMutation(input => PdfIncrementalUpdater.UpdateFormFields(input, fieldValues, keepNeedAppearances));
    }

    /// <summary>
    /// Appends a simple AcroForm field-value revision, including multi-value fields, without rewriting the existing PDF bytes.
    /// </summary>
    public PdfDocument AppendRevision(IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfIncrementalFormFieldUpdateOptions? formOptions) {
        return _document.ApplyMutation(input => PdfIncrementalUpdater.UpdateFormFields(input, fieldValues, formOptions));
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
            _ => AppendRevisionWithReadOptions(fieldValues, CreateIncrementalOptions(keepNeedAppearances), options ?? _document.ReadOptions),
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
            _ => AppendRevisionWithReadOptions(fieldValues, formOptions, readOptions ?? _document.ReadOptions),
            fieldValues.Keys,
            readOptions,
            PdfMutationExecutionPreference.RequireAppendOnly);
    }

    /// <summary>
    /// Creates a new PDF with simple form fields flattened.
    /// </summary>
    public PdfDocument Flatten() {
        return _document.ApplyMutation(input => PdfFormFiller.FlattenFields(input));
    }

    /// <summary>
    /// Creates a new PDF with simple form fields flattened.
    /// </summary>
    public PdfDocument Flatten(PdfFormFillerOptions formOptions) {
        return _document.ApplyMutation(input => PdfFormFiller.FlattenFields(input, formOptions));
    }

    /// <summary>Creates a new PDF with only the named simple form fields flattened.</summary>
    public PdfDocument Flatten(params string[] fieldNames) {
        return _document.ApplyMutation(input => PdfFormFiller.FlattenFields(input, fieldNames));
    }

    /// <summary>Creates a new PDF with only the named simple form fields flattened.</summary>
    public PdfDocument Flatten(IReadOnlyCollection<string> fieldNames, PdfFormFillerOptions formOptions) {
        return _document.ApplyMutation(input => PdfFormFiller.FlattenFields(input, fieldNames, formOptions));
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
        return _document.ApplyMutation(input => PdfFormFiller.FillAndFlattenFields(input, fieldValues));
    }

    /// <summary>
    /// Creates a new PDF with simple form fields filled and flattened.
    /// </summary>
    public PdfDocument FillAndFlatten(IReadOnlyDictionary<string, string> fieldValues, PdfFormFillerOptions formOptions) {
        return _document.ApplyMutation(input => PdfFormFiller.FillAndFlattenFields(input, fieldValues, formOptions));
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
        return _document.ApplyMutation(input => PdfFormFiller.FillAndFlattenFields(input, fieldValues));
    }

    /// <summary>
    /// Creates a new PDF with simple form fields filled and flattened, including multi-value fields.
    /// </summary>
    public PdfDocument FillAndFlatten(IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfFormFillerOptions formOptions) {
        return _document.ApplyMutation(input => PdfFormFiller.FillAndFlattenFields(input, fieldValues, formOptions));
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

    /// <summary>Exports readable field values as a typed data set.</summary>
    public PdfFormDataSet ExportData() => PdfFormData.Export(_document.GetBytesForOperation(), _document.ReadOptions);

    /// <summary>Exports readable field values as XFDF.</summary>
    public string ExportXfdf() => ExportData().ToXfdf();

    /// <summary>Imports a typed data set through the validated form filler.</summary>
    public PdfDocument ImportData(PdfFormDataSet data, PdfFormFillerOptions? options = null) => _document.ApplyMutation(input => PdfFormData.Import(input, data, options));

    /// <summary>Imports XFDF through the validated form filler.</summary>
    public PdfDocument ImportXfdf(string xfdf, PdfFormFillerOptions? options = null) => _document.ApplyMutation(input => PdfFormData.ImportXfdf(input, xfdf, options));

    /// <summary>Transactionally creates or edits fields, widgets, ordering, and selective flattening.</summary>
    public PdfAcroFormEditResult Edit(Action<PdfAcroFormEditSession> edit) => PdfAcroFormEditor.Edit(_document.GetBytesForOperation(), edit, _document.ReadOptions);

    private static PdfIncrementalFormFieldUpdateOptions CreateIncrementalOptions(PdfFormFillerOptions? formOptions) {
        if (formOptions?.HasAppearanceFontFamily == true || formOptions?.HasAppearanceFontFallbacks == true) {
            throw new NotSupportedException("Append-only form updates cannot yet embed custom appearance fonts. Use the default appearance policy or a PDF that permits full rewrite.");
        }

        return new PdfIncrementalFormFieldUpdateOptions {
            KeepNeedAppearances = formOptions?.KeepNeedAppearances ?? false,
            GenerateAppearanceStreams = true
        };
    }

    private static PdfIncrementalFormFieldUpdateOptions CreateIncrementalOptions(bool keepNeedAppearances) => new PdfIncrementalFormFieldUpdateOptions {
        KeepNeedAppearances = keepNeedAppearances,
        GenerateAppearanceStreams = !keepNeedAppearances
    };

    private PdfDocument AppendRevisionWithReadOptions(
        IReadOnlyDictionary<string, string> fieldValues,
        PdfIncrementalFormFieldUpdateOptions? formOptions,
        PdfReadOptions? readOptions) => _document.ApplyMutation(
            input => PdfIncrementalUpdater.UpdateFormFields(input, fieldValues, formOptions, readOptions),
            operationName: "AppendRevision");

    private PdfDocument AppendRevisionWithReadOptions(
        IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues,
        PdfIncrementalFormFieldUpdateOptions? formOptions,
        PdfReadOptions? readOptions) => _document.ApplyMutation(
            input => PdfIncrementalUpdater.UpdateFormFields(input, fieldValues, formOptions, readOptions),
            operationName: "AppendRevision");
}
