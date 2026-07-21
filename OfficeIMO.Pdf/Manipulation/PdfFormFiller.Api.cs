using OfficeIMO.Drawing.Internal;
namespace OfficeIMO.Pdf;

internal static partial class PdfFormFiller {
    /// <summary>
    /// Returns a new PDF with simple AcroForm field values updated by fully qualified field name.
    /// </summary>
    public static byte[] FillFields(byte[] pdf, IReadOnlyDictionary<string, string> fieldValues) {
        return FillFields(pdf, ToFormFieldValues(fieldValues), options: null, readOptions: null);
    }

    /// <summary>
    /// Returns a new PDF with simple AcroForm field values updated by fully qualified field name.
    /// </summary>
    public static byte[] FillFields(byte[] pdf, IReadOnlyDictionary<string, string> fieldValues, PdfFormFillerOptions? options) {
        return FillFields(pdf, ToFormFieldValues(fieldValues), options, readOptions: null);
    }

    internal static byte[] FillFields(byte[] pdf, IReadOnlyDictionary<string, string> fieldValues, PdfFormFillerOptions? options, PdfReadOptions? readOptions) =>
        FillFields(pdf, ToFormFieldValues(fieldValues), options, readOptions);

    /// <summary>
    /// Returns a new PDF with simple AcroForm field values updated by fully qualified field name.
    /// </summary>
    public static byte[] FillFields(byte[] pdf, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        return FillFields(pdf, fieldValues, options: null, readOptions: null);
    }

    /// <summary>
    /// Returns a new PDF with simple AcroForm field values updated by fully qualified field name.
    /// </summary>
    public static byte[] FillFields(byte[] pdf, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfFormFillerOptions? options) {
        return FillFields(pdf, fieldValues, options, readOptions: null);
    }

    internal static byte[] FillFields(byte[] pdf, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfFormFillerOptions? options, PdfReadOptions? readOptions) =>
        FillFieldsCore(pdf, fieldValues, options, readOptions, requireMutationPlan: true);

    internal static byte[] FillFieldsWithinPlannedRewrite(byte[] pdf, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfFormFillerOptions? options = null) {
        return FillFieldsCore(pdf, fieldValues, options, readOptions: null, requireMutationPlan: false);
    }

    private static byte[] FillFieldsCore(byte[] pdf, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfFormFillerOptions? options, PdfReadOptions? readOptions, bool requireMutationPlan) {
        Guard.NotNull(pdf, nameof(pdf));
        ValidateFieldValues(fieldValues);
        if (requireMutationPlan) _ = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.FillFormFields, readOptions, fieldNames: fieldValues.Keys);

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, readOptions);
        int catalogObjectNumber = FindCatalogObjectNumber(objects, trailerRaw);
        if (catalogObjectNumber == 0 ||
            objects[catalogObjectNumber].Value is not PdfDictionary catalog ||
            !catalog.Items.TryGetValue("AcroForm", out var acroFormObject) ||
            ResolveDictionary(objects, acroFormObject) is not PdfDictionary acroForm ||
            !acroForm.Items.TryGetValue("Fields", out var fieldsObject) ||
            ResolveObject(objects, fieldsObject) is not PdfArray fields) {
            throw new ArgumentException("PDF does not contain a readable AcroForm field tree.", nameof(pdf));
        }

        var remaining = new HashSet<string>(fieldValues.Keys, StringComparer.Ordinal);
        int nextObjectNumber = objects.Keys.Count == 0 ? 1 : objects.Keys.Max() + 1;
        int? acroFormQuadding = ReadFieldQuadding(objects, acroForm, null);
        PdfDictionary? acroFormDefaultResources = TryReadDefaultResources(objects, acroForm);
        string? acroFormDefaultAppearance = TryReadText(objects, acroForm, "DA");
        for (int i = 0; i < fields.Items.Count; i++) {
            FillField(objects, fields.Items[i], null, null, 0, acroFormQuadding, null, acroFormDefaultResources, acroFormDefaultAppearance, null, fieldValues, options, remaining, new HashSet<int>(), ref nextObjectNumber);
        }

        if (remaining.Count > 0) {
            throw new ArgumentException("PDF form field was not found: " + string.Join(", ", remaining), nameof(fieldValues));
        }

        acroForm.Items["NeedAppearances"] = new PdfBoolean(options?.KeepNeedAppearances == true);
        return RewriteAllObjects(objects, catalogObjectNumber, PdfReadDocument.Open(pdf, readOptions).Metadata, pdf);
    }

    /// <summary>
    /// Returns a new PDF with simple AcroForm field values updated from the current position of a readable stream.
    /// </summary>
    public static byte[] FillFields(Stream stream, IReadOnlyDictionary<string, string> fieldValues) {
        return FillFields(stream, fieldValues, options: null);
    }

    /// <summary>
    /// Returns a new PDF with simple AcroForm field values updated from the current position of a readable stream.
    /// </summary>
    public static byte[] FillFields(Stream stream, IReadOnlyDictionary<string, string> fieldValues, PdfFormFillerOptions? options) {
        return FillFields(ReadStream(stream, nameof(stream)), fieldValues, options);
    }

    /// <summary>
    /// Returns a new PDF with simple AcroForm field values updated from the current position of a readable stream.
    /// </summary>
    public static byte[] FillFields(Stream stream, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        return FillFields(stream, fieldValues, options: null);
    }

    /// <summary>
    /// Returns a new PDF with simple AcroForm field values updated from the current position of a readable stream.
    /// </summary>
    public static byte[] FillFields(Stream stream, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfFormFillerOptions? options) {
        return FillFields(ReadStream(stream, nameof(stream)), fieldValues, options);
    }

    /// <summary>
    /// Writes a new PDF with simple AcroForm field values updated.
    /// </summary>
    public static void FillFields(byte[] pdf, Stream outputStream, IReadOnlyDictionary<string, string> fieldValues) {
        FillFields(pdf, outputStream, fieldValues, options: null);
    }

    /// <summary>
    /// Writes a new PDF with simple AcroForm field values updated.
    /// </summary>
    public static void FillFields(byte[] pdf, Stream outputStream, IReadOnlyDictionary<string, string> fieldValues, PdfFormFillerOptions? options) {
        WriteOutput(outputStream, FillFields(pdf, fieldValues, options));
    }

    /// <summary>
    /// Writes a new PDF with simple AcroForm field values updated.
    /// </summary>
    public static void FillFields(byte[] pdf, Stream outputStream, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        FillFields(pdf, outputStream, fieldValues, options: null);
    }

    /// <summary>
    /// Writes a new PDF with simple AcroForm field values updated.
    /// </summary>
    public static void FillFields(byte[] pdf, Stream outputStream, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfFormFillerOptions? options) {
        WriteOutput(outputStream, FillFields(pdf, fieldValues, options));
    }

    /// <summary>
    /// Writes a new PDF with simple AcroForm field values updated from the current position of a readable stream.
    /// </summary>
    public static void FillFields(Stream inputStream, Stream outputStream, IReadOnlyDictionary<string, string> fieldValues) {
        FillFields(inputStream, outputStream, fieldValues, options: null);
    }

    /// <summary>
    /// Writes a new PDF with simple AcroForm field values updated from the current position of a readable stream.
    /// </summary>
    public static void FillFields(Stream inputStream, Stream outputStream, IReadOnlyDictionary<string, string> fieldValues, PdfFormFillerOptions? options) {
        WriteOutput(outputStream, FillFields(inputStream, fieldValues, options));
    }

    /// <summary>
    /// Writes a new PDF with simple AcroForm field values updated from the current position of a readable stream.
    /// </summary>
    public static void FillFields(Stream inputStream, Stream outputStream, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        FillFields(inputStream, outputStream, fieldValues, options: null);
    }

    /// <summary>
    /// Writes a new PDF with simple AcroForm field values updated from the current position of a readable stream.
    /// </summary>
    public static void FillFields(Stream inputStream, Stream outputStream, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfFormFillerOptions? options) {
        WriteOutput(outputStream, FillFields(inputStream, fieldValues, options));
    }

    /// <summary>
    /// Writes a new PDF file with simple AcroForm field values updated.
    /// </summary>
    public static void FillFields(string inputPath, string outputPath, IReadOnlyDictionary<string, string> fieldValues) {
        FillFields(inputPath, outputPath, fieldValues, options: null);
    }

    /// <summary>
    /// Writes a new PDF file with simple AcroForm field values updated.
    /// </summary>
    public static void FillFields(string inputPath, string outputPath, IReadOnlyDictionary<string, string> fieldValues, PdfFormFillerOptions? options) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        string fullOutputPath = ValidateOutputPath(outputPath);
        byte[] bytes = FillFields(File.ReadAllBytes(inputPath), fieldValues, options);
        var directory = Path.GetDirectoryName(fullOutputPath);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        OfficeFileCommit.WriteAllBytes(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF file with simple AcroForm field values updated.
    /// </summary>
    public static void FillFields(string inputPath, string outputPath, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        FillFields(inputPath, outputPath, fieldValues, options: null);
    }

    /// <summary>
    /// Writes a new PDF file with simple AcroForm field values updated.
    /// </summary>
    public static void FillFields(string inputPath, string outputPath, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfFormFillerOptions? options) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        string fullOutputPath = ValidateOutputPath(outputPath);
        byte[] bytes = FillFields(File.ReadAllBytes(inputPath), fieldValues, options);
        var directory = Path.GetDirectoryName(fullOutputPath);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        OfficeFileCommit.WriteAllBytes(fullOutputPath, bytes);
    }

    /// <summary>
    /// Reads a PDF file and writes a new PDF with simple AcroForm field values updated.
    /// </summary>
    public static void FillFields(string inputPath, Stream outputStream, IReadOnlyDictionary<string, string> fieldValues) {
        FillFields(inputPath, outputStream, fieldValues, options: null);
    }

    /// <summary>
    /// Reads a PDF file and writes a new PDF with simple AcroForm field values updated.
    /// </summary>
    public static void FillFields(string inputPath, Stream outputStream, IReadOnlyDictionary<string, string> fieldValues, PdfFormFillerOptions? options) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        byte[] bytes = FillFields(File.ReadAllBytes(inputPath), fieldValues, options);
        WriteOutput(outputStream, bytes);
    }

    /// <summary>
    /// Reads a PDF file and writes a new PDF with simple AcroForm field values updated.
    /// </summary>
    public static void FillFields(string inputPath, Stream outputStream, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        FillFields(inputPath, outputStream, fieldValues, options: null);
    }

    /// <summary>
    /// Reads a PDF file and writes a new PDF with simple AcroForm field values updated.
    /// </summary>
    public static void FillFields(string inputPath, Stream outputStream, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfFormFillerOptions? options) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        byte[] bytes = FillFields(File.ReadAllBytes(inputPath), fieldValues, options);
        WriteOutput(outputStream, bytes);
    }

    /// <summary>
    /// Reads a PDF file and returns new PDF bytes with simple AcroForm field values updated.
    /// </summary>
    public static byte[] FillFieldsToBytes(string inputPath, IReadOnlyDictionary<string, string> fieldValues) {
        return FillFieldsToBytes(inputPath, fieldValues, options: null);
    }

    /// <summary>
    /// Reads a PDF file and returns new PDF bytes with simple AcroForm field values updated.
    /// </summary>
    public static byte[] FillFieldsToBytes(string inputPath, IReadOnlyDictionary<string, string> fieldValues, PdfFormFillerOptions? options) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return FillFields(File.ReadAllBytes(inputPath), fieldValues, options);
    }

    /// <summary>
    /// Reads a PDF file and returns new PDF bytes with simple AcroForm field values updated.
    /// </summary>
    public static byte[] FillFieldsToBytes(string inputPath, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        return FillFieldsToBytes(inputPath, fieldValues, options: null);
    }

    /// <summary>
    /// Reads a PDF file and returns new PDF bytes with simple AcroForm field values updated.
    /// </summary>
    public static byte[] FillFieldsToBytes(string inputPath, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfFormFillerOptions? options) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return FillFields(File.ReadAllBytes(inputPath), fieldValues, options);
    }

    /// <summary>
    /// Returns a new PDF with simple text, choice, and button AcroForm widgets painted into page content and removed from the form tree.
    /// </summary>
    public static byte[] FlattenFields(byte[] pdf) {
        return FlattenFields(pdf, options: null, readOptions: null);
    }

    /// <summary>
    /// Returns a new PDF with simple text, choice, and button AcroForm widgets painted into page content and removed from the form tree.
    /// </summary>
    public static byte[] FlattenFields(byte[] pdf, PdfFormFillerOptions? options) {
        return FlattenFields(pdf, options, readOptions: null);
    }

    internal static byte[] FlattenFields(byte[] pdf, PdfFormFillerOptions? options, PdfReadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        _ = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.FlattenFormFields, readOptions);

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, readOptions);
        int catalogObjectNumber = FindCatalogObjectNumber(objects, trailerRaw);
        if (catalogObjectNumber == 0 ||
            objects[catalogObjectNumber].Value is not PdfDictionary catalog ||
            !catalog.Items.TryGetValue("AcroForm", out var acroFormObject) ||
            ResolveDictionary(objects, acroFormObject) is not PdfDictionary acroForm ||
            !acroForm.Items.TryGetValue("Fields", out var fieldsObject) ||
            ResolveObject(objects, fieldsObject) is not PdfArray fields) {
            throw new ArgumentException("PDF does not contain a readable AcroForm field tree.", nameof(pdf));
        }

        int nextObjectNumber = objects.Keys.Count == 0 ? 1 : objects.Keys.Max() + 1;
        PdfDictionary? acroFormDefaultResources = TryReadDefaultResources(objects, acroForm);
        string? acroFormDefaultAppearance = TryReadText(objects, acroForm, "DA");
        var widgets = new Dictionary<int, FlattenWidgetState>();
        var removableObjects = new HashSet<int>();
        for (int i = 0; i < fields.Items.Count; i++) {
            CollectFlattenWidgets(objects, fields.Items[i], null, 0, null, acroFormDefaultResources, acroFormDefaultAppearance, null, null, null, null, options, widgets, removableObjects, new HashSet<int>(), ref nextObjectNumber);
        }

        if (widgets.Count == 0) {
            throw new NotSupportedException(UnsupportedFlattenWidgetMessage);
        }

        int flattenedWidgetCount = FlattenPageWidgets(objects, widgets, ref nextObjectNumber);
        if (flattenedWidgetCount != widgets.Count) {
            throw new NotSupportedException(UnsupportedFlattenAnnotationMessage);
        }

        catalog.Items.Remove("AcroForm");
        if (acroFormObject is PdfReference acroFormReference) {
            removableObjects.Add(acroFormReference.ObjectNumber);
        }

        foreach (int objectNumber in removableObjects) {
            objects.Remove(objectNumber);
        }

        return RewriteAllObjects(objects, catalogObjectNumber, PdfReadDocument.Open(pdf, readOptions).Metadata, pdf);
    }

    /// <summary>
    /// Returns a new PDF with only the named simple text, choice, and button AcroForm fields painted into page content and removed from the form tree.
    /// </summary>
    public static byte[] FlattenFields(byte[] pdf, IReadOnlyCollection<string> fieldNames, PdfFormFillerOptions? options = null) {
        return FlattenFieldsCore(pdf, fieldNames, options, readOptions: null, requireMutationPlan: true);
    }

    internal static byte[] FlattenFields(byte[] pdf, IReadOnlyCollection<string> fieldNames, PdfFormFillerOptions? options, PdfReadOptions? readOptions) =>
        FlattenFieldsCore(pdf, fieldNames, options, readOptions, requireMutationPlan: true);

    internal static byte[] FlattenFieldsWithinPlannedRewrite(byte[] pdf, IReadOnlyCollection<string> fieldNames, PdfFormFillerOptions? options = null) {
        return FlattenFieldsCore(pdf, fieldNames, options, readOptions: null, requireMutationPlan: false);
    }

    private static byte[] FlattenFieldsCore(byte[] pdf, IReadOnlyCollection<string> fieldNames, PdfFormFillerOptions? options, PdfReadOptions? readOptions, bool requireMutationPlan) {
        Guard.NotNull(pdf, nameof(pdf));
        ValidateFlattenFieldNames(fieldNames);
        if (requireMutationPlan) _ = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.FlattenFormFields, readOptions, fieldNames: fieldNames);

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, readOptions);
        int catalogObjectNumber = FindCatalogObjectNumber(objects, trailerRaw);
        if (catalogObjectNumber == 0 ||
            objects[catalogObjectNumber].Value is not PdfDictionary catalog ||
            !catalog.Items.TryGetValue("AcroForm", out var acroFormObject) ||
            ResolveDictionary(objects, acroFormObject) is not PdfDictionary acroForm ||
            !acroForm.Items.TryGetValue("Fields", out var fieldsObject) ||
            ResolveObject(objects, fieldsObject) is not PdfArray fields) {
            throw new ArgumentException("PDF does not contain a readable AcroForm field tree.", nameof(pdf));
        }

        var requested = new HashSet<string>(fieldNames, StringComparer.Ordinal);
        var matched = new HashSet<string>(StringComparer.Ordinal);
        var widgets = new Dictionary<int, FlattenWidgetState>();
        var removableObjects = new HashSet<int>();
        int nextObjectNumber = objects.Keys.Count == 0 ? 1 : objects.Keys.Max() + 1;
        CollectSelectedFlattenFields(
            objects,
            fields,
            requested,
            matched,
            inheritedFieldType: null,
            inheritedFlags: 0,
            inheritedMaxLength: null,
            inheritedDefaultResources: TryReadDefaultResources(objects, acroForm),
            inheritedDefaultAppearance: TryReadText(objects, acroForm, "DA"),
            inheritedDisplayValue: null,
            inheritedRichAppearanceRuns: null,
            inheritedName: null,
            inheritedChoiceOptions: null,
            options,
            widgets,
            removableObjects,
            ref nextObjectNumber);

        if (matched.Count != requested.Count) {
            throw new ArgumentException("PDF form field was not found: " + string.Join(", ", requested.Where(name => !matched.Contains(name))), nameof(fieldNames));
        }
        if (widgets.Count == 0) {
            throw new NotSupportedException(UnsupportedFlattenWidgetMessage);
        }

        int flattenedWidgetCount = FlattenPageWidgets(objects, widgets, ref nextObjectNumber);
        if (flattenedWidgetCount != widgets.Count) {
            throw new NotSupportedException(UnsupportedFlattenAnnotationMessage);
        }

        FilterCalculationOrder(objects, acroForm, removableObjects);
        if (fields.Items.Count == 0) {
            catalog.Items.Remove("AcroForm");
            if (acroFormObject is PdfReference acroFormReference) removableObjects.Add(acroFormReference.ObjectNumber);
        }

        foreach (int objectNumber in removableObjects) objects.Remove(objectNumber);
        return RewriteAllObjects(objects, catalogObjectNumber, PdfReadDocument.Open(pdf, readOptions).Metadata, pdf);
    }

    /// <summary>
    /// Returns a new PDF with simple text, choice, and button AcroForm widgets flattened from the current position of a readable stream.
    /// </summary>
    public static byte[] FlattenFields(Stream stream) {
        return FlattenFields(stream, options: null);
    }

    /// <summary>
    /// Returns a new PDF with simple text, choice, and button AcroForm widgets flattened from the current position of a readable stream.
    /// </summary>
    public static byte[] FlattenFields(Stream stream, PdfFormFillerOptions? options) {
        return FlattenFields(ReadStream(stream, nameof(stream)), options);
    }

    /// <summary>
    /// Writes a new PDF with simple text, choice, and button AcroForm widgets flattened.
    /// </summary>
    public static void FlattenFields(byte[] pdf, Stream outputStream) {
        FlattenFields(pdf, outputStream, options: null);
    }

    /// <summary>
    /// Writes a new PDF with simple text, choice, and button AcroForm widgets flattened.
    /// </summary>
    public static void FlattenFields(byte[] pdf, Stream outputStream, PdfFormFillerOptions? options) {
        WriteOutput(outputStream, FlattenFields(pdf, options));
    }

    /// <summary>
    /// Writes a new PDF with simple text, choice, and button AcroForm widgets flattened from the current position of a readable stream.
    /// </summary>
    public static void FlattenFields(Stream inputStream, Stream outputStream) {
        FlattenFields(inputStream, outputStream, options: null);
    }

    /// <summary>
    /// Writes a new PDF with simple text, choice, and button AcroForm widgets flattened from the current position of a readable stream.
    /// </summary>
    public static void FlattenFields(Stream inputStream, Stream outputStream, PdfFormFillerOptions? options) {
        WriteOutput(outputStream, FlattenFields(inputStream, options));
    }

    /// <summary>
    /// Writes a new PDF file with simple text, choice, and button AcroForm widgets flattened.
    /// </summary>
    public static void FlattenFields(string inputPath, string outputPath) {
        FlattenFields(inputPath, outputPath, options: null);
    }

    /// <summary>
    /// Writes a new PDF file with simple text, choice, and button AcroForm widgets flattened.
    /// </summary>
    public static void FlattenFields(string inputPath, string outputPath, PdfFormFillerOptions? options) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        string fullOutputPath = ValidateOutputPath(outputPath);
        byte[] bytes = FlattenFields(File.ReadAllBytes(inputPath), options);
        var directory = Path.GetDirectoryName(fullOutputPath);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        OfficeFileCommit.WriteAllBytes(fullOutputPath, bytes);
    }

    /// <summary>
    /// Reads a PDF file and writes a new PDF with simple text, choice, and button AcroForm widgets flattened.
    /// </summary>
    public static void FlattenFields(string inputPath, Stream outputStream) {
        FlattenFields(inputPath, outputStream, options: null);
    }

    /// <summary>
    /// Reads a PDF file and writes a new PDF with simple text, choice, and button AcroForm widgets flattened.
    /// </summary>
    public static void FlattenFields(string inputPath, Stream outputStream, PdfFormFillerOptions? options) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        byte[] bytes = FlattenFields(File.ReadAllBytes(inputPath), options);
        WriteOutput(outputStream, bytes);
    }

    /// <summary>
    /// Reads a PDF file and returns new PDF bytes with simple text, choice, and button AcroForm widgets flattened.
    /// </summary>
    public static byte[] FlattenFieldsToBytes(string inputPath) {
        return FlattenFieldsToBytes(inputPath, options: null);
    }

    /// <summary>
    /// Reads a PDF file and returns new PDF bytes with simple text, choice, and button AcroForm widgets flattened.
    /// </summary>
    public static byte[] FlattenFieldsToBytes(string inputPath, PdfFormFillerOptions? options) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return FlattenFields(File.ReadAllBytes(inputPath), options);
    }

    /// <summary>
    /// Returns a new PDF with simple AcroForm field values updated and then flattened into page content.
    /// </summary>
    public static byte[] FillAndFlattenFields(byte[] pdf, IReadOnlyDictionary<string, string> fieldValues) {
        return FillAndFlattenFields(pdf, fieldValues, options: null);
    }

    /// <summary>
    /// Returns a new PDF with simple AcroForm field values updated and then flattened into page content.
    /// </summary>
    public static byte[] FillAndFlattenFields(byte[] pdf, IReadOnlyDictionary<string, string> fieldValues, PdfFormFillerOptions? options) {
        return FlattenFields(FillFields(pdf, fieldValues, options), options);
    }

    internal static byte[] FillAndFlattenFields(byte[] pdf, IReadOnlyDictionary<string, string> fieldValues, PdfFormFillerOptions? options, PdfReadOptions? readOptions) =>
        FlattenFields(FillFields(pdf, fieldValues, options, readOptions), options);

    /// <summary>
    /// Returns a new PDF with simple AcroForm field values updated and then flattened into page content.
    /// </summary>
    public static byte[] FillAndFlattenFields(byte[] pdf, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        return FillAndFlattenFields(pdf, fieldValues, options: null);
    }

    /// <summary>
    /// Returns a new PDF with simple AcroForm field values updated and then flattened into page content.
    /// </summary>
    public static byte[] FillAndFlattenFields(byte[] pdf, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfFormFillerOptions? options) {
        return FlattenFields(FillFields(pdf, fieldValues, options), options);
    }

    internal static byte[] FillAndFlattenFields(byte[] pdf, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfFormFillerOptions? options, PdfReadOptions? readOptions) =>
        FlattenFields(FillFields(pdf, fieldValues, options, readOptions), options);

    /// <summary>
    /// Returns a new PDF with simple AcroForm field values updated and flattened from the current position of a readable stream.
    /// </summary>
    public static byte[] FillAndFlattenFields(Stream stream, IReadOnlyDictionary<string, string> fieldValues) {
        return FillAndFlattenFields(stream, fieldValues, options: null);
    }

    /// <summary>
    /// Returns a new PDF with simple AcroForm field values updated and flattened from the current position of a readable stream.
    /// </summary>
    public static byte[] FillAndFlattenFields(Stream stream, IReadOnlyDictionary<string, string> fieldValues, PdfFormFillerOptions? options) {
        return FillAndFlattenFields(ReadStream(stream, nameof(stream)), fieldValues, options);
    }

    /// <summary>
    /// Returns a new PDF with simple AcroForm field values updated and flattened from the current position of a readable stream.
    /// </summary>
    public static byte[] FillAndFlattenFields(Stream stream, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        return FillAndFlattenFields(stream, fieldValues, options: null);
    }

    /// <summary>
    /// Returns a new PDF with simple AcroForm field values updated and flattened from the current position of a readable stream.
    /// </summary>
    public static byte[] FillAndFlattenFields(Stream stream, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfFormFillerOptions? options) {
        return FillAndFlattenFields(ReadStream(stream, nameof(stream)), fieldValues, options);
    }

    /// <summary>
    /// Writes a new PDF with simple AcroForm field values updated and flattened.
    /// </summary>
    public static void FillAndFlattenFields(byte[] pdf, Stream outputStream, IReadOnlyDictionary<string, string> fieldValues) {
        FillAndFlattenFields(pdf, outputStream, fieldValues, options: null);
    }

    /// <summary>
    /// Writes a new PDF with simple AcroForm field values updated and flattened.
    /// </summary>
    public static void FillAndFlattenFields(byte[] pdf, Stream outputStream, IReadOnlyDictionary<string, string> fieldValues, PdfFormFillerOptions? options) {
        WriteOutput(outputStream, FillAndFlattenFields(pdf, fieldValues, options));
    }

    /// <summary>
    /// Writes a new PDF with simple AcroForm field values updated and flattened.
    /// </summary>
    public static void FillAndFlattenFields(byte[] pdf, Stream outputStream, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        FillAndFlattenFields(pdf, outputStream, fieldValues, options: null);
    }

    /// <summary>
    /// Writes a new PDF with simple AcroForm field values updated and flattened.
    /// </summary>
    public static void FillAndFlattenFields(byte[] pdf, Stream outputStream, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfFormFillerOptions? options) {
        WriteOutput(outputStream, FillAndFlattenFields(pdf, fieldValues, options));
    }

    /// <summary>
    /// Writes a new PDF with simple AcroForm field values updated and flattened from the current position of a readable stream.
    /// </summary>
    public static void FillAndFlattenFields(Stream inputStream, Stream outputStream, IReadOnlyDictionary<string, string> fieldValues) {
        FillAndFlattenFields(inputStream, outputStream, fieldValues, options: null);
    }

    /// <summary>
    /// Writes a new PDF with simple AcroForm field values updated and flattened from the current position of a readable stream.
    /// </summary>
    public static void FillAndFlattenFields(Stream inputStream, Stream outputStream, IReadOnlyDictionary<string, string> fieldValues, PdfFormFillerOptions? options) {
        WriteOutput(outputStream, FillAndFlattenFields(inputStream, fieldValues, options));
    }

    /// <summary>
    /// Writes a new PDF with simple AcroForm field values updated and flattened from the current position of a readable stream.
    /// </summary>
    public static void FillAndFlattenFields(Stream inputStream, Stream outputStream, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        FillAndFlattenFields(inputStream, outputStream, fieldValues, options: null);
    }

    /// <summary>
    /// Writes a new PDF with simple AcroForm field values updated and flattened from the current position of a readable stream.
    /// </summary>
    public static void FillAndFlattenFields(Stream inputStream, Stream outputStream, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfFormFillerOptions? options) {
        WriteOutput(outputStream, FillAndFlattenFields(inputStream, fieldValues, options));
    }

    /// <summary>
    /// Writes a new PDF file with simple AcroForm field values updated and flattened.
    /// </summary>
    public static void FillAndFlattenFields(string inputPath, string outputPath, IReadOnlyDictionary<string, string> fieldValues) {
        FillAndFlattenFields(inputPath, outputPath, fieldValues, options: null);
    }

    /// <summary>
    /// Writes a new PDF file with simple AcroForm field values updated and flattened.
    /// </summary>
    public static void FillAndFlattenFields(string inputPath, string outputPath, IReadOnlyDictionary<string, string> fieldValues, PdfFormFillerOptions? options) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        string fullOutputPath = ValidateOutputPath(outputPath);
        byte[] bytes = FillAndFlattenFields(File.ReadAllBytes(inputPath), fieldValues, options);
        var directory = Path.GetDirectoryName(fullOutputPath);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        OfficeFileCommit.WriteAllBytes(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF file with simple AcroForm field values updated and flattened.
    /// </summary>
    public static void FillAndFlattenFields(string inputPath, string outputPath, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        FillAndFlattenFields(inputPath, outputPath, fieldValues, options: null);
    }

    /// <summary>
    /// Writes a new PDF file with simple AcroForm field values updated and flattened.
    /// </summary>
    public static void FillAndFlattenFields(string inputPath, string outputPath, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfFormFillerOptions? options) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        string fullOutputPath = ValidateOutputPath(outputPath);
        byte[] bytes = FillAndFlattenFields(File.ReadAllBytes(inputPath), fieldValues, options);
        var directory = Path.GetDirectoryName(fullOutputPath);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        OfficeFileCommit.WriteAllBytes(fullOutputPath, bytes);
    }

    /// <summary>
    /// Reads a PDF file and writes a new PDF with simple AcroForm field values updated and flattened.
    /// </summary>
    public static void FillAndFlattenFields(string inputPath, Stream outputStream, IReadOnlyDictionary<string, string> fieldValues) {
        FillAndFlattenFields(inputPath, outputStream, fieldValues, options: null);
    }

    /// <summary>
    /// Reads a PDF file and writes a new PDF with simple AcroForm field values updated and flattened.
    /// </summary>
    public static void FillAndFlattenFields(string inputPath, Stream outputStream, IReadOnlyDictionary<string, string> fieldValues, PdfFormFillerOptions? options) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        byte[] bytes = FillAndFlattenFields(File.ReadAllBytes(inputPath), fieldValues, options);
        WriteOutput(outputStream, bytes);
    }

    /// <summary>
    /// Reads a PDF file and writes a new PDF with simple AcroForm field values updated and flattened.
    /// </summary>
    public static void FillAndFlattenFields(string inputPath, Stream outputStream, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        FillAndFlattenFields(inputPath, outputStream, fieldValues, options: null);
    }

    /// <summary>
    /// Reads a PDF file and writes a new PDF with simple AcroForm field values updated and flattened.
    /// </summary>
    public static void FillAndFlattenFields(string inputPath, Stream outputStream, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfFormFillerOptions? options) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        byte[] bytes = FillAndFlattenFields(File.ReadAllBytes(inputPath), fieldValues, options);
        WriteOutput(outputStream, bytes);
    }

    /// <summary>
    /// Reads a PDF file and returns new PDF bytes with simple AcroForm field values updated and flattened.
    /// </summary>
    public static byte[] FillAndFlattenFieldsToBytes(string inputPath, IReadOnlyDictionary<string, string> fieldValues) {
        return FillAndFlattenFieldsToBytes(inputPath, fieldValues, options: null);
    }

    /// <summary>
    /// Reads a PDF file and returns new PDF bytes with simple AcroForm field values updated and flattened.
    /// </summary>
    public static byte[] FillAndFlattenFieldsToBytes(string inputPath, IReadOnlyDictionary<string, string> fieldValues, PdfFormFillerOptions? options) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return FillAndFlattenFields(File.ReadAllBytes(inputPath), fieldValues, options);
    }

    /// <summary>
    /// Reads a PDF file and returns new PDF bytes with simple AcroForm field values updated and flattened.
    /// </summary>
    public static byte[] FillAndFlattenFieldsToBytes(string inputPath, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        return FillAndFlattenFieldsToBytes(inputPath, fieldValues, options: null);
    }

    /// <summary>
    /// Reads a PDF file and returns new PDF bytes with simple AcroForm field values updated and flattened.
    /// </summary>
    public static byte[] FillAndFlattenFieldsToBytes(string inputPath, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfFormFillerOptions? options) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return FillAndFlattenFields(File.ReadAllBytes(inputPath), fieldValues, options);
    }
}
