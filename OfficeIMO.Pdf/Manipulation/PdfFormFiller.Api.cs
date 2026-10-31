using OfficeIMO.Core.Internal;
using System.Threading;
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
        return FillFields(pdf, ToFormFieldValues(fieldValues, GetCancellationToken(options)), options, readOptions: null);
    }

    internal static byte[] FillFields(byte[] pdf, IReadOnlyDictionary<string, string> fieldValues, PdfFormFillerOptions? options, PdfLoadOptions? readOptions) =>
        FillFields(pdf, ToFormFieldValues(fieldValues, GetCancellationToken(options)), options, readOptions);

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

    internal static byte[] FillFields(byte[] pdf, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfFormFillerOptions? options,
        PdfLoadOptions? readOptions, CancellationToken cancellationToken = default) =>
        FillFieldsCore(pdf, fieldValues, options, readOptions, requireMutationPlan: true, cancellationToken);

    internal static byte[] FillFieldsWithinPlannedRewrite(byte[] pdf, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfFormFillerOptions? options = null, PdfLoadOptions? readOptions = null, CancellationToken cancellationToken = default) =>
        FillFieldsWithinPlannedRewrite(pdf, fieldValues, options, readOptions, out _, cancellationToken);

    internal static byte[] FillFieldsWithinPlannedRewrite(byte[] pdf, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues,
        PdfFormFillerOptions? options, PdfLoadOptions? readOptions, out PdfLoadOptions outputReadOptions, CancellationToken cancellationToken = default) =>
        FillFieldsCore(pdf, fieldValues, options, readOptions, requireMutationPlan: false, out outputReadOptions, cancellationToken);

    internal static byte[] FillFields(byte[] pdf, IReadOnlyDictionary<string, string> fieldValues, PdfFormFillerOptions? options,
        PdfLoadOptions? readOptions, out PdfLoadOptions outputReadOptions) =>
        FillFields(pdf, ToFormFieldValues(fieldValues, GetCancellationToken(options)), options, readOptions, out outputReadOptions);

    internal static byte[] FillFields(byte[] pdf, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfFormFillerOptions? options,
        PdfLoadOptions? readOptions, out PdfLoadOptions outputReadOptions, CancellationToken cancellationToken = default) =>
        FillFieldsCore(pdf, fieldValues, options, readOptions, requireMutationPlan: true, out outputReadOptions, cancellationToken);

    private static byte[] FillFieldsCore(byte[] pdf, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfFormFillerOptions? options,
        PdfLoadOptions? readOptions, bool requireMutationPlan, CancellationToken cancellationToken = default) =>
        FillFieldsCore(pdf, fieldValues, options, readOptions, requireMutationPlan, out _, cancellationToken);

    private static byte[] FillFieldsCore(byte[] pdf, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfFormFillerOptions? options,
        PdfLoadOptions? readOptions, bool requireMutationPlan, out PdfLoadOptions outputReadOptions, CancellationToken cancellationToken = default) {
        using var linkedCancellation = LinkCancellation(options?.CancellationToken ?? default, ref cancellationToken);
        Guard.NotNull(pdf, nameof(pdf));
        cancellationToken.ThrowIfCancellationRequested();
        ValidateFieldValues(fieldValues, cancellationToken);
        PdfReadDocument source = RejectPushButtonFillValues(pdf, fieldValues.Keys, readOptions, cancellationToken);
        if (requireMutationPlan) _ = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.FillFormFields, readOptions,
            fieldNames: fieldValues.Keys, cancellationToken);
        PdfFlattenedContentGrowth.ValidateSourceAppearances(source, cancellationToken);

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, readOptions, out _, out _, cancellationToken);
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
        int firstGeneratedObject = nextObjectNumber;
        int? acroFormQuadding = ReadFieldQuadding(objects, acroForm, null);
        PdfDictionary? acroFormDefaultResources = TryReadDefaultResources(objects, acroForm);
        string? acroFormDefaultAppearance = TryReadText(objects, acroForm, "DA");
        for (int i = 0; i < fields.Items.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            FillField(objects, fields.Items[i], null, null, 0, acroFormQuadding, null, acroFormDefaultResources, acroFormDefaultAppearance, null, fieldValues, options, remaining, new HashSet<int>(), ref nextObjectNumber,
                source.ReadOptions.Limits.MaxDecodedTextCharacters, cancellationToken);
        }

        if (remaining.Count > 0) {
            throw new ArgumentException("PDF form field was not found: " + string.Join(", ", remaining), nameof(fieldValues));
        }

        acroForm.Items["NeedAppearances"] = new PdfBoolean(options?.KeepNeedAppearances == true);
        PdfGeneratedOutputGrowth growth = GetFilledAppearanceGrowth(objects, firstGeneratedObject, cancellationToken);
        return RewriteAllObjects(objects, catalogObjectNumber, source, pdf, growth, out outputReadOptions, cancellationToken);
    }

    private static PdfReadDocument RejectPushButtonFillValues(byte[] pdf, IEnumerable<string> fieldNames, PdfLoadOptions? readOptions,
        CancellationToken cancellationToken) {
        PdfReadDocument document = PdfReadDocument.Open(pdf, readOptions, cancellationToken);
        IReadOnlyDictionary<string, PdfFormField> fields = PdfInspector.Inspect(pdf, document, cancellationToken).FormFieldsByName;
        foreach (string fieldName in fieldNames) {
            cancellationToken.ThrowIfCancellationRequested();
            if (fields.TryGetValue(fieldName, out PdfFormField? field) && field.IsPushButton) {
                throw new ArgumentException("Push-button fields do not have a fillable value: " + fieldName, nameof(fieldNames));
            }
        }
        return document;
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
        return FillFields(stream, fieldValues, options, readOptions: null);
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
        return FillFields(stream, fieldValues, options, readOptions: null);
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
        WriteOutput(outputStream, FillFields(pdf, fieldValues, options), options);
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
        WriteOutput(outputStream, FillFields(pdf, fieldValues, options), options);
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
        WriteOutput(outputStream, FillFields(inputStream, fieldValues, options), options);
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
        WriteOutput(outputStream, FillFields(inputStream, fieldValues, options), options);
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
        byte[] bytes = FillFields(ReadPath(inputPath, options), fieldValues, options);
        var directory = Path.GetDirectoryName(fullOutputPath);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        OfficeFileCommit.Write(fullOutputPath, output => WriteOutput(output, bytes, options), GetCancellationToken(options));
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
        byte[] bytes = FillFields(ReadPath(inputPath, options), fieldValues, options);
        var directory = Path.GetDirectoryName(fullOutputPath);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        OfficeFileCommit.Write(fullOutputPath, output => WriteOutput(output, bytes, options), GetCancellationToken(options));
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

        byte[] bytes = FillFields(ReadPath(inputPath, options), fieldValues, options);
        WriteOutput(outputStream, bytes, options);
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

        byte[] bytes = FillFields(ReadPath(inputPath, options), fieldValues, options);
        WriteOutput(outputStream, bytes, options);
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
        return FillFieldsToBytes(inputPath, fieldValues, options, readOptions: null);
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
        return FillFields(ReadPath(inputPath, options), fieldValues, options);
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
        return FlattenFields(stream, options, readOptions: null);
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
        WriteOutput(outputStream, FlattenFields(pdf, options), options);
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
        WriteOutput(outputStream, FlattenFields(inputStream, options), options);
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
        byte[] bytes = FlattenFields(ReadPath(inputPath, options), options);
        var directory = Path.GetDirectoryName(fullOutputPath);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        OfficeFileCommit.Write(fullOutputPath, output => WriteOutput(output, bytes, options), GetCancellationToken(options));
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

        byte[] bytes = FlattenFields(ReadPath(inputPath, options), options);
        WriteOutput(outputStream, bytes, options);
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
        return FlattenFieldsToBytes(inputPath, options, readOptions: null);
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
        return FillAndFlattenFields(pdf, fieldValues, options, readOptions: null);
    }

    internal static byte[] FillAndFlattenFields(byte[] pdf, IReadOnlyDictionary<string, string> fieldValues, PdfFormFillerOptions? options, PdfLoadOptions? readOptions) =>
        FillAndFlattenFields(pdf, fieldValues, options, readOptions, out _);

    internal static byte[] FillAndFlattenFields(byte[] pdf, IReadOnlyDictionary<string, string> fieldValues,
        PdfFormFillerOptions? options, PdfLoadOptions? readOptions, out PdfLoadOptions outputReadOptions) {
        byte[] filled = FillFields(pdf, fieldValues, options, readOptions, out PdfLoadOptions filledReadOptions);
        return FlattenFields(filled, options, filledReadOptions, out outputReadOptions);
    }

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
        return FillAndFlattenFields(pdf, fieldValues, options, readOptions: null);
    }

    internal static byte[] FillAndFlattenFields(byte[] pdf, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfFormFillerOptions? options, PdfLoadOptions? readOptions) =>
        FillAndFlattenFields(pdf, fieldValues, options, readOptions, out _);

    internal static byte[] FillAndFlattenFields(byte[] pdf, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues,
        PdfFormFillerOptions? options, PdfLoadOptions? readOptions, out PdfLoadOptions outputReadOptions) {
        byte[] filled = FillFields(pdf, fieldValues, options, readOptions, out PdfLoadOptions filledReadOptions);
        return FlattenFields(filled, options, filledReadOptions, out outputReadOptions);
    }

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
        return FillAndFlattenFields(stream, fieldValues, options, readOptions: null);
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
        return FillAndFlattenFields(stream, fieldValues, options, readOptions: null);
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
        WriteOutput(outputStream, FillAndFlattenFields(pdf, fieldValues, options), options);
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
        WriteOutput(outputStream, FillAndFlattenFields(pdf, fieldValues, options), options);
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
        WriteOutput(outputStream, FillAndFlattenFields(inputStream, fieldValues, options), options);
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
        WriteOutput(outputStream, FillAndFlattenFields(inputStream, fieldValues, options), options);
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
        byte[] bytes = FillAndFlattenFields(ReadPath(inputPath, options), fieldValues, options);
        var directory = Path.GetDirectoryName(fullOutputPath);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        OfficeFileCommit.Write(fullOutputPath, output => WriteOutput(output, bytes, options), GetCancellationToken(options));
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
        byte[] bytes = FillAndFlattenFields(ReadPath(inputPath, options), fieldValues, options);
        var directory = Path.GetDirectoryName(fullOutputPath);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        OfficeFileCommit.Write(fullOutputPath, output => WriteOutput(output, bytes, options), GetCancellationToken(options));
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

        byte[] bytes = FillAndFlattenFields(ReadPath(inputPath, options), fieldValues, options);
        WriteOutput(outputStream, bytes, options);
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

        byte[] bytes = FillAndFlattenFields(ReadPath(inputPath, options), fieldValues, options);
        WriteOutput(outputStream, bytes, options);
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
        return FillAndFlattenFieldsToBytes(inputPath, fieldValues, options, readOptions: null);
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
        return FillAndFlattenFields(ReadPath(inputPath, options), fieldValues, options);
    }
}
