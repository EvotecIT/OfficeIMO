namespace OfficeIMO.Pdf;

public static partial class PdfFormFiller {
    /// <summary>
    /// Returns a new PDF with simple AcroForm field values updated by fully qualified field name.
    /// </summary>
    public static byte[] FillFields(byte[] pdf, IReadOnlyDictionary<string, string> fieldValues) {
        return FillFields(pdf, ToFormFieldValues(fieldValues));
    }

    /// <summary>
    /// Returns a new PDF with simple AcroForm field values updated by fully qualified field name.
    /// </summary>
    public static byte[] FillFields(byte[] pdf, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        Guard.NotNull(pdf, nameof(pdf));
        ValidateFieldValues(fieldValues);

        if (PdfSyntax.HasSignatureMarkers(pdf)) {
            throw new NotSupportedException("Signed PDF files are not supported for form filling by OfficeIMO.Pdf yet.");
        }

        if (PdfSyntax.HasActiveContentMarkers(pdf)) {
            throw new NotSupportedException("PDF active content is not supported for form filling by OfficeIMO.Pdf yet.");
        }

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf);
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
        for (int i = 0; i < fields.Items.Count; i++) {
            FillField(objects, fields.Items[i], null, null, 0, null, fieldValues, remaining, new HashSet<int>(), ref nextObjectNumber);
        }

        if (remaining.Count > 0) {
            throw new ArgumentException("PDF form field was not found: " + string.Join(", ", remaining), nameof(fieldValues));
        }

        acroForm.Items["NeedAppearances"] = new PdfBoolean(true);
        return RewriteAllObjects(objects, catalogObjectNumber, PdfReadDocument.Load(pdf).Metadata);
    }

    /// <summary>
    /// Returns a new PDF with simple AcroForm field values updated from the current position of a readable stream.
    /// </summary>
    public static byte[] FillFields(Stream stream, IReadOnlyDictionary<string, string> fieldValues) {
        return FillFields(ReadStream(stream, nameof(stream)), fieldValues);
    }

    /// <summary>
    /// Returns a new PDF with simple AcroForm field values updated from the current position of a readable stream.
    /// </summary>
    public static byte[] FillFields(Stream stream, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        return FillFields(ReadStream(stream, nameof(stream)), fieldValues);
    }

    /// <summary>
    /// Writes a new PDF with simple AcroForm field values updated.
    /// </summary>
    public static void FillFields(byte[] pdf, Stream outputStream, IReadOnlyDictionary<string, string> fieldValues) {
        WriteOutput(outputStream, FillFields(pdf, fieldValues));
    }

    /// <summary>
    /// Writes a new PDF with simple AcroForm field values updated.
    /// </summary>
    public static void FillFields(byte[] pdf, Stream outputStream, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        WriteOutput(outputStream, FillFields(pdf, fieldValues));
    }

    /// <summary>
    /// Writes a new PDF with simple AcroForm field values updated from the current position of a readable stream.
    /// </summary>
    public static void FillFields(Stream inputStream, Stream outputStream, IReadOnlyDictionary<string, string> fieldValues) {
        WriteOutput(outputStream, FillFields(inputStream, fieldValues));
    }

    /// <summary>
    /// Writes a new PDF with simple AcroForm field values updated from the current position of a readable stream.
    /// </summary>
    public static void FillFields(Stream inputStream, Stream outputStream, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        WriteOutput(outputStream, FillFields(inputStream, fieldValues));
    }

    /// <summary>
    /// Writes a new PDF file with simple AcroForm field values updated.
    /// </summary>
    public static void FillFields(string inputPath, string outputPath, IReadOnlyDictionary<string, string> fieldValues) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        string fullOutputPath = ValidateOutputPath(outputPath);
        byte[] bytes = FillFields(File.ReadAllBytes(inputPath), fieldValues);
        var directory = Path.GetDirectoryName(fullOutputPath);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        File.WriteAllBytes(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF file with simple AcroForm field values updated.
    /// </summary>
    public static void FillFields(string inputPath, string outputPath, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        string fullOutputPath = ValidateOutputPath(outputPath);
        byte[] bytes = FillFields(File.ReadAllBytes(inputPath), fieldValues);
        var directory = Path.GetDirectoryName(fullOutputPath);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        File.WriteAllBytes(fullOutputPath, bytes);
    }

    /// <summary>
    /// Reads a PDF file and writes a new PDF with simple AcroForm field values updated.
    /// </summary>
    public static void FillFields(string inputPath, Stream outputStream, IReadOnlyDictionary<string, string> fieldValues) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        byte[] bytes = FillFields(File.ReadAllBytes(inputPath), fieldValues);
        WriteOutput(outputStream, bytes);
    }

    /// <summary>
    /// Reads a PDF file and writes a new PDF with simple AcroForm field values updated.
    /// </summary>
    public static void FillFields(string inputPath, Stream outputStream, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        byte[] bytes = FillFields(File.ReadAllBytes(inputPath), fieldValues);
        WriteOutput(outputStream, bytes);
    }

    /// <summary>
    /// Reads a PDF file and returns new PDF bytes with simple AcroForm field values updated.
    /// </summary>
    public static byte[] FillFieldsToBytes(string inputPath, IReadOnlyDictionary<string, string> fieldValues) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return FillFields(File.ReadAllBytes(inputPath), fieldValues);
    }

    /// <summary>
    /// Reads a PDF file and returns new PDF bytes with simple AcroForm field values updated.
    /// </summary>
    public static byte[] FillFieldsToBytes(string inputPath, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return FillFields(File.ReadAllBytes(inputPath), fieldValues);
    }

    /// <summary>
    /// Returns a new PDF with simple text, choice, and button AcroForm widgets painted into page content and removed from the form tree.
    /// </summary>
    public static byte[] FlattenFields(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        if (PdfSyntax.HasSignatureMarkers(pdf)) {
            throw new NotSupportedException("Signed PDF files are not supported for form flattening by OfficeIMO.Pdf yet.");
        }

        if (PdfSyntax.HasActiveContentMarkers(pdf)) {
            throw new NotSupportedException("PDF active content is not supported for form flattening by OfficeIMO.Pdf yet.");
        }

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf);
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
        var widgets = new Dictionary<int, FlattenWidgetState>();
        var removableObjects = new HashSet<int>();
        for (int i = 0; i < fields.Items.Count; i++) {
            CollectFlattenWidgets(objects, fields.Items[i], null, null, null, null, widgets, removableObjects, new HashSet<int>(), ref nextObjectNumber);
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

        return RewriteAllObjects(objects, catalogObjectNumber, PdfReadDocument.Load(pdf).Metadata);
    }

    /// <summary>
    /// Returns a new PDF with simple text, choice, and button AcroForm widgets flattened from the current position of a readable stream.
    /// </summary>
    public static byte[] FlattenFields(Stream stream) {
        return FlattenFields(ReadStream(stream, nameof(stream)));
    }

    /// <summary>
    /// Writes a new PDF with simple text, choice, and button AcroForm widgets flattened.
    /// </summary>
    public static void FlattenFields(byte[] pdf, Stream outputStream) {
        WriteOutput(outputStream, FlattenFields(pdf));
    }

    /// <summary>
    /// Writes a new PDF with simple text, choice, and button AcroForm widgets flattened from the current position of a readable stream.
    /// </summary>
    public static void FlattenFields(Stream inputStream, Stream outputStream) {
        WriteOutput(outputStream, FlattenFields(inputStream));
    }

    /// <summary>
    /// Writes a new PDF file with simple text, choice, and button AcroForm widgets flattened.
    /// </summary>
    public static void FlattenFields(string inputPath, string outputPath) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        string fullOutputPath = ValidateOutputPath(outputPath);
        byte[] bytes = FlattenFields(File.ReadAllBytes(inputPath));
        var directory = Path.GetDirectoryName(fullOutputPath);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        File.WriteAllBytes(fullOutputPath, bytes);
    }

    /// <summary>
    /// Reads a PDF file and writes a new PDF with simple text, choice, and button AcroForm widgets flattened.
    /// </summary>
    public static void FlattenFields(string inputPath, Stream outputStream) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        byte[] bytes = FlattenFields(File.ReadAllBytes(inputPath));
        WriteOutput(outputStream, bytes);
    }

    /// <summary>
    /// Reads a PDF file and returns new PDF bytes with simple text, choice, and button AcroForm widgets flattened.
    /// </summary>
    public static byte[] FlattenFieldsToBytes(string inputPath) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return FlattenFields(File.ReadAllBytes(inputPath));
    }

    /// <summary>
    /// Returns a new PDF with simple AcroForm field values updated and then flattened into page content.
    /// </summary>
    public static byte[] FillAndFlattenFields(byte[] pdf, IReadOnlyDictionary<string, string> fieldValues) {
        return FlattenFields(FillFields(pdf, fieldValues));
    }

    /// <summary>
    /// Returns a new PDF with simple AcroForm field values updated and then flattened into page content.
    /// </summary>
    public static byte[] FillAndFlattenFields(byte[] pdf, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        return FlattenFields(FillFields(pdf, fieldValues));
    }

    /// <summary>
    /// Returns a new PDF with simple AcroForm field values updated and flattened from the current position of a readable stream.
    /// </summary>
    public static byte[] FillAndFlattenFields(Stream stream, IReadOnlyDictionary<string, string> fieldValues) {
        return FillAndFlattenFields(ReadStream(stream, nameof(stream)), fieldValues);
    }

    /// <summary>
    /// Returns a new PDF with simple AcroForm field values updated and flattened from the current position of a readable stream.
    /// </summary>
    public static byte[] FillAndFlattenFields(Stream stream, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        return FillAndFlattenFields(ReadStream(stream, nameof(stream)), fieldValues);
    }

    /// <summary>
    /// Writes a new PDF with simple AcroForm field values updated and flattened.
    /// </summary>
    public static void FillAndFlattenFields(byte[] pdf, Stream outputStream, IReadOnlyDictionary<string, string> fieldValues) {
        WriteOutput(outputStream, FillAndFlattenFields(pdf, fieldValues));
    }

    /// <summary>
    /// Writes a new PDF with simple AcroForm field values updated and flattened.
    /// </summary>
    public static void FillAndFlattenFields(byte[] pdf, Stream outputStream, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        WriteOutput(outputStream, FillAndFlattenFields(pdf, fieldValues));
    }

    /// <summary>
    /// Writes a new PDF with simple AcroForm field values updated and flattened from the current position of a readable stream.
    /// </summary>
    public static void FillAndFlattenFields(Stream inputStream, Stream outputStream, IReadOnlyDictionary<string, string> fieldValues) {
        WriteOutput(outputStream, FillAndFlattenFields(inputStream, fieldValues));
    }

    /// <summary>
    /// Writes a new PDF with simple AcroForm field values updated and flattened from the current position of a readable stream.
    /// </summary>
    public static void FillAndFlattenFields(Stream inputStream, Stream outputStream, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        WriteOutput(outputStream, FillAndFlattenFields(inputStream, fieldValues));
    }

    /// <summary>
    /// Writes a new PDF file with simple AcroForm field values updated and flattened.
    /// </summary>
    public static void FillAndFlattenFields(string inputPath, string outputPath, IReadOnlyDictionary<string, string> fieldValues) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        string fullOutputPath = ValidateOutputPath(outputPath);
        byte[] bytes = FillAndFlattenFields(File.ReadAllBytes(inputPath), fieldValues);
        var directory = Path.GetDirectoryName(fullOutputPath);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        File.WriteAllBytes(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF file with simple AcroForm field values updated and flattened.
    /// </summary>
    public static void FillAndFlattenFields(string inputPath, string outputPath, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        string fullOutputPath = ValidateOutputPath(outputPath);
        byte[] bytes = FillAndFlattenFields(File.ReadAllBytes(inputPath), fieldValues);
        var directory = Path.GetDirectoryName(fullOutputPath);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        File.WriteAllBytes(fullOutputPath, bytes);
    }

    /// <summary>
    /// Reads a PDF file and writes a new PDF with simple AcroForm field values updated and flattened.
    /// </summary>
    public static void FillAndFlattenFields(string inputPath, Stream outputStream, IReadOnlyDictionary<string, string> fieldValues) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        byte[] bytes = FillAndFlattenFields(File.ReadAllBytes(inputPath), fieldValues);
        WriteOutput(outputStream, bytes);
    }

    /// <summary>
    /// Reads a PDF file and writes a new PDF with simple AcroForm field values updated and flattened.
    /// </summary>
    public static void FillAndFlattenFields(string inputPath, Stream outputStream, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        byte[] bytes = FillAndFlattenFields(File.ReadAllBytes(inputPath), fieldValues);
        WriteOutput(outputStream, bytes);
    }

    /// <summary>
    /// Reads a PDF file and returns new PDF bytes with simple AcroForm field values updated and flattened.
    /// </summary>
    public static byte[] FillAndFlattenFieldsToBytes(string inputPath, IReadOnlyDictionary<string, string> fieldValues) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return FillAndFlattenFields(File.ReadAllBytes(inputPath), fieldValues);
    }

    /// <summary>
    /// Reads a PDF file and returns new PDF bytes with simple AcroForm field values updated and flattened.
    /// </summary>
    public static byte[] FillAndFlattenFieldsToBytes(string inputPath, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return FillAndFlattenFields(File.ReadAllBytes(inputPath), fieldValues);
    }
}
