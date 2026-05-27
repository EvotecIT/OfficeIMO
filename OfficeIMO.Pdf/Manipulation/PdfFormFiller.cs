namespace OfficeIMO.Pdf;

/// <summary>
/// Updates simple AcroForm field values in parser-supported PDFs.
/// </summary>
public static class PdfFormFiller {
    private const string UnsupportedFlattenWidgetMessage = "Only simple text, choice, and button AcroForm widgets with rectangles are supported for flattening by OfficeIMO.Pdf yet.";
    private const string UnsupportedFlattenAnnotationMessage = "Only simple text, choice, and button AcroForm widgets referenced from page annotations are supported for flattening by OfficeIMO.Pdf yet.";

    private readonly struct ChoiceFillValue {
        public string ExportValue { get; }
        public string DisplayValue { get; }

        public ChoiceFillValue(string exportValue, string displayValue) {
            ExportValue = exportValue;
            DisplayValue = displayValue;
        }
    }

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
            FillField(objects, fields.Items[i], null, null, fieldValues, remaining, new HashSet<int>(), ref nextObjectNumber);
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
            CollectFlattenWidgets(objects, fields.Items[i], null, null, null, widgets, removableObjects, new HashSet<int>(), ref nextObjectNumber);
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

    private static Dictionary<string, PdfFormFieldValue> ToFormFieldValues(IReadOnlyDictionary<string, string> fieldValues) {
        Guard.NotNull(fieldValues, nameof(fieldValues));

        var converted = new Dictionary<string, PdfFormFieldValue>(fieldValues.Count, StringComparer.Ordinal);
        foreach (var pair in fieldValues) {
            if (pair.Value is null) {
                throw new ArgumentException("Field values cannot be null. Use an empty string to clear a text value.", nameof(fieldValues));
            }

            converted[pair.Key] = PdfFormFieldValue.From(pair.Value);
        }

        return converted;
    }

    private static void ValidateFieldValues(IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        Guard.NotNull(fieldValues, nameof(fieldValues));
        if (fieldValues.Count == 0) {
            throw new ArgumentException("At least one field value is required.", nameof(fieldValues));
        }

        foreach (var pair in fieldValues) {
            if (string.IsNullOrWhiteSpace(pair.Key)) {
                throw new ArgumentException("Field names cannot be empty or whitespace.", nameof(fieldValues));
            }

            if (pair.Value is null) {
                throw new ArgumentException("Field values cannot be null. Use an empty string to clear a text value.", nameof(fieldValues));
            }

            if (pair.Value.Values.Count == 0) {
                throw new ArgumentException("Field values must contain at least one entry. Use an empty string to clear a text value.", nameof(fieldValues));
            }

            for (int i = 0; i < pair.Value.Values.Count; i++) {
                if (pair.Value.Values[i] is null) {
                    throw new ArgumentException("Field values cannot contain null entries. Use an empty string to clear a text value.", nameof(fieldValues));
                }
            }
        }
    }

    private static void CollectFlattenWidgets(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject fieldObject,
        string? inheritedFieldType,
        string? inheritedDisplayValue,
        string? inheritedName,
        Dictionary<int, FlattenWidgetState> widgets,
        HashSet<int> removableObjects,
        HashSet<int> visited,
        ref int nextObjectNumber) {
        int? fieldObjectNumber = null;
        if (fieldObject is PdfReference reference) {
            fieldObjectNumber = reference.ObjectNumber;
            if (!visited.Add(reference.ObjectNumber)) {
                return;
            }
        }

        if (ResolveObject(objects, fieldObject) is not PdfDictionary field) {
            return;
        }

        if (fieldObjectNumber.HasValue) {
            removableObjects.Add(fieldObjectNumber.Value);
        }

        string? partialName = TryReadText(objects, field, "T");
        string? fullName = CombineFieldName(inheritedName, partialName);
        string? fieldType = TryReadName(objects, field, "FT") ?? inheritedFieldType;
        IReadOnlyList<string>? values = TryReadSimpleValues(objects, field, "V");
        string? value = values is { Count: > 0 } ? values[0] : inheritedDisplayValue;
        bool isButtonField = string.Equals(fieldType, "Btn", StringComparison.Ordinal);
        string? appearanceValue = string.Equals(fieldType, "Ch", StringComparison.Ordinal)
            ? TryResolveChoiceDisplayValue(objects, field, values) ?? JoinSimpleValues(values) ?? inheritedDisplayValue
            : value;

        if (IsWidget(field)) {
            if (!fieldObjectNumber.HasValue ||
                !TryReadRectCoordinates(field, out double x, out double y, out double width, out double height)) {
                throw new NotSupportedException(UnsupportedFlattenWidgetMessage);
            }

            int appearanceObjectNumber;
            if (isButtonField) {
                if (!TryGetButtonAppearanceReference(objects, field, value, out PdfReference? appearanceReference)) {
                    EnsureButtonWidgetAppearances(objects, field, value ?? "Off", ref nextObjectNumber);
                    if (!TryGetButtonAppearanceReference(objects, field, value, out appearanceReference)) {
                        throw new NotSupportedException(UnsupportedFlattenWidgetMessage);
                    }
                }

                appearanceObjectNumber = appearanceReference!.ObjectNumber;
            } else if (TryGetNormalAppearanceReference(objects, field, out PdfReference? appearanceReference)) {
                appearanceObjectNumber = appearanceReference!.ObjectNumber;
            } else {
                appearanceObjectNumber = nextObjectNumber++;
                objects[appearanceObjectNumber] = new PdfIndirectObject(appearanceObjectNumber, 0, CreateTextAppearanceStream(appearanceValue ?? string.Empty, width, height));
            }

            widgets[fieldObjectNumber.Value] = new FlattenWidgetState(fieldObjectNumber.Value, x, y, width, height, appearanceObjectNumber);
            return;
        }

        if (!field.Items.TryGetValue("Kids", out var kidsObject) ||
            ResolveObject(objects, kidsObject) is not PdfArray kids) {
            throw new NotSupportedException(UnsupportedFlattenWidgetMessage);
        }

        for (int i = 0; i < kids.Items.Count; i++) {
            CollectFlattenWidgets(objects, kids.Items[i], fieldType, appearanceValue, fullName, widgets, removableObjects, visited, ref nextObjectNumber);
        }
    }

    private static int FlattenPageWidgets(Dictionary<int, PdfIndirectObject> objects, Dictionary<int, FlattenWidgetState> widgets, ref int nextObjectNumber) {
        int flattenedWidgetCount = 0;
        foreach (var entry in objects.OrderBy(pair => pair.Key).ToArray()) {
            if (entry.Value.Value is not PdfDictionary page ||
                page.Get<PdfName>("Type")?.Name != "Page" ||
                !page.Items.TryGetValue("Annots", out var annotsObject) ||
                ResolveObject(objects, annotsObject) is not PdfArray annots) {
                continue;
            }

            var pageWidgets = new List<FlattenWidgetState>();
            var remainingAnnots = new PdfArray();
            for (int i = 0; i < annots.Items.Count; i++) {
                PdfObject annot = annots.Items[i];
                if (annot is PdfReference annotReference && widgets.TryGetValue(annotReference.ObjectNumber, out var widget)) {
                    pageWidgets.Add(widget);
                    flattenedWidgetCount++;
                } else {
                    remainingAnnots.Items.Add(annot);
                }
            }

            if (pageWidgets.Count == 0) {
                continue;
            }

            if (remainingAnnots.Items.Count == 0) {
                page.Items.Remove("Annots");
            } else {
                page.Items["Annots"] = remainingAnnots;
            }

            string content = BuildFlattenContent(objects, page, pageWidgets);
            int contentObjectNumber = nextObjectNumber++;
            objects[contentObjectNumber] = new PdfIndirectObject(contentObjectNumber, 0, CreateContentStream(content));
            AppendPageContent(objects, page, contentObjectNumber);
        }

        return flattenedWidgetCount;
    }

    private static string BuildFlattenContent(Dictionary<int, PdfIndirectObject> objects, PdfDictionary page, List<FlattenWidgetState> widgets) {
        PdfDictionary xObjects = EnsurePageXObjects(objects, page);
        var builder = new StringBuilder();
        for (int i = 0; i < widgets.Count; i++) {
            FlattenWidgetState widget = widgets[i];
            string xObjectName = CreateUniqueXObjectName(xObjects);
            xObjects.Items[xObjectName] = new PdfReference(widget.AppearanceObjectNumber, 0);
            builder.Append("q\n");
            builder.Append(FormatNumber(widget.Width)).Append(" 0 0 ").Append(FormatNumber(widget.Height)).Append(' ')
                .Append(FormatNumber(widget.X)).Append(' ').Append(FormatNumber(widget.Y)).Append(" cm\n");
            builder.Append('/').Append(xObjectName).Append(" Do\n");
            builder.Append("Q\n");
        }

        return builder.ToString();
    }

    private static PdfDictionary EnsurePageXObjects(Dictionary<int, PdfIndirectObject> objects, PdfDictionary page) {
        PdfDictionary resources;
        if (page.Items.TryGetValue("Resources", out var resourcesObject)) {
            resources = ResolveDictionary(objects, resourcesObject) ?? throw new NotSupportedException("Indirect page resources must resolve to a dictionary before form flattening.");
        } else {
            resources = new PdfDictionary();
            page.Items["Resources"] = resources;
        }

        if (resources.Items.TryGetValue("XObject", out var xObjectObject)) {
            if (ResolveObject(objects, xObjectObject) is PdfDictionary existing) {
                return existing;
            }

            throw new NotSupportedException("Page XObject resources must be a dictionary before form flattening.");
        }

        var xObjects = new PdfDictionary();
        resources.Items["XObject"] = xObjects;
        return xObjects;
    }

    private static string CreateUniqueXObjectName(PdfDictionary xObjects) {
        int index = 1;
        string name;
        do {
            name = "OfficeIMOForm" + index.ToString(System.Globalization.CultureInfo.InvariantCulture);
            index++;
        } while (xObjects.Items.ContainsKey(name));

        return name;
    }

    private static PdfStream CreateContentStream(string content) {
        var dictionary = new PdfDictionary();
        return new PdfStream(dictionary, PdfEncoding.Latin1GetBytes(content));
    }

    private static void AppendPageContent(Dictionary<int, PdfIndirectObject> objects, PdfDictionary page, int contentObjectNumber) {
        var newReference = new PdfReference(contentObjectNumber, 0);
        if (!page.Items.TryGetValue("Contents", out var contents)) {
            page.Items["Contents"] = newReference;
            return;
        }

        if (contents is PdfArray contentsArray) {
            contentsArray.Items.Add(newReference);
            return;
        }

        var array = new PdfArray();
        AppendContentEntries(objects, array, contents);
        array.Items.Add(newReference);
        page.Items["Contents"] = array;
    }

    private static void AppendContentEntries(Dictionary<int, PdfIndirectObject> objects, PdfArray target, PdfObject contents) {
        if (contents is PdfArray directArray) {
            foreach (var item in directArray.Items) {
                target.Items.Add(item);
            }

            return;
        }

        if (contents is PdfReference reference &&
            PdfObjectLookup.TryGet(objects, reference, out var indirect) &&
            indirect.Value is PdfArray referencedArray) {
            foreach (var item in referencedArray.Items) {
                target.Items.Add(item);
            }

            return;
        }

        target.Items.Add(contents);
    }

    private static void FillField(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject fieldObject,
        string? parentName,
        string? inheritedFieldType,
        IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues,
        HashSet<string> remaining,
        HashSet<int> visited,
        ref int nextObjectNumber) {
        if (fieldObject is PdfReference reference && !visited.Add(reference.ObjectNumber)) {
            return;
        }

        if (ResolveObject(objects, fieldObject) is not PdfDictionary field) {
            return;
        }

        string? partialName = TryReadText(objects, field, "T");
        string? fullName = CombineFieldName(parentName, partialName);
        string? fieldType = TryReadName(objects, field, "FT") ?? inheritedFieldType;
        if (fullName is not null && remaining.Contains(fullName) && fieldValues.TryGetValue(fullName, out PdfFormFieldValue? value)) {
            SetFieldValue(objects, field, fieldType, value, ref nextObjectNumber);
            remaining.Remove(fullName);
        }

        if (!field.Items.TryGetValue("Kids", out var kidsObject) ||
            ResolveObject(objects, kidsObject) is not PdfArray kids) {
            return;
        }

        for (int i = 0; i < kids.Items.Count; i++) {
            FillField(objects, kids.Items[i], fullName, fieldType, fieldValues, remaining, visited, ref nextObjectNumber);
        }
    }

    private static void SetFieldValue(Dictionary<int, PdfIndirectObject> objects, PdfDictionary field, string? fieldType, PdfFormFieldValue value, ref int nextObjectNumber) {
        IReadOnlyList<string> values = value.Values;
        string firstValue = values[0];
        if (string.Equals(fieldType, "Btn", StringComparison.Ordinal)) {
            string name = string.IsNullOrEmpty(firstValue) ? "Off" : firstValue;
            field.Items["V"] = new PdfName(name);
            field.Items["AS"] = new PdfName(name);
            SetWidgetAppearanceStates(objects, field, name, new HashSet<int>(), ref nextObjectNumber);
            return;
        }

        if (string.Equals(fieldType, "Ch", StringComparison.Ordinal)) {
            IReadOnlyList<ChoiceFillValue> choiceValues = ResolveChoiceFillValues(objects, field, values);
            if (choiceValues.Count > 1) {
                field.Items["V"] = CreateStringArray(choiceValues.Select(item => item.ExportValue));
                SetTextWidgetAppearances(objects, field, string.Join(", ", choiceValues.Select(item => item.DisplayValue)), new HashSet<int>(), ref nextObjectNumber);
                return;
            }

            ChoiceFillValue choiceValue = choiceValues[0];
            field.Items["V"] = new PdfStringObj(choiceValue.ExportValue);
            SetTextWidgetAppearances(objects, field, choiceValue.DisplayValue, new HashSet<int>(), ref nextObjectNumber);
            return;
        }

        field.Items["V"] = new PdfStringObj(firstValue);
        SetTextWidgetAppearances(objects, field, firstValue, new HashSet<int>(), ref nextObjectNumber);
    }

    private static void SetWidgetAppearanceStates(Dictionary<int, PdfIndirectObject> objects, PdfDictionary field, string name, HashSet<int> visited, ref int nextObjectNumber) {
        if (IsWidget(field)) {
            field.Items["AS"] = new PdfName(name);
            EnsureButtonWidgetAppearances(objects, field, name, ref nextObjectNumber);
        }

        if (!field.Items.TryGetValue("Kids", out var kidsObject) ||
            ResolveObject(objects, kidsObject) is not PdfArray kids) {
            return;
        }

        for (int i = 0; i < kids.Items.Count; i++) {
            PdfObject kidObject = kids.Items[i];
            if (kidObject is PdfReference reference && !visited.Add(reference.ObjectNumber)) {
                continue;
            }

            if (ResolveObject(objects, kidObject) is PdfDictionary kid) {
                kid.Items["AS"] = new PdfName(name);
                SetWidgetAppearanceStates(objects, kid, name, visited, ref nextObjectNumber);
            }
        }
    }

    private static void EnsureButtonWidgetAppearances(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget, string selectedName, ref int nextObjectNumber) {
        if (!TryReadRect(widget, out double width, out double height)) {
            return;
        }

        PdfDictionary normalAppearances = GetOrCreateButtonNormalAppearanceDictionary(objects, widget);
        if (!normalAppearances.Items.ContainsKey("Off")) {
            int offAppearanceObjectNumber = nextObjectNumber++;
            objects[offAppearanceObjectNumber] = new PdfIndirectObject(offAppearanceObjectNumber, 0, CreateButtonAppearanceStream(width, height, selected: false));
            normalAppearances.Items["Off"] = new PdfReference(offAppearanceObjectNumber, 0);
        }

        if (!string.Equals(selectedName, "Off", StringComparison.Ordinal) && !normalAppearances.Items.ContainsKey(selectedName)) {
            int selectedAppearanceObjectNumber = nextObjectNumber++;
            objects[selectedAppearanceObjectNumber] = new PdfIndirectObject(selectedAppearanceObjectNumber, 0, CreateButtonAppearanceStream(width, height, selected: true));
            normalAppearances.Items[selectedName] = new PdfReference(selectedAppearanceObjectNumber, 0);
        }
    }

    private static PdfDictionary GetOrCreateButtonNormalAppearanceDictionary(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget) {
        PdfDictionary appearance;
        if (widget.Items.TryGetValue("AP", out var appearanceObject) &&
            ResolveDictionary(objects, appearanceObject) is PdfDictionary existingAppearance) {
            appearance = existingAppearance;
        } else {
            appearance = new PdfDictionary();
            widget.Items["AP"] = appearance;
        }

        if (appearance.Items.TryGetValue("N", out var normalAppearanceObject) &&
            ResolveDictionary(objects, normalAppearanceObject) is PdfDictionary existingNormalAppearance) {
            return existingNormalAppearance;
        }

        var normalAppearances = new PdfDictionary();
        appearance.Items["N"] = normalAppearances;
        return normalAppearances;
    }

    private static void SetTextWidgetAppearances(Dictionary<int, PdfIndirectObject> objects, PdfDictionary field, string value, HashSet<int> visited, ref int nextObjectNumber) {
        if (IsWidget(field) && TryReadRect(field, out double width, out double height)) {
            int appearanceObjectNumber = nextObjectNumber++;
            objects[appearanceObjectNumber] = new PdfIndirectObject(appearanceObjectNumber, 0, CreateTextAppearanceStream(value, width, height));

            var appearance = new PdfDictionary();
            appearance.Items["N"] = new PdfReference(appearanceObjectNumber, 0);
            field.Items["AP"] = appearance;
        }

        if (!field.Items.TryGetValue("Kids", out var kidsObject) ||
            ResolveObject(objects, kidsObject) is not PdfArray kids) {
            return;
        }

        for (int i = 0; i < kids.Items.Count; i++) {
            PdfObject kidObject = kids.Items[i];
            if (kidObject is PdfReference reference && !visited.Add(reference.ObjectNumber)) {
                continue;
            }

            if (ResolveObject(objects, kidObject) is PdfDictionary kid) {
                SetTextWidgetAppearances(objects, kid, value, visited, ref nextObjectNumber);
            }
        }
    }

    private static PdfStream CreateTextAppearanceStream(string value, double width, double height) {
        double fontSize = Math.Max(6D, Math.Min(12D, height - 4D));
        double baseline = Math.Max(2D, (height - fontSize) / 2D);
        string escapedValue = PdfSyntaxEscaper.EscapeLiteralContent(value);
        string content =
            "q\n" +
            "1 1 1 rg 0 0 " + FormatNumber(width) + " " + FormatNumber(height) + " re f\n" +
            "0.75 0.75 0.75 RG 0.5 0.5 " + FormatNumber(Math.Max(0D, width - 1D)) + " " + FormatNumber(Math.Max(0D, height - 1D)) + " re S\n" +
            "BT /Helv " + FormatNumber(fontSize) + " Tf 0 0 0 rg 2 " + FormatNumber(baseline) + " Td (" + escapedValue + ") Tj ET\n" +
            "Q\n";

        var dictionary = new PdfDictionary();
        dictionary.Items["Type"] = new PdfName("XObject");
        dictionary.Items["Subtype"] = new PdfName("Form");
        dictionary.Items["BBox"] = CreateNumberArray(0D, 0D, width, height);
        dictionary.Items["Resources"] = CreateAppearanceResources();
        return new PdfStream(dictionary, PdfEncoding.Latin1GetBytes(content));
    }

    private static PdfStream CreateButtonAppearanceStream(double width, double height, bool selected) {
        double boxWidth = Math.Max(0D, width - 1D);
        double boxHeight = Math.Max(0D, height - 1D);
        string content =
            "q\n" +
            "1 1 1 rg 0 0 " + FormatNumber(width) + " " + FormatNumber(height) + " re f\n" +
            "0.75 0.75 0.75 RG 0.5 0.5 " + FormatNumber(boxWidth) + " " + FormatNumber(boxHeight) + " re S\n";

        if (selected) {
            double markLeft = Math.Max(2D, width * 0.2D);
            double markMidX = Math.Max(markLeft + 1D, width * 0.42D);
            double markRight = Math.Max(markMidX + 1D, width * 0.8D);
            double markMidY = Math.Max(2D, height * 0.25D);
            double markLeftY = Math.Min(height - 2D, height * 0.52D);
            double markRightY = Math.Min(height - 2D, height * 0.78D);
            content +=
                "0 0 0 RG 1.25 w " +
                FormatNumber(markLeft) + " " + FormatNumber(markLeftY) + " m " +
                FormatNumber(markMidX) + " " + FormatNumber(markMidY) + " l " +
                FormatNumber(markRight) + " " + FormatNumber(markRightY) + " l S\n";
        }

        content += "Q\n";

        var dictionary = new PdfDictionary();
        dictionary.Items["Type"] = new PdfName("XObject");
        dictionary.Items["Subtype"] = new PdfName("Form");
        dictionary.Items["BBox"] = CreateNumberArray(0D, 0D, width, height);
        return new PdfStream(dictionary, PdfEncoding.Latin1GetBytes(content));
    }

    private static PdfDictionary CreateAppearanceResources() {
        var font = new PdfDictionary();
        font.Items["Type"] = new PdfName("Font");
        font.Items["Subtype"] = new PdfName("Type1");
        font.Items["BaseFont"] = new PdfName("Helvetica");

        var fonts = new PdfDictionary();
        fonts.Items["Helv"] = font;

        var resources = new PdfDictionary();
        resources.Items["Font"] = fonts;
        return resources;
    }

    private static PdfArray CreateNumberArray(params double[] values) {
        var array = new PdfArray();
        foreach (double value in values) {
            array.Items.Add(new PdfNumber(value));
        }

        return array;
    }

    private static PdfArray CreateStringArray(IEnumerable<string> values) {
        var array = new PdfArray();
        foreach (string value in values) {
            array.Items.Add(new PdfStringObj(value));
        }

        return array;
    }

    private static bool IsWidget(PdfDictionary dictionary) {
        return dictionary.Items.TryGetValue("Subtype", out var subtype) &&
            subtype is PdfName name &&
            string.Equals(name.Name, "Widget", StringComparison.Ordinal);
    }

    private static bool TryReadRect(PdfDictionary dictionary, out double width, out double height) {
        if (TryReadRectCoordinates(dictionary, out _, out _, out width, out height)) {
            return true;
        }

        width = 0D;
        height = 0D;
        return false;
    }

    private static bool TryReadRectCoordinates(PdfDictionary dictionary, out double x, out double y, out double width, out double height) {
        x = 0D;
        y = 0D;
        width = 0D;
        height = 0D;
        if (!dictionary.Items.TryGetValue("Rect", out var rectObject) ||
            rectObject is not PdfArray rect ||
            rect.Items.Count < 4 ||
            rect.Items[0] is not PdfNumber x1 ||
            rect.Items[1] is not PdfNumber y1 ||
            rect.Items[2] is not PdfNumber x2 ||
            rect.Items[3] is not PdfNumber y2) {
            return false;
        }

        x = Math.Min(x1.Value, x2.Value);
        y = Math.Min(y1.Value, y2.Value);
        width = Math.Abs(x2.Value - x1.Value);
        height = Math.Abs(y2.Value - y1.Value);
        return width > 0D && height > 0D;
    }

    private static bool TryGetNormalAppearanceReference(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget, out PdfReference? reference) {
        reference = null;
        if (!TryGetNormalAppearanceObject(objects, widget, out PdfObject? normalAppearance) ||
            normalAppearance is not PdfReference normalAppearanceReference) {
            return false;
        }

        reference = normalAppearanceReference;
        return true;
    }

    private static bool TryGetButtonAppearanceReference(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget, string? inheritedValue, out PdfReference? reference) {
        reference = null;
        if (!TryGetNormalAppearanceObject(objects, widget, out PdfObject? normalAppearance)) {
            return false;
        }

        if (normalAppearance is PdfReference singleAppearanceReference) {
            reference = singleAppearanceReference;
            return true;
        }

        if (normalAppearance is not PdfDictionary appearanceStates) {
            return false;
        }

        string? stateName = TryReadName(objects, widget, "AS") ?? inheritedValue;
        if (stateName is { Length: > 0 } &&
            TryGetAppearanceStateReference(appearanceStates, stateName, out reference)) {
            return true;
        }

        return TryGetAppearanceStateReference(appearanceStates, "Off", out reference);
    }

    private static bool TryGetNormalAppearanceObject(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget, out PdfObject? normalAppearance) {
        normalAppearance = null;
        if (!widget.Items.TryGetValue("AP", out var appearanceObject) ||
            ResolveDictionary(objects, appearanceObject) is not PdfDictionary appearance ||
            !appearance.Items.TryGetValue("N", out var normalAppearanceObject)) {
            return false;
        }

        if (normalAppearanceObject is PdfReference normalAppearanceReference) {
            PdfObject? resolved = ResolveObject(objects, normalAppearanceReference);
            normalAppearance = resolved is PdfStream ? normalAppearanceReference : resolved;
            return normalAppearance is not null;
        }

        normalAppearance = normalAppearanceObject;
        return true;
    }

    private static bool TryGetAppearanceStateReference(PdfDictionary appearanceStates, string stateName, out PdfReference? reference) {
        reference = null;
        if (!appearanceStates.Items.TryGetValue(stateName, out var stateAppearance) ||
            stateAppearance is not PdfReference stateAppearanceReference) {
            return false;
        }

        reference = stateAppearanceReference;
        return true;
    }

    private static string? TryReadSimpleValue(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, string key) {
        if (!dictionary.Items.TryGetValue(key, out var value)) {
            return null;
        }

        return ResolveObject(objects, value) switch {
            PdfStringObj text => text.Value,
            PdfName name => name.Name,
            PdfNumber number => FormatNumber(number.Value),
            PdfBoolean boolean => boolean.Value ? "true" : "false",
            _ => null
        };
    }

    private static IReadOnlyList<string>? TryReadSimpleValues(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, string key) {
        if (!dictionary.Items.TryGetValue(key, out var value)) {
            return null;
        }

        PdfObject? resolved = ResolveObject(objects, value);
        if (resolved is PdfArray array) {
            var values = new List<string>();
            for (int i = 0; i < array.Items.Count; i++) {
                if (TryFormatSimpleValue(objects, array.Items[i], out string? item)) {
                    values.Add(item!);
                }
            }

            return values.Count == 0 ? null : values.AsReadOnly();
        }

        return TryFormatSimpleValue(objects, resolved, out string? single)
            ? new[] { single! }
            : null;
    }

    private static bool TryFormatSimpleValue(Dictionary<int, PdfIndirectObject> objects, PdfObject? value, out string? text) {
        text = null;
        switch (ResolveObject(objects, value)) {
            case PdfStringObj stringObj:
                text = stringObj.Value;
                return true;
            case PdfName name:
                text = name.Name;
                return true;
            case PdfNumber number:
                text = FormatNumber(number.Value);
                return true;
            case PdfBoolean boolean:
                text = boolean.Value ? "true" : "false";
                return true;
            default:
                return false;
        }
    }

    private static string? JoinSimpleValues(IReadOnlyList<string>? values) {
        return values is { Count: > 0 }
            ? string.Join(", ", values)
            : null;
    }

    private static string? TryResolveChoiceDisplayValue(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, IReadOnlyList<string>? exportValues) {
        if (exportValues is not { Count: > 0 } ||
            !dictionary.Items.TryGetValue("Opt", out var optionsObject) ||
            ResolveObject(objects, optionsObject) is not PdfArray options ||
            options.Items.Count == 0) {
            return null;
        }

        var displayValues = new List<string>(exportValues.Count);
        for (int i = 0; i < exportValues.Count; i++) {
            displayValues.Add(ResolveSingleChoiceDisplayValue(objects, options, exportValues[i]) ?? exportValues[i]);
        }

        return displayValues.Count == 0 ? null : string.Join(", ", displayValues);
    }

    private static string? ResolveSingleChoiceDisplayValue(Dictionary<int, PdfIndirectObject> objects, PdfArray options, string exportValue) {
        for (int i = 0; i < options.Items.Count; i++) {
            PdfObject? optionObject = ResolveObject(objects, options.Items[i]);
            if (optionObject is PdfArray pair &&
                pair.Items.Count >= 2 &&
                TryReadOptionText(objects, pair.Items[0], out string? pairExportValue) &&
                string.Equals(pairExportValue, exportValue, StringComparison.Ordinal) &&
                TryReadOptionText(objects, pair.Items[1], out string? pairDisplayText)) {
                return pairDisplayText;
            }

            if (optionObject is not null &&
                TryReadOptionText(objects, optionObject, out string? singleValue) &&
                string.Equals(singleValue, exportValue, StringComparison.Ordinal)) {
                return singleValue;
            }
        }

        return null;
    }

    private static List<ChoiceFillValue> ResolveChoiceFillValues(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, IReadOnlyList<string> values) {
        PdfArray? options = null;
        if (dictionary.Items.TryGetValue("Opt", out var optionsObject)) {
            options = ResolveObject(objects, optionsObject) as PdfArray;
        }

        var resolved = new List<ChoiceFillValue>(values.Count);
        for (int i = 0; i < values.Count; i++) {
            resolved.Add(options is null || options.Items.Count == 0
                ? new ChoiceFillValue(values[i], values[i])
                : ResolveSingleChoiceFillValue(objects, options, values[i]));
        }

        return resolved;
    }

    private static ChoiceFillValue ResolveSingleChoiceFillValue(Dictionary<int, PdfIndirectObject> objects, PdfArray options, string value) {
        for (int i = 0; i < options.Items.Count; i++) {
            PdfObject? optionObject = ResolveObject(objects, options.Items[i]);
            if (optionObject is PdfArray pair &&
                pair.Items.Count >= 2 &&
                TryReadOptionText(objects, pair.Items[0], out string? pairExportValue) &&
                pairExportValue is not null &&
                TryReadOptionText(objects, pair.Items[1], out string? pairDisplayText) &&
                pairDisplayText is not null &&
                (string.Equals(pairExportValue, value, StringComparison.Ordinal) ||
                 string.Equals(pairDisplayText, value, StringComparison.Ordinal))) {
                return new ChoiceFillValue(pairExportValue, pairDisplayText);
            }

            if (optionObject is not null &&
                TryReadOptionText(objects, optionObject, out string? singleValue) &&
                singleValue is not null &&
                string.Equals(singleValue, value, StringComparison.Ordinal)) {
                return new ChoiceFillValue(singleValue, singleValue);
            }
        }

        return new ChoiceFillValue(value, value);
    }

    private static bool TryReadOptionText(Dictionary<int, PdfIndirectObject> objects, PdfObject value, out string? text) {
        switch (ResolveObject(objects, value)) {
            case PdfStringObj stringObj:
                text = stringObj.Value;
                return true;
            case PdfName name:
                text = name.Name;
                return true;
            case PdfNumber number:
                text = FormatNumber(number.Value);
                return true;
            default:
                text = null;
                return false;
        }
    }

    private static string FormatNumber(double value) {
        if (Math.Abs(value % 1D) < 0.0000001D) {
            return ((long)Math.Round(value)).ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        return value.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);
    }

    private static byte[] RewriteAllObjects(Dictionary<int, PdfIndirectObject> objects, int catalogObjectNumber, PdfMetadata metadata) {
        var sourceIds = objects.Keys.OrderBy(id => id).ToArray();
        var numberMap = new Dictionary<int, int>(sourceIds.Length);
        for (int i = 0; i < sourceIds.Length; i++) {
            numberMap[sourceIds[i]] = i + 1;
        }

        var context = new PdfPageExtractor.SerializationContext(numberMap, pagesObjectId: 0, new Dictionary<int, Dictionary<string, PdfObject>>(), objects);
        var rewritten = new List<byte[]>(sourceIds.Length + 1);
        foreach (int sourceId in sourceIds) {
            byte[] body = PdfPageExtractor.SerializeObject(objects[sourceId].Value, context);
            rewritten.Add(PdfPageExtractor.WrapObject(numberMap[sourceId], body));
        }

        int infoId = rewritten.Count + 1;
        rewritten.Add(PdfPageExtractor.WrapObject(infoId, PdfEncoding.Latin1GetBytes(PdfPageExtractor.BuildInfoDictionary(metadata))));

        return PdfPageExtractor.Assemble(rewritten, numberMap[catalogObjectNumber], infoId);
    }

    private static int FindCatalogObjectNumber(Dictionary<int, PdfIndirectObject> objects, string? trailerRaw) {
        PdfDictionary? catalog = PdfSyntax.FindCatalog(objects, trailerRaw);
        if (catalog is null) {
            return 0;
        }

        foreach (var entry in objects) {
            if (ReferenceEquals(entry.Value.Value, catalog)) {
                return entry.Key;
            }
        }

        return 0;
    }

    private static PdfObject? ResolveObject(Dictionary<int, PdfIndirectObject> objects, PdfObject? value) {
        return PdfObjectLookup.Resolve(objects, value);
    }

    private static PdfDictionary? ResolveDictionary(Dictionary<int, PdfIndirectObject> objects, PdfObject? value) {
        return ResolveObject(objects, value) as PdfDictionary;
    }

    private static string? TryReadText(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, string key) {
        return dictionary.Items.TryGetValue(key, out var value) &&
            ResolveObject(objects, value) is PdfStringObj text &&
            !string.IsNullOrEmpty(text.Value)
            ? text.Value
            : null;
    }

    private static string? TryReadName(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, string key) {
        return dictionary.Items.TryGetValue(key, out var value) &&
            ResolveObject(objects, value) is PdfName name &&
            !string.IsNullOrEmpty(name.Name)
            ? name.Name
            : null;
    }

    private static string? CombineFieldName(string? parentName, string? partialName) {
        if (string.IsNullOrEmpty(parentName)) {
            return string.IsNullOrEmpty(partialName) ? null : partialName;
        }

        if (string.IsNullOrEmpty(partialName)) {
            return parentName;
        }

        return parentName + "." + partialName;
    }

    private static byte[] ReadStream(Stream stream, string paramName) {
        Guard.NotNull(stream, paramName);
        if (!stream.CanRead) {
            throw new ArgumentException("Stream must be readable.", paramName);
        }

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }

    private static void WriteOutput(Stream outputStream, byte[] bytes) {
        ValidateWritableOutputStream(outputStream);

        outputStream.Write(bytes, 0, bytes.Length);
    }

    private static void ValidateWritableOutputStream(Stream outputStream) {
        Guard.NotNull(outputStream, nameof(outputStream));
        if (!outputStream.CanWrite) {
            throw new ArgumentException("Stream must be writable.", nameof(outputStream));
        }
    }

    private static string ValidateOutputPath(string outputPath) {
        Guard.NotNull(outputPath, nameof(outputPath));
        if (string.IsNullOrWhiteSpace(outputPath)) {
            throw new ArgumentException("Output path cannot be empty or whitespace.", nameof(outputPath));
        }

        string fullPath;
        try {
            fullPath = Path.GetFullPath(outputPath);
        } catch (Exception ex) {
            throw new ArgumentException("Output path is invalid.", nameof(outputPath), ex);
        }

        if (Directory.Exists(fullPath) && (File.GetAttributes(fullPath) & FileAttributes.Directory) == FileAttributes.Directory) {
            throw new ArgumentException("Output path refers to a directory; a file path is required.", nameof(outputPath));
        }

        string fileName = Path.GetFileName(fullPath);
        if (string.IsNullOrEmpty(fileName)) {
            throw new ArgumentException("Output path must include a file name.", nameof(outputPath));
        }

        if (fileName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) {
            throw new ArgumentException("Output path contains invalid file name characters.", nameof(outputPath));
        }

        return fullPath;
    }

    private sealed class FlattenWidgetState {
        internal FlattenWidgetState(int widgetObjectNumber, double x, double y, double width, double height, int appearanceObjectNumber) {
            WidgetObjectNumber = widgetObjectNumber;
            X = x;
            Y = y;
            Width = width;
            Height = height;
            AppearanceObjectNumber = appearanceObjectNumber;
        }

        internal int WidgetObjectNumber { get; }
        internal double X { get; }
        internal double Y { get; }
        internal double Width { get; }
        internal double Height { get; }
        internal int AppearanceObjectNumber { get; }
    }
}
