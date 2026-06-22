using System.Globalization;

namespace OfficeIMO.Pdf;

public static partial class PdfIncrementalUpdater {
    /// <summary>
    /// Appends a simple AcroForm field-value revision to a PDF byte array without rewriting the existing bytes.
    /// </summary>
    public static byte[] UpdateFormFields(byte[] pdf, IReadOnlyDictionary<string, string> fieldValues, bool keepNeedAppearances = true) {
        Guard.NotNull(pdf, nameof(pdf));
        ValidateFieldValues(fieldValues);

        PdfDocumentSecurityInfo security = PdfSyntax.ReadDocumentSecurityInfo(pdf);
        ValidateAppendOnlyFormInput(security);

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf);
        if (!security.RootObjectNumber.HasValue ||
            !objects.TryGetValue(security.RootObjectNumber.Value, out PdfIndirectObject? rootObject) ||
            rootObject.Value is not PdfDictionary catalog ||
            !catalog.Items.TryGetValue("AcroForm", out PdfObject? acroFormObject) ||
            ResolveDictionary(objects, acroFormObject) is not PdfDictionary acroForm ||
            !acroForm.Items.TryGetValue("Fields", out PdfObject? fieldsObject) ||
            ResolveObject(objects, fieldsObject) is not PdfArray fields) {
            throw new ArgumentException("PDF does not contain a readable AcroForm field tree.", nameof(pdf));
        }

        var remaining = new HashSet<string>(fieldValues.Keys, StringComparer.Ordinal);
        var changedObjectNumbers = new HashSet<int>();
        int inheritedFlags = 0;
        for (int i = 0; i < fields.Items.Count; i++) {
            UpdateFormField(objects, fields.Items[i], null, null, inheritedFlags, fieldValues, remaining, changedObjectNumbers, new HashSet<int>());
        }

        if (remaining.Count > 0) {
            throw new ArgumentException("PDF form field was not found: " + string.Join(", ", remaining), nameof(fieldValues));
        }

        acroForm.Items["NeedAppearances"] = new PdfBoolean(keepNeedAppearances);
        if (acroFormObject is PdfReference acroFormReference) {
            changedObjectNumbers.Add(acroFormReference.ObjectNumber);
        } else {
            changedObjectNumbers.Add(security.RootObjectNumber.Value);
        }

        if (changedObjectNumbers.Count == 0) {
            throw new ArgumentException("No supported AcroForm fields were updated.", nameof(fieldValues));
        }

        return AppendIncrementalObjects(pdf, objects, security, trailerRaw, changedObjectNumbers);
    }

    /// <summary>Appends a simple AcroForm field-value revision to a readable PDF stream.</summary>
    public static byte[] UpdateFormFields(Stream input, IReadOnlyDictionary<string, string> fieldValues, bool keepNeedAppearances = true) {
        Guard.NotNull(input, nameof(input));
        if (!input.CanRead) {
            throw new ArgumentException("Stream must be readable.", nameof(input));
        }

        using var buffer = new MemoryStream();
        input.CopyTo(buffer);
        return UpdateFormFields(buffer.ToArray(), fieldValues, keepNeedAppearances);
    }

    /// <summary>Appends a simple AcroForm field-value revision to a PDF file and writes the result to <paramref name="outputPath"/>.</summary>
    public static void UpdateFormFields(string inputPath, string outputPath, IReadOnlyDictionary<string, string> fieldValues, bool keepNeedAppearances = true) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNullOrWhiteSpace(outputPath, nameof(outputPath));
        File.WriteAllBytes(outputPath, UpdateFormFields(File.ReadAllBytes(inputPath), fieldValues, keepNeedAppearances));
    }

    private static void ValidateAppendOnlyFormInput(PdfDocumentSecurityInfo security) {
        PdfAppendOnlyMutationReport report = BuildAppendOnlyMutationReport(security);
        if (!report.SupportedActions.Contains("FormFill", StringComparer.Ordinal)) {
            throw new NotSupportedException("Incremental form field updates are not supported for this PDF: " + string.Join(", ", report.Blockers));
        }
    }

    private static void ValidateFieldValues(IReadOnlyDictionary<string, string> fieldValues) {
        Guard.NotNull(fieldValues, nameof(fieldValues));
        if (fieldValues.Count == 0) {
            throw new ArgumentException("At least one form field value must be provided.", nameof(fieldValues));
        }

        foreach (KeyValuePair<string, string> entry in fieldValues) {
            if (string.IsNullOrWhiteSpace(entry.Key)) {
                throw new ArgumentException("Form field names cannot be empty.", nameof(fieldValues));
            }
        }
    }

    private static void UpdateFormField(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject fieldObject,
        string? parentName,
        string? inheritedFieldType,
        int inheritedFlags,
        IReadOnlyDictionary<string, string> fieldValues,
        HashSet<string> remaining,
        HashSet<int> changedObjectNumbers,
        HashSet<int> visited) {
        int? objectNumber = null;
        if (fieldObject is PdfReference reference) {
            objectNumber = reference.ObjectNumber;
            if (!visited.Add(reference.ObjectNumber)) {
                return;
            }
        }

        if (ResolveObject(objects, fieldObject) is not PdfDictionary field) {
            return;
        }

        string? partialName = TryReadText(objects, field, "T");
        string? fullName = CombineFieldName(parentName, partialName);
        string? fieldType = TryReadName(objects, field, "FT") ?? inheritedFieldType;
        int fieldFlags = ReadFieldFlags(objects, field, inheritedFlags);

        if (fullName is not null && remaining.Contains(fullName) && fieldValues.TryGetValue(fullName, out string? value)) {
            SetIncrementalFieldValue(field, fieldType, value ?? string.Empty);
            if (objectNumber.HasValue) {
                changedObjectNumbers.Add(objectNumber.Value);
            }

            remaining.Remove(fullName);
        }

        if (!field.Items.TryGetValue("Kids", out PdfObject? kidsObject) ||
            ResolveObject(objects, kidsObject) is not PdfArray kids) {
            return;
        }

        for (int i = 0; i < kids.Items.Count; i++) {
            UpdateFormField(objects, kids.Items[i], fullName, fieldType, fieldFlags, fieldValues, remaining, changedObjectNumbers, visited);
        }
    }

    private static void SetIncrementalFieldValue(PdfDictionary field, string? fieldType, string value) {
        if (string.Equals(fieldType, "Btn", StringComparison.Ordinal)) {
            string name = IsOffButtonValue(value) ? "Off" : value;
            field.Items["V"] = new PdfName(name);
            field.Items["AS"] = new PdfName(name);
            return;
        }

        field.Items["V"] = new PdfStringObj(value, useTextStringEncoding: true);
    }

    private static bool IsOffButtonValue(string value) =>
        string.IsNullOrWhiteSpace(value) ||
        string.Equals(value, "false", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(value, "off", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(value, "0", StringComparison.Ordinal);

    private static byte[] AppendIncrementalObjects(
        byte[] pdf,
        Dictionary<int, PdfIndirectObject> objects,
        PdfDocumentSecurityInfo security,
        string trailerRaw,
        HashSet<int> changedObjectNumbers) {
        if (!security.RootObjectNumber.HasValue) {
            throw new InvalidOperationException("PDF root catalog reference is required for an incremental update.");
        }

        if (!security.LastStartXrefOffset.HasValue) {
            throw new InvalidOperationException("PDF startxref offset is required for an incremental update.");
        }

        var identityMap = objects.Keys.ToDictionary(static objectNumber => objectNumber, static objectNumber => objectNumber);
        var context = new PdfPageExtractor.SerializationContext(identityMap, pagesObjectId: 0, new Dictionary<int, Dictionary<string, PdfObject>>(), objects);
        int[] objectNumbers = changedObjectNumbers.OrderBy(static objectNumber => objectNumber).ToArray();
        var serialized = new List<(int ObjectNumber, byte[] Bytes)>(objectNumbers.Length);
        foreach (int objectNumber in objectNumbers) {
            if (!objects.TryGetValue(objectNumber, out PdfIndirectObject? indirect)) {
                throw new InvalidOperationException("PDF object " + objectNumber.ToString(CultureInfo.InvariantCulture) + " was changed but could not be found.");
            }

            serialized.Add((objectNumber, PdfObjectBytes.WrapIndirectObject(objectNumber, PdfPageExtractor.SerializeObject(indirect.Value, context))));
        }

        using var output = new MemoryStream(pdf.Length + serialized.Sum(static item => item.Bytes.Length) + (serialized.Count * 32) + 256);
        output.Write(pdf, 0, pdf.Length);
        if (pdf.Length == 0 || (pdf[pdf.Length - 1] != (byte)'\n' && pdf[pdf.Length - 1] != (byte)'\r')) {
            output.WriteByte((byte)'\n');
        }

        var offsets = new Dictionary<int, long>();
        foreach (var item in serialized) {
            offsets[item.ObjectNumber] = output.Position;
            output.Write(item.Bytes, 0, item.Bytes.Length);
        }

        long xrefOffset = output.Position;
        int size = Math.Max(objects.Keys.Max(), objectNumbers.Max()) + 1;

        using var writer = new StreamWriter(output, Encoding.ASCII, 1024, leaveOpen: true) { NewLine = "\n" };
        writer.WriteLine("xref");
        foreach (int objectNumber in objectNumbers) {
            writer.WriteLine(objectNumber.ToString(CultureInfo.InvariantCulture) + " 1");
            writer.WriteLine(offsets[objectNumber].ToString("0000000000", CultureInfo.InvariantCulture) + " 00000 n ");
        }

        writer.WriteLine("trailer");
        writer.WriteLine("<< /Size " + size.ToString(CultureInfo.InvariantCulture) +
            " /Root " + PdfSyntaxEscaper.IndirectReference(security.RootObjectNumber.Value) +
            (security.InfoObjectNumber.HasValue ? " /Info " + PdfSyntaxEscaper.IndirectReference(security.InfoObjectNumber.Value) : string.Empty) +
            " /Prev " + security.LastStartXrefOffset.Value.ToString(CultureInfo.InvariantCulture) +
            ReadTrailerIdEntry(trailerRaw) +
            " >>");
        writer.WriteLine("startxref");
        writer.WriteLine(xrefOffset.ToString(CultureInfo.InvariantCulture));
        writer.WriteLine("%%EOF");
        writer.Flush();

        return output.ToArray();
    }

    private static PdfObject? ResolveObject(Dictionary<int, PdfIndirectObject> objects, PdfObject? value) =>
        PdfObjectLookup.Resolve(objects, value);

    private static PdfDictionary? ResolveDictionary(Dictionary<int, PdfIndirectObject> objects, PdfObject? value) =>
        ResolveObject(objects, value) as PdfDictionary;

    private static string? TryReadName(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, string key) =>
        dictionary.Items.TryGetValue(key, out PdfObject? value) &&
        ResolveObject(objects, value) is PdfName name &&
        !string.IsNullOrEmpty(name.Name)
            ? name.Name
            : null;

    private static string? TryReadText(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, string key) =>
        dictionary.Items.TryGetValue(key, out PdfObject? value) &&
        ResolveObject(objects, value) is PdfStringObj text
            ? text.Value
            : null;

    private static int ReadFieldFlags(Dictionary<int, PdfIndirectObject> objects, PdfDictionary field, int inheritedFlags) {
        if (!field.Items.TryGetValue("Ff", out PdfObject? flagsObject) ||
            ResolveObject(objects, flagsObject) is not PdfNumber flagsNumber) {
            return inheritedFlags;
        }

        return (int)flagsNumber.Value;
    }

    private static string? CombineFieldName(string? parentName, string? partialName) {
        if (string.IsNullOrEmpty(partialName)) {
            return parentName;
        }

        return string.IsNullOrEmpty(parentName) ? partialName : parentName + "." + partialName;
    }
}
