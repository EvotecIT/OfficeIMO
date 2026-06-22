using System.Globalization;

namespace OfficeIMO.Pdf;

public static partial class PdfIncrementalUpdater {
    private const string SignatureByteRangePlaceholder =
        "00000000000000000000 00000000000000000000 00000000000000000000 00000000000000000000";

    /// <summary>
    /// Appends an AcroForm signature field and a detached-signature placeholder as a new incremental revision.
    /// The returned byte ranges can be signed by an external CMS/CAdES/TSA provider without adding cryptographic dependencies.
    /// </summary>
    public static PdfExternalSignaturePreparation PrepareExternalSignature(byte[] pdf, PdfExternalSignatureOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        PdfExternalSignatureOptions effectiveOptions = options ?? new PdfExternalSignatureOptions();
        ValidateExternalSignatureOptions(effectiveOptions);

        PdfDocumentSecurityInfo security = PdfSyntax.ReadDocumentSecurityInfo(pdf);
        ValidateExternalSignatureInput(security);

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf);
        if (!security.RootObjectNumber.HasValue ||
            !objects.TryGetValue(security.RootObjectNumber.Value, out PdfIndirectObject? rootObject) ||
            rootObject.Value is not PdfDictionary catalog) {
            throw new InvalidOperationException("PDF root catalog dictionary is required for external signature preparation.");
        }

        EnsureSignatureFieldNameAvailable(pdf, effectiveOptions.FieldName);

        int nextObjectNumber = objects.Keys.Count == 0 ? 1 : objects.Keys.Max() + 1;
        int signatureObjectNumber = nextObjectNumber++;
        int signatureFieldObjectNumber = nextObjectNumber++;
        int? acroFormObjectNumber = EnsureAcroForm(objects, catalog, security.RootObjectNumber.Value, ref nextObjectNumber, out PdfDictionary acroForm, out bool catalogChanged);

        PdfArray fields = EnsureAcroFormFieldsArray(objects, acroForm, ref nextObjectNumber, out int? fieldsArrayObjectNumber);
        fields.Items.Add(new PdfReference(signatureFieldObjectNumber, 0));
        acroForm.Items["SigFlags"] = new PdfNumber(3);

        var signatureField = new PdfDictionary();
        signatureField.Items["FT"] = new PdfName("Sig");
        signatureField.Items["T"] = new PdfStringObj(effectiveOptions.FieldName, useTextStringEncoding: true);
        signatureField.Items["V"] = new PdfReference(signatureObjectNumber, 0);
        signatureField.Items["Ff"] = new PdfNumber(0);
        objects[signatureFieldObjectNumber] = new PdfIndirectObject(signatureFieldObjectNumber, 0, signatureField);

        var changedObjects = new HashSet<int> { signatureFieldObjectNumber };
        if (catalogChanged) {
            changedObjects.Add(security.RootObjectNumber.Value);
        }

        if (acroFormObjectNumber.HasValue) {
            changedObjects.Add(acroFormObjectNumber.Value);
        }

        if (fieldsArrayObjectNumber.HasValue) {
            changedObjects.Add(fieldsArrayObjectNumber.Value);
        }

        byte[] signatureBytes = PdfObjectBytes.WrapIndirectObject(
            signatureObjectNumber,
            BuildSignaturePlaceholderDictionary(effectiveOptions));

        byte[] prepared = AppendIncrementalObjectsWithRawObjects(
            pdf,
            objects,
            security,
            trailerRaw,
            changedObjects,
            new[] { (ObjectNumber: signatureObjectNumber, Bytes: signatureBytes) });

        return PatchSignatureByteRange(prepared, effectiveOptions, signatureObjectNumber);
    }

    /// <summary>Appends an external signature placeholder to a readable PDF stream.</summary>
    public static PdfExternalSignaturePreparation PrepareExternalSignature(Stream input, PdfExternalSignatureOptions? options = null) {
        Guard.NotNull(input, nameof(input));
        if (!input.CanRead) {
            throw new ArgumentException("Stream must be readable.", nameof(input));
        }

        using var buffer = new MemoryStream();
        input.CopyTo(buffer);
        return PrepareExternalSignature(buffer.ToArray(), options);
    }

    /// <summary>Appends an external signature placeholder to a PDF file and writes the prepared PDF to <paramref name="outputPath"/>.</summary>
    public static PdfExternalSignaturePreparation PrepareExternalSignature(string inputPath, string outputPath, PdfExternalSignatureOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNullOrWhiteSpace(outputPath, nameof(outputPath));
        PdfExternalSignaturePreparation preparation = PrepareExternalSignature(File.ReadAllBytes(inputPath), options);
        File.WriteAllBytes(outputPath, preparation.PreparedPdf);
        return preparation;
    }

    /// <summary>Injects externally produced CMS/CAdES/TSA bytes into a prepared signature placeholder.</summary>
    public static byte[] ApplyExternalSignature(PdfExternalSignaturePreparation preparation, byte[] signatureContents) {
        Guard.NotNull(preparation, nameof(preparation));
        Guard.NotNull(signatureContents, nameof(signatureContents));
        return ApplyExternalSignature(
            preparation.PreparedPdf,
            signatureContents,
            preparation.ContentsHexOffset,
            preparation.ContentsHexLength);
    }

    /// <summary>Injects externally produced CMS/CAdES/TSA bytes into the first zero-filled prepared signature placeholder found in a PDF.</summary>
    public static byte[] ApplyExternalSignature(byte[] preparedPdf, byte[] signatureContents) {
        Guard.NotNull(preparedPdf, nameof(preparedPdf));
        Guard.NotNull(signatureContents, nameof(signatureContents));
        if (!TryFindZeroFilledSignatureContents(preparedPdf, out int contentsHexOffset, out int contentsHexLength)) {
            throw new ArgumentException("PDF does not contain a zero-filled external signature /Contents placeholder.", nameof(preparedPdf));
        }

        return ApplyExternalSignature(preparedPdf, signatureContents, contentsHexOffset, contentsHexLength);
    }

    /// <summary>Injects externally produced CMS/CAdES/TSA bytes into a prepared signature placeholder in a file.</summary>
    public static void ApplyExternalSignature(string inputPath, string outputPath, byte[] signatureContents) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNullOrWhiteSpace(outputPath, nameof(outputPath));
        File.WriteAllBytes(outputPath, ApplyExternalSignature(File.ReadAllBytes(inputPath), signatureContents));
    }

    private static void ValidateExternalSignatureOptions(PdfExternalSignatureOptions options) {
        if (string.IsNullOrWhiteSpace(options.FieldName)) {
            throw new ArgumentException("Signature field name cannot be empty.", nameof(options));
        }

        if (string.IsNullOrWhiteSpace(options.Filter)) {
            throw new ArgumentException("Signature filter cannot be empty.", nameof(options));
        }
    }

    private static void ValidateExternalSignatureInput(PdfDocumentSecurityInfo security) {
        var blockers = new List<string>();
        if (security.HasEncryption) {
            blockers.Add("Encrypted");
        }

        if (security.HasXrefStreams) {
            blockers.Add("XrefStream");
        }

        if (!security.RootObjectNumber.HasValue) {
            blockers.Add("MissingRoot");
        }

        if (!security.LastStartXrefOffset.HasValue) {
            blockers.Add("MissingStartXref");
        }

        if (security.SignatureFieldCount > 0 || security.SignatureCount > 0 || security.HasByteRange) {
            blockers.Add("Signed");
        }

        if (security.HasDocMDPPermissions) {
            blockers.Add("DocMDP");
        }

        if (blockers.Count > 0) {
            throw new NotSupportedException("External signature preparation is not supported for this PDF: " + string.Join(", ", blockers));
        }
    }

    private static void EnsureSignatureFieldNameAvailable(byte[] pdf, string fieldName) {
        PdfDocumentInfo info = PdfInspector.Inspect(pdf);
        if (info.FormFields.Any(field => string.Equals(field.Name, fieldName, StringComparison.Ordinal))) {
            throw new ArgumentException("PDF already contains a form field named " + fieldName + ".", nameof(fieldName));
        }
    }

    private static int? EnsureAcroForm(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary catalog,
        int catalogObjectNumber,
        ref int nextObjectNumber,
        out PdfDictionary acroForm,
        out bool catalogChanged) {
        catalogChanged = false;
        if (catalog.Items.TryGetValue("AcroForm", out PdfObject? acroFormObject)) {
            if (acroFormObject is PdfReference reference &&
                ResolveDictionary(objects, reference) is PdfDictionary referencedAcroForm) {
                acroForm = referencedAcroForm;
                return reference.ObjectNumber;
            }

            if (ResolveDictionary(objects, acroFormObject) is PdfDictionary directAcroForm) {
                int objectNumber = nextObjectNumber++;
                objects[objectNumber] = new PdfIndirectObject(objectNumber, 0, directAcroForm);
                catalog.Items["AcroForm"] = new PdfReference(objectNumber, 0);
                catalogChanged = true;
                _ = catalogObjectNumber;
                acroForm = directAcroForm;
                return objectNumber;
            }
        }

        int acroFormObjectNumber = nextObjectNumber++;
        acroForm = new PdfDictionary();
        objects[acroFormObjectNumber] = new PdfIndirectObject(acroFormObjectNumber, 0, acroForm);
        catalog.Items["AcroForm"] = new PdfReference(acroFormObjectNumber, 0);
        catalogChanged = true;
        _ = catalogObjectNumber;
        return acroFormObjectNumber;
    }

    private static PdfArray EnsureAcroFormFieldsArray(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary acroForm,
        ref int nextObjectNumber,
        out int? fieldsArrayObjectNumber) {
        fieldsArrayObjectNumber = null;
        if (acroForm.Items.TryGetValue("Fields", out PdfObject? fieldsObject)) {
            if (fieldsObject is PdfReference reference &&
                ResolveObject(objects, reference) is PdfArray referencedFields) {
                fieldsArrayObjectNumber = reference.ObjectNumber;
                return referencedFields;
            }

            if (ResolveObject(objects, fieldsObject) is PdfArray directFields) {
                return directFields;
            }
        }

        var fields = new PdfArray();
        acroForm.Items["Fields"] = fields;
        _ = nextObjectNumber;
        return fields;
    }

    private static string BuildSignaturePlaceholderDictionary(PdfExternalSignatureOptions options) {
        string zeros = new string('0', options.ReservedSignatureContentsBytes * 2);
        var builder = new StringBuilder();
        builder.Append("<< /Type /");
        builder.Append(options.SubFilter == PdfExternalSignatureSubFilter.DocumentTimestamp ? "DocTimeStamp" : "Sig");
        builder.Append(" /Filter /").Append(PdfSyntaxEscaper.Name(options.Filter));
        builder.Append(" /SubFilter /").Append(PdfSyntaxEscaper.Name(ToSubFilterName(options.SubFilter)));
        builder.Append(" /ByteRange [").Append(SignatureByteRangePlaceholder).Append(']');
        builder.Append(" /Contents <").Append(zeros).Append('>');
        AppendSignatureTextEntry(builder, "Name", options.Name);
        AppendSignatureTextEntry(builder, "Reason", options.Reason);
        AppendSignatureTextEntry(builder, "Location", options.Location);
        AppendSignatureTextEntry(builder, "ContactInfo", options.ContactInfo);
        builder.Append(" /M ").Append(PdfSyntaxEscaper.TextString(FormatSignatureDate(options.SigningTime ?? DateTimeOffset.UtcNow)));
        builder.Append(" >>\n");
        return builder.ToString();
    }

    private static void AppendSignatureTextEntry(StringBuilder builder, string key, string? value) {
        if (!string.IsNullOrWhiteSpace(value)) {
            builder.Append(" /").Append(key).Append(' ').Append(PdfSyntaxEscaper.TextString(value!));
        }
    }

    private static string ToSubFilterName(PdfExternalSignatureSubFilter subFilter) {
        switch (subFilter) {
            case PdfExternalSignatureSubFilter.DetachedCms:
                return "adbe.pkcs7.detached";
            case PdfExternalSignatureSubFilter.CadesDetached:
                return "ETSI.CAdES.detached";
            case PdfExternalSignatureSubFilter.DocumentTimestamp:
                return "ETSI.RFC3161";
            default:
                throw new ArgumentOutOfRangeException(nameof(subFilter), "Unsupported PDF signature subfilter.");
        }
    }

    private static string FormatSignatureDate(DateTimeOffset value) {
        DateTimeOffset local = value;
        TimeSpan offset = local.Offset;
        char sign = offset < TimeSpan.Zero ? '-' : '+';
        offset = offset.Duration();
        return string.Concat(
            "D:",
            local.Year.ToString("0000", CultureInfo.InvariantCulture),
            local.Month.ToString("00", CultureInfo.InvariantCulture),
            local.Day.ToString("00", CultureInfo.InvariantCulture),
            local.Hour.ToString("00", CultureInfo.InvariantCulture),
            local.Minute.ToString("00", CultureInfo.InvariantCulture),
            local.Second.ToString("00", CultureInfo.InvariantCulture),
            sign,
            offset.Hours.ToString("00", CultureInfo.InvariantCulture),
            "'",
            offset.Minutes.ToString("00", CultureInfo.InvariantCulture),
            "'");
    }

    private static PdfExternalSignaturePreparation PatchSignatureByteRange(byte[] prepared, PdfExternalSignatureOptions options, int signatureObjectNumber) {
        byte[] output = (byte[])prepared.Clone();
        byte[] objectHeader = PdfEncoding.Latin1GetBytes(signatureObjectNumber.ToString(CultureInfo.InvariantCulture) + " 0 obj");
        int objectStart = IndexOf(output, objectHeader, 0);
        if (objectStart < 0) {
            throw new InvalidOperationException("Prepared PDF does not contain appended signature object " + signatureObjectNumber.ToString(CultureInfo.InvariantCulture) + ".");
        }

        int objectEnd = IndexOf(output, PdfEncoding.Latin1GetBytes("endobj"), objectStart);
        if (objectEnd < 0) {
            objectEnd = output.Length;
        }

        int byteRangeOffset = IndexOf(output, PdfEncoding.Latin1GetBytes(SignatureByteRangePlaceholder), objectStart, objectEnd);
        if (byteRangeOffset < 0) {
            throw new InvalidOperationException("Prepared signature object " + signatureObjectNumber.ToString(CultureInfo.InvariantCulture) + " does not contain the expected /ByteRange placeholder.");
        }

        byte[] contentsMarker = PdfEncoding.Latin1GetBytes("/Contents <" + new string('0', options.ReservedSignatureContentsBytes * 2) + ">");
        int contentsMarkerOffset = IndexOf(output, contentsMarker, byteRangeOffset, objectEnd);
        if (contentsMarkerOffset < 0) {
            throw new InvalidOperationException("Prepared signature object " + signatureObjectNumber.ToString(CultureInfo.InvariantCulture) + " does not contain the expected /Contents placeholder.");
        }

        int contentsLiteralStart = contentsMarkerOffset + "/Contents ".Length;
        int contentsLiteralEndExclusive = contentsLiteralStart + 1 + (options.ReservedSignatureContentsBytes * 2) + 1;
        long[] ranges = {
            0,
            contentsLiteralStart,
            contentsLiteralEndExclusive,
            output.LongLength - contentsLiteralEndExclusive
        };

        string patchedRange = string.Join(" ", ranges.Select(static value => value.ToString("00000000000000000000", CultureInfo.InvariantCulture)).ToArray());
        byte[] patchedRangeBytes = PdfEncoding.Latin1GetBytes(patchedRange);
        Buffer.BlockCopy(patchedRangeBytes, 0, output, byteRangeOffset, patchedRangeBytes.Length);

        return new PdfExternalSignaturePreparation(
            output,
            options.FieldName,
            options.Filter,
            ToSubFilterName(options.SubFilter),
            ranges,
            contentsLiteralStart + 1,
            options.ReservedSignatureContentsBytes * 2,
            options.ReservedSignatureContentsBytes);
    }

    private static byte[] ApplyExternalSignature(byte[] preparedPdf, byte[] signatureContents, int contentsHexOffset, int contentsHexLength) {
        if (signatureContents.Length * 2 > contentsHexLength) {
            throw new ArgumentException("Signature contents require " + signatureContents.Length.ToString(CultureInfo.InvariantCulture) + " bytes, but the prepared PDF reserved " + (contentsHexLength / 2).ToString(CultureInfo.InvariantCulture) + " bytes.", nameof(signatureContents));
        }

        byte[] output = (byte[])preparedPdf.Clone();
        string signatureHex = ToHex(signatureContents);
        byte[] signatureHexBytes = PdfEncoding.Latin1GetBytes(signatureHex);
        Buffer.BlockCopy(signatureHexBytes, 0, output, contentsHexOffset, signatureHexBytes.Length);
        return output;
    }

    private static string ToHex(byte[] bytes) {
        var builder = new StringBuilder(bytes.Length * 2);
        for (int i = 0; i < bytes.Length; i++) {
            builder.Append(bytes[i].ToString("X2", CultureInfo.InvariantCulture));
        }

        return builder.ToString();
    }

    private static bool TryFindZeroFilledSignatureContents(byte[] pdf, out int contentsHexOffset, out int contentsHexLength) {
        contentsHexOffset = 0;
        contentsHexLength = 0;
        byte[] marker = PdfEncoding.Latin1GetBytes("/Contents <");
        int searchOffset = 0;
        while (true) {
            int markerOffset = IndexOf(pdf, marker, searchOffset);
            if (markerOffset < 0) {
                return false;
            }

            int start = markerOffset + marker.Length;
            int end = start;
            while (end < pdf.Length && pdf[end] != (byte)'>') {
                byte value = pdf[end];
                if (!IsHexDigit(value)) {
                    break;
                }

                end++;
            }

            if (end < pdf.Length &&
                pdf[end] == (byte)'>' &&
                end > start &&
                IsZeroFilled(pdf, start, end - start)) {
                contentsHexOffset = start;
                contentsHexLength = end - start;
                return true;
            }

            searchOffset = markerOffset + marker.Length;
        }
    }

    private static bool IsZeroFilled(byte[] bytes, int offset, int length) {
        for (int i = 0; i < length; i++) {
            if (bytes[offset + i] != (byte)'0') {
                return false;
            }
        }

        return true;
    }

    private static bool IsHexDigit(byte value) =>
        (value >= (byte)'0' && value <= (byte)'9') ||
        (value >= (byte)'A' && value <= (byte)'F') ||
        (value >= (byte)'a' && value <= (byte)'f');

    private static int IndexOf(byte[] haystack, byte[] needle, int startOffset) {
        return IndexOf(haystack, needle, startOffset, haystack.Length);
    }

    private static int IndexOf(byte[] haystack, byte[] needle, int startOffset, int endExclusive) {
        if (needle.Length == 0) {
            return startOffset;
        }

        int lastStart = Math.Min(endExclusive, haystack.Length) - needle.Length;
        for (int i = Math.Max(0, startOffset); i <= lastStart; i++) {
            bool match = true;
            for (int j = 0; j < needle.Length; j++) {
                if (haystack[i + j] != needle[j]) {
                    match = false;
                    break;
                }
            }

            if (match) {
                return i;
            }
        }

        return -1;
    }

    private static byte[] AppendIncrementalObjectsWithRawObjects(
        byte[] pdf,
        Dictionary<int, PdfIndirectObject> objects,
        PdfDocumentSecurityInfo security,
        string trailerRaw,
        HashSet<int> changedObjectNumbers,
        IReadOnlyList<(int ObjectNumber, byte[] Bytes)> rawObjects) {
        if (!security.RootObjectNumber.HasValue) {
            throw new InvalidOperationException("PDF root catalog reference is required for an incremental update.");
        }

        if (!security.LastStartXrefOffset.HasValue) {
            throw new InvalidOperationException("PDF startxref offset is required for an incremental update.");
        }

        var contextObjects = new Dictionary<int, PdfIndirectObject>(objects);
        foreach (var rawObject in rawObjects) {
            if (!contextObjects.ContainsKey(rawObject.ObjectNumber)) {
                contextObjects[rawObject.ObjectNumber] = new PdfIndirectObject(rawObject.ObjectNumber, 0, PdfNull.Instance);
            }
        }

        var identityMap = contextObjects.Keys.ToDictionary(static objectNumber => objectNumber, static objectNumber => objectNumber);
        var context = new PdfPageExtractor.SerializationContext(identityMap, pagesObjectId: 0, new Dictionary<int, Dictionary<string, PdfObject>>(), contextObjects, preserveReferenceGenerations: true);
        int[] objectNumbers = changedObjectNumbers
            .Concat(rawObjects.Select(static item => item.ObjectNumber))
            .Distinct()
            .OrderBy(static objectNumber => objectNumber)
            .ToArray();

        var rawByObjectNumber = rawObjects.ToDictionary(static item => item.ObjectNumber, static item => item.Bytes);
        var serialized = new List<(int ObjectNumber, int Generation, byte[] Bytes)>(objectNumbers.Length);
        foreach (int objectNumber in objectNumbers) {
            if (rawByObjectNumber.TryGetValue(objectNumber, out byte[]? rawBytes)) {
                serialized.Add((objectNumber, 0, rawBytes));
                continue;
            }

            if (!objects.TryGetValue(objectNumber, out PdfIndirectObject? indirect)) {
                throw new InvalidOperationException("PDF object " + objectNumber.ToString(CultureInfo.InvariantCulture) + " was changed but could not be found.");
            }

            serialized.Add((objectNumber, indirect.Generation, PdfObjectBytes.WrapIndirectObject(objectNumber, indirect.Generation, PdfPageExtractor.SerializeObject(indirect.Value, context))));
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
        int size = Math.Max(objects.Keys.Concat(rawObjects.Select(static item => item.ObjectNumber)).Max(), objectNumbers.Max()) + 1;

        using var writer = new StreamWriter(output, Encoding.ASCII, 1024, leaveOpen: true) { NewLine = "\n" };
        writer.WriteLine("xref");
        foreach (int objectNumber in objectNumbers) {
            int generation = serialized.First(item => item.ObjectNumber == objectNumber).Generation;
            writer.WriteLine(objectNumber.ToString(CultureInfo.InvariantCulture) + " 1");
            writer.WriteLine(offsets[objectNumber].ToString("0000000000", CultureInfo.InvariantCulture) + " " + generation.ToString("00000", CultureInfo.InvariantCulture) + " n ");
        }

        writer.WriteLine("trailer");
        writer.WriteLine("<< /Size " + size.ToString(CultureInfo.InvariantCulture) +
            " /Root " + BuildExistingTrailerReference(objects, security.RootObjectNumber.Value) +
            (security.InfoObjectNumber.HasValue ? " /Info " + BuildExistingTrailerReference(objects, security.InfoObjectNumber.Value) : string.Empty) +
            " /Prev " + security.LastStartXrefOffset.Value.ToString(CultureInfo.InvariantCulture) +
            ReadTrailerIdEntry(trailerRaw) +
            " >>");
        writer.WriteLine("startxref");
        writer.WriteLine(xrefOffset.ToString(CultureInfo.InvariantCulture));
        writer.WriteLine("%%EOF");
        writer.Flush();

        return output.ToArray();
    }
}
