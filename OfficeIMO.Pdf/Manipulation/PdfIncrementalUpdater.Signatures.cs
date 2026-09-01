using OfficeIMO.Core.Internal;
using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfIncrementalUpdater {
    private const string SignatureByteRangePlaceholder =
        "00000000000000000000 00000000000000000000 00000000000000000000 00000000000000000000";
    private static readonly byte[] SignatureContentsMarker = PdfEncoding.Latin1GetBytes("/Contents <");
    private static readonly byte[] SignatureContentsName = PdfEncoding.Latin1GetBytes("/Contents");
    private static readonly byte[] SignatureTypeMarker = PdfEncoding.Latin1GetBytes("/Type /Sig");
    private static readonly byte[] DocumentTimestampTypeMarker = PdfEncoding.Latin1GetBytes("/Type /DocTimeStamp");
    private static readonly byte[] SignatureByteRangeMarker = PdfEncoding.Latin1GetBytes("/ByteRange [");
    private static readonly byte[] PdfObjectKeyword = PdfEncoding.Latin1GetBytes("obj");
    private static readonly byte[] PdfEndObjectKeyword = PdfEncoding.Latin1GetBytes("endobj");
    private static readonly byte[] PdfStreamKeyword = PdfEncoding.Latin1GetBytes("stream");
    private static readonly byte[] PdfEndStreamKeyword = PdfEncoding.Latin1GetBytes("endstream");
    private static readonly byte[] PdfLengthName = PdfEncoding.Latin1GetBytes("/Length");

    /// <summary>
    /// Appends an AcroForm signature field and a detached-signature placeholder as a new incremental revision.
    /// The returned byte ranges can be signed by an external CMS/CAdES/TSA provider without adding cryptographic dependencies.
    /// </summary>
    public static PdfExternalSignaturePreparation PrepareExternalSignature(byte[] pdf, PdfExternalSignatureOptions? options = null) =>
        PrepareExternalSignature(pdf, options, readOptions: null);

    internal static PdfExternalSignaturePreparation PrepareExternalSignature(
        byte[] pdf,
        PdfExternalSignatureOptions? options,
        PdfLoadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        PdfExternalSignatureOptions effectiveOptions = options ?? new PdfExternalSignatureOptions();
        ValidateSigningInput(pdf.LongLength, effectiveOptions);
        effectiveOptions.CancellationToken.ThrowIfCancellationRequested();
        ValidateExternalSignatureOptions(effectiveOptions);
        PdfSignatureProfile signatureProfile = ResolveSignatureProfile(effectiveOptions);
        _ = PdfMutationPlanner.RequireAppendOnly(pdf, PdfMutationOperation.PrepareExternalSignature, readOptions);

        PdfDocumentSecurityInfo security = PdfSyntax.ReadDocumentSecurityInfo(pdf, readOptions);

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, readOptions);
        if (!security.RootObjectNumber.HasValue ||
            !objects.TryGetValue(security.RootObjectNumber.Value, out PdfIndirectObject? rootObject) ||
            rootObject.Value is not PdfDictionary catalog) {
            throw new InvalidOperationException("PDF root catalog dictionary is required for external signature preparation.");
        }

        EnsureSignatureFieldNameAvailable(pdf, effectiveOptions.FieldName, readOptions);

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
        var profileChangedObjects = new HashSet<int>();
        ApplySignatureProfile(
            pdf,
            objects,
            catalog,
            signatureField,
            signatureObjectNumber,
            effectiveOptions,
            signatureProfile,
            ref nextObjectNumber,
            ref catalogChanged,
            profileChangedObjects);
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

        foreach (int objectNumber in profileChangedObjects) {
            changedObjects.Add(objectNumber);
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

        return PatchSignatureByteRange(prepared, effectiveOptions, signatureObjectNumber, readOptions);
    }

    /// <summary>Appends an external signature placeholder to a readable PDF stream.</summary>
    public static PdfExternalSignaturePreparation PrepareExternalSignature(Stream input, PdfExternalSignatureOptions? options = null) {
        Guard.NotNull(input, nameof(input));
        if (!input.CanRead) {
            throw new ArgumentException("Stream must be readable.", nameof(input));
        }

        PdfExternalSignatureOptions effectiveOptions = options ?? new PdfExternalSignatureOptions();
        return PrepareExternalSignature(ReadSigningInput(input, effectiveOptions), effectiveOptions);
    }

    /// <summary>Appends an external signature placeholder to a PDF file and writes the prepared PDF to <paramref name="outputPath"/>.</summary>
    public static PdfExternalSignaturePreparation PrepareExternalSignature(string inputPath, string outputPath, PdfExternalSignatureOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNullOrWhiteSpace(outputPath, nameof(outputPath));
        PdfExternalSignatureOptions effectiveOptions = options ?? new PdfExternalSignatureOptions();
        PdfExternalSignaturePreparation preparation;
        using (var input = new FileStream(inputPath, FileMode.Open, FileAccess.Read, FileShare.Read)) {
            preparation = PrepareExternalSignature(ReadSigningInput(input, effectiveOptions), effectiveOptions);
        }
        OfficeFileCommit.WriteAllBytes(outputPath, preparation.PreparedPdf);
        return preparation;
    }

    /// <summary>Injects externally produced CMS/CAdES/TSA bytes into a prepared signature placeholder.</summary>
    public static byte[] ApplyExternalSignature(PdfExternalSignaturePreparation preparation, byte[] signatureContents) {
        Guard.NotNull(preparation, nameof(preparation));
        Guard.NotNull(signatureContents, nameof(signatureContents));
        _ = PdfMutationPlanner.RequireAppendOnly(
            preparation.PreparedPdf,
            PdfMutationOperation.FinalizeExternalSignature,
            preparation.GetCompletionReadOptions(preparation.PreparedPdf.LongLength));
        return ApplyExternalSignature(
            preparation.PreparedPdf,
            signatureContents,
            preparation.ContentsHexOffset,
            preparation.ContentsHexLength);
    }

    /// <summary>Injects externally produced CMS/CAdES/TSA bytes into the only zero-filled prepared signature placeholder found in a PDF.</summary>
    public static byte[] ApplyExternalSignature(byte[] preparedPdf, byte[] signatureContents) =>
        ApplyExternalSignature(preparedPdf, signatureContents, readOptions: null);

    internal static byte[] ApplyExternalSignature(
        byte[] preparedPdf,
        byte[] signatureContents,
        PdfLoadOptions? readOptions) {
        Guard.NotNull(preparedPdf, nameof(preparedPdf));
        Guard.NotNull(signatureContents, nameof(signatureContents));
        _ = PdfReadDocument.Open(preparedPdf, readOptions);
        int placeholderCount = FindZeroFilledSignatureContents(preparedPdf, out int contentsHexOffset, out int contentsHexLength, out _);
        if (placeholderCount == 0) {
            throw new ArgumentException("PDF does not contain a zero-filled external signature /Contents placeholder.", nameof(preparedPdf));
        }

        if (placeholderCount > 1) {
            throw new ArgumentException("PDF contains multiple zero-filled external signature placeholders. Complete the intended PdfExternalSignaturePreparation instead.", nameof(preparedPdf));
        }

        _ = PdfMutationPlanner.RequireAppendOnly(
            preparedPdf,
            PdfMutationOperation.FinalizeExternalSignature,
            readOptions);
        return ApplyExternalSignature(preparedPdf, signatureContents, contentsHexOffset, contentsHexLength);
    }

    internal static bool HasFinalizableExternalSignatureReservation(
        byte[] preparedPdf,
        PdfDocumentSecurityInfo security) {
        int placeholderCount = FindZeroFilledSignatureContents(preparedPdf, out int contentsHexOffset, out int contentsHexLength, out int placeholderObjectNumber);
        if (placeholderCount != 1 || contentsHexLength <= 0 || (contentsHexLength & 1) != 0) return false;

        long contentsLiteralStart = contentsHexOffset - 1L;
        long contentsLiteralEnd = contentsHexOffset + (long)contentsHexLength + 1L;
        if (contentsLiteralStart < 0L || contentsLiteralEnd > preparedPdf.LongLength) return false;

        PdfSignatureInfo[] candidates = security.Signatures.Where(signature =>
            signature.ObjectNumber == placeholderObjectNumber &&
            signature.HasUnsignedContentsPlaceholder &&
            signature.ContentsSizeBytes == contentsHexLength / 2 &&
            HasExactFinalizationByteRange(signature.ByteRangeValues, contentsLiteralStart, contentsLiteralEnd, preparedPdf.LongLength)).ToArray();
        if (candidates.Length != 1) return false;

        PdfSignatureInfo candidate = candidates[0];
        for (int index = 0; index < security.Signatures.Count; index++) {
            PdfSignatureInfo signature = security.Signatures[index];
            if (ReferenceEquals(signature, candidate)) continue;
            if (ByteRangesOverlap(signature.ByteRangeValues, contentsLiteralStart, contentsLiteralEnd)) return false;
        }
        return true;
    }

    private static bool HasExactFinalizationByteRange(
        IReadOnlyList<long> values,
        long contentsLiteralStart,
        long contentsLiteralEnd,
        long fileLength) =>
        values.Count == 4 &&
        values[0] == 0L &&
        values[1] == contentsLiteralStart &&
        values[2] == contentsLiteralEnd &&
        values[3] == fileLength - contentsLiteralEnd;

    private static bool ByteRangesOverlap(IReadOnlyList<long> values, long start, long end) {
        if ((values.Count & 1) != 0) return true;
        for (int index = 0; index < values.Count; index += 2) {
            long rangeStart = values[index];
            long rangeLength = values[index + 1];
            if (rangeStart < 0L || rangeLength < 0L || rangeStart > long.MaxValue - rangeLength) return true;
            long rangeEnd = rangeStart + rangeLength;
            if (rangeStart < end && rangeEnd > start) return true;
        }
        return false;
    }

    /// <summary>Injects externally produced CMS/CAdES/TSA bytes into a prepared signature placeholder in a file.</summary>
    public static void ApplyExternalSignature(string inputPath, string outputPath, byte[] signatureContents) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNullOrWhiteSpace(outputPath, nameof(outputPath));
        OfficeFileCommit.WriteAllBytes(outputPath, ApplyExternalSignature(File.ReadAllBytes(inputPath), signatureContents));
    }

    private static void ValidateExternalSignatureOptions(PdfExternalSignatureOptions options) {
        if (string.IsNullOrWhiteSpace(options.FieldName)) {
            throw new ArgumentException("Signature field name cannot be empty.", nameof(options));
        }

        if (string.IsNullOrWhiteSpace(options.Filter)) {
            throw new ArgumentException("Signature filter cannot be empty.", nameof(options));
        }

        ResolveSignatureProfile(options);
        ResolveSignatureSubFilter(options);
    }

    private static void EnsureSignatureFieldNameAvailable(byte[] pdf, string fieldName, PdfLoadOptions? readOptions) {
        PdfDocumentInfo info = PdfInspector.Inspect(pdf, readOptions);
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
        PdfSignatureProfile profile = ResolveSignatureProfile(options);
        PdfExternalSignatureSubFilter subFilter = ResolveSignatureSubFilter(options);
        string zeros = new string('0', options.ReservedSignatureContentsBytes * 2);
        var builder = new StringBuilder();
        builder.Append("<< /Type /");
        builder.Append(profile == PdfSignatureProfile.DocumentTimestamp ? "DocTimeStamp" : "Sig");
        builder.Append(" /Filter /").Append(PdfSyntaxEscaper.Name(options.Filter));
        builder.Append(" /SubFilter /").Append(PdfSyntaxEscaper.Name(ToSubFilterName(subFilter)));
        builder.Append(" /ByteRange [").Append(SignatureByteRangePlaceholder).Append(']');
        builder.Append(" /Contents <").Append(zeros).Append('>');
        AppendSignatureTextEntry(builder, "Name", options.Name);
        AppendSignatureTextEntry(builder, "Reason", options.Reason);
        AppendSignatureTextEntry(builder, "Location", options.Location);
        AppendSignatureTextEntry(builder, "ContactInfo", options.ContactInfo);
        builder.Append(" /M ").Append(PdfSyntaxEscaper.TextString(FormatSignatureDate(options.SigningTime ?? DateTimeOffset.UtcNow)));
        if (profile == PdfSignatureProfile.Certification) {
            builder.Append(" /Reference [<< /Type /SigRef /TransformMethod /DocMDP /TransformParams << /Type /TransformParams /P ")
                .Append(((int)options.CertificationPermission).ToString(CultureInfo.InvariantCulture))
                .Append(" /V /1.2 >> >>]");
        }
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

    private static PdfExternalSignaturePreparation PatchSignatureByteRange(
        byte[] prepared,
        PdfExternalSignatureOptions options,
        int signatureObjectNumber,
        PdfLoadOptions? readOptions) {
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
            ToSubFilterName(ResolveSignatureSubFilter(options)),
            ResolveSignatureProfile(options),
            ranges,
            contentsLiteralStart + 1,
            options.ReservedSignatureContentsBytes * 2,
            options.ReservedSignatureContentsBytes,
            readOptions);
    }

    private static byte[] ApplyExternalSignature(byte[] preparedPdf, byte[] signatureContents, int contentsHexOffset, int contentsHexLength) {
        if (signatureContents.Length == 0) {
            throw new ArgumentException("Signature contents cannot be empty.", nameof(signatureContents));
        }

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

    private static int FindZeroFilledSignatureContents(
        byte[] pdf,
        out int contentsHexOffset,
        out int contentsHexLength,
        out int contentsObjectNumber) {
        contentsHexOffset = 0;
        contentsHexLength = 0;
        contentsObjectNumber = 0;
        int placeholderCount = 0;
        byte[] marker = SignatureContentsMarker;
        Dictionary<(int ObjectNumber, int Generation), int> indirectLengths =
            ReadIndirectPdfLengthValues(pdf);
        List<PdfIndirectObjectSpan> objectSpans = FindIndirectObjectSpans(pdf, indirectLengths);
        List<PdfByteSpan> streamSpans = FindPdfStreamSpans(pdf, indirectLengths);
        int objectSpanIndex = 0;
        int streamSpanIndex = 0;
        int searchOffset = 0;
        while (true) {
            int markerOffset = IndexOf(pdf, marker, searchOffset);
            if (markerOffset < 0) {
                return placeholderCount;
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
                IsZeroFilled(pdf, start, end - start) &&
                !IsOffsetInsideSpan(streamSpans, markerOffset, ref streamSpanIndex) &&
                TryGetContainingObjectSpan(objectSpans, markerOffset, end + 1, ref objectSpanIndex, out PdfIndirectObjectSpan objectSpan) &&
                IsSignatureContentsPlaceholder(pdf, markerOffset, objectSpan.Start, objectSpan.End, out int objectNumber)) {
                if (placeholderCount == 0) {
                    contentsHexOffset = start;
                    contentsHexLength = end - start;
                    contentsObjectNumber = objectNumber;
                }

                placeholderCount++;
            }

            searchOffset = markerOffset + marker.Length;
        }
    }

    private static bool IsSignatureContentsPlaceholder(
        byte[] pdf,
        int contentsMarkerOffset,
        int objectStart,
        int objectEnd,
        out int objectNumber) {
        objectNumber = 0;
        if (!TryReadIndirectObjectNumber(pdf, objectStart, objectEnd, out objectNumber) ||
            !TryFindSolePdfNameOffset(pdf, SignatureContentsName, objectStart, objectEnd, out int contentsNameOffset) ||
            contentsNameOffset != contentsMarkerOffset) {
            return false;
        }

        bool hasSignatureType =
            IndexOf(pdf, SignatureTypeMarker, objectStart, objectEnd) >= 0 ||
            IndexOf(pdf, DocumentTimestampTypeMarker, objectStart, objectEnd) >= 0;
        return hasSignatureType &&
            IndexOf(pdf, SignatureByteRangeMarker, objectStart, objectEnd) >= 0;
    }

    private static bool TryReadIndirectObjectNumber(
        byte[] pdf,
        int objectStart,
        int objectEnd,
        out int objectNumber) {
        objectNumber = 0;
        int index = objectStart;
        while (index < objectEnd && IsPdfWhitespace(pdf[index])) index++;
        int digitStart = index;
        while (index < objectEnd && pdf[index] >= (byte)'0' && pdf[index] <= (byte)'9') {
            int digit = pdf[index] - (byte)'0';
            if (objectNumber > (int.MaxValue - digit) / 10) return false;
            objectNumber = (objectNumber * 10) + digit;
            index++;
        }
        return index > digitStart && objectNumber > 0 && index < objectEnd && IsPdfWhitespace(pdf[index]);
    }

    private static bool TryFindSolePdfNameOffset(
        byte[] pdf,
        byte[] marker,
        int start,
        int endExclusive,
        out int nameOffset) {
        nameOffset = -1;
        int index = start;
        while (index < endExclusive) {
            if (SkipPdfLexicalValue(pdf, ref index, endExclusive)) continue;
            byte value = pdf[index];
            if (value == (byte)'<' && index + 1 < endExclusive && pdf[index + 1] == (byte)'<') {
                index += 2;
                continue;
            }
            if (value == (byte)'/' && MatchesAt(pdf, marker, index, endExclusive)) {
                int after = index + marker.Length;
                if (after >= endExclusive || IsPdfWhitespace(pdf[after]) || IsPdfDelimiter(pdf[after])) {
                    if (nameOffset >= 0) return false;
                    nameOffset = index;
                }
            }
            index++;
        }
        return nameOffset >= 0;
    }

    private static void SkipPdfLiteralString(byte[] pdf, ref int index, int endExclusive) {
        int depth = 1;
        index++;
        while (index < endExclusive && depth > 0) {
            byte value = pdf[index++];
            if (value == (byte)'\\') {
                if (index < endExclusive) index++;
            } else if (value == (byte)'(') {
                depth++;
            } else if (value == (byte)')') {
                depth--;
            }
        }
    }

    private static bool MatchesAt(byte[] value, byte[] expected, int offset, int endExclusive) {
        if (offset < 0 || offset > endExclusive - expected.Length) return false;
        for (int index = 0; index < expected.Length; index++) {
            if (value[offset + index] != expected[index]) return false;
        }
        return true;
    }

    private static bool IsPdfWhitespace(byte value) =>
        value == 0 || value == (byte)'\t' || value == (byte)'\n' || value == (byte)'\f' || value == (byte)'\r' || value == (byte)' ';

    private static bool IsPdfDelimiter(byte value) =>
        value == (byte)'(' || value == (byte)')' || value == (byte)'<' || value == (byte)'>' ||
        value == (byte)'[' || value == (byte)']' || value == (byte)'{' || value == (byte)'}' ||
        value == (byte)'/' || value == (byte)'%';

    private static List<PdfIndirectObjectSpan> FindIndirectObjectSpans(
        byte[] pdf,
        IReadOnlyDictionary<(int ObjectNumber, int Generation), int> indirectLengths) {
        var spans = new List<PdfIndirectObjectSpan>();
        int index = 0;
        int objectStart = -1;
        while (index < pdf.Length) {
            if (SkipPdfLexicalValue(pdf, ref index, pdf.Length, indirectLengths)) continue;
            if (objectStart >= 0 && MatchesPdfKeyword(pdf, index, pdf.Length, PdfEndObjectKeyword)) {
                spans.Add(new PdfIndirectObjectSpan(objectStart, index));
                objectStart = -1;
                index += PdfEndObjectKeyword.Length;
                continue;
            }
            if (objectStart < 0 && TryMatchIndirectObjectHeader(pdf, index, pdf.Length)) {
                objectStart = index;
            }
            index++;
        }
        return spans;
    }

    private static bool TryGetContainingObjectSpan(
        List<PdfIndirectObjectSpan> spans,
        int markerOffset,
        int literalEndExclusive,
        ref int spanIndex,
        out PdfIndirectObjectSpan span) {
        while (spanIndex < spans.Count && spans[spanIndex].End < markerOffset) spanIndex++;
        if (spanIndex < spans.Count && spans[spanIndex].Start <= markerOffset && spans[spanIndex].End >= literalEndExclusive) {
            span = spans[spanIndex];
            return true;
        }
        span = default;
        return false;
    }

    private static List<PdfByteSpan> FindPdfStreamSpans(
        byte[] pdf,
        IReadOnlyDictionary<(int ObjectNumber, int Generation), int> indirectLengths) {
        var spans = new List<PdfByteSpan>();
        int index = 0;
        while (index < pdf.Length) {
            if (MatchesPdfKeyword(pdf, index, pdf.Length, PdfStreamKeyword)) {
                int start = index + PdfStreamKeyword.Length;
                if (start < pdf.Length && pdf[start] == (byte)'\r') start++;
                if (start < pdf.Length && pdf[start] == (byte)'\n') start++;
                if (!TryFindPdfStreamBoundary(pdf, index, start, pdf.Length, indirectLengths, out int end, out int endStream)) break;
                spans.Add(new PdfByteSpan(start, end));
                index = endStream + PdfEndStreamKeyword.Length;
                continue;
            }
            if (SkipPdfLexicalValue(pdf, ref index, pdf.Length, indirectLengths)) continue;
            index++;
        }
        return spans;
    }

    private static bool IsOffsetInsideSpan(List<PdfByteSpan> spans, int offset, ref int spanIndex) {
        while (spanIndex < spans.Count && spans[spanIndex].End <= offset) spanIndex++;
        return spanIndex < spans.Count && spans[spanIndex].Start <= offset && offset < spans[spanIndex].End;
    }

    private static Dictionary<(int ObjectNumber, int Generation), int> ReadIndirectPdfLengthValues(byte[] pdf) {
        var values = new Dictionary<(int ObjectNumber, int Generation), int>();
        try {
            var (objects, _) = PdfSyntax.ParseObjects(pdf);
            foreach (PdfIndirectObject indirect in objects.Values) {
                if (indirect.Value is not PdfNumber number ||
                    number.Value < 0d ||
                    number.Value > int.MaxValue ||
                    number.Value != Math.Truncate(number.Value)) {
                    continue;
                }
                values[(indirect.ObjectNumber, indirect.Generation)] = (int)number.Value;
            }
        } catch (Exception ex) when (ex is not OutOfMemoryException) {
            // The raw placeholder scan still uses structurally paired boundaries when
            // a malformed PDF cannot provide a trustworthy indirect length value.
        }
        return values;
    }

    private static bool TryFindPdfStreamBoundary(
        byte[] pdf,
        int streamKeywordOffset,
        int streamDataStart,
        int endExclusive,
        IReadOnlyDictionary<(int ObjectNumber, int Generation), int>? indirectLengths,
        out int streamDataEnd,
        out int endStreamOffset) {
        streamDataEnd = -1;
        endStreamOffset = -1;
        if (TryReadPdfStreamLength(pdf, streamKeywordOffset, indirectLengths, out int declaredLength) &&
            declaredLength <= endExclusive - streamDataStart) {
            int declaredEnd = streamDataStart + declaredLength;
            int markerOffset = SkipPdfWhitespaceAndComments(pdf, declaredEnd, endExclusive);
            if (MatchesPdfKeyword(pdf, markerOffset, endExclusive, PdfEndStreamKeyword)) {
                streamDataEnd = declaredEnd;
                endStreamOffset = markerOffset;
                return true;
            }
        }

        int searchOffset = streamDataStart;
        while (searchOffset < endExclusive) {
            int candidate = IndexOfPdfKeyword(pdf, PdfEndStreamKeyword, searchOffset, endExclusive);
            if (candidate < 0) return false;
            int nextToken = SkipPdfWhitespaceAndComments(
                pdf,
                candidate + PdfEndStreamKeyword.Length,
                endExclusive);
            if (MatchesPdfKeyword(pdf, nextToken, endExclusive, PdfEndObjectKeyword)) {
                streamDataEnd = candidate;
                endStreamOffset = candidate;
                return true;
            }
            searchOffset = candidate + PdfEndStreamKeyword.Length;
        }
        return false;
    }

    private static bool TryReadPdfStreamLength(
        byte[] pdf,
        int streamKeywordOffset,
        IReadOnlyDictionary<(int ObjectNumber, int Generation), int>? indirectLengths,
        out int length) {
        length = 0;
        int cursor = streamKeywordOffset - 1;
        while (cursor >= 0 && IsPdfWhitespace(pdf[cursor])) cursor--;
        if (cursor < 1 || pdf[cursor - 1] != (byte)'>' || pdf[cursor] != (byte)'>') return false;

        int dictionaryEnd = cursor + 1;
        int dictionaryStart = -1;
        int depth = 1;
        cursor -= 2;
        while (cursor >= 1) {
            if (pdf[cursor - 1] == (byte)'>' && pdf[cursor] == (byte)'>') {
                depth++;
                cursor -= 2;
                continue;
            }
            if (pdf[cursor - 1] == (byte)'<' && pdf[cursor] == (byte)'<') {
                depth--;
                if (depth == 0) {
                    dictionaryStart = cursor - 1;
                    break;
                }
                cursor -= 2;
                continue;
            }
            cursor--;
        }
        if (dictionaryStart < 0) return false;

        depth = 1;
        int index = dictionaryStart + 2;
        while (index < dictionaryEnd - 1) {
            if (pdf[index] == (byte)'%') {
                while (index < dictionaryEnd && pdf[index] != (byte)'\r' && pdf[index] != (byte)'\n') index++;
                continue;
            }
            if (pdf[index] == (byte)'(') {
                SkipPdfLiteralString(pdf, ref index, dictionaryEnd);
                continue;
            }
            if (pdf[index] == (byte)'<' && index + 1 < dictionaryEnd && pdf[index + 1] != (byte)'<') {
                index++;
                while (index < dictionaryEnd && pdf[index] != (byte)'>') index++;
                if (index < dictionaryEnd) index++;
                continue;
            }
            if (index + 1 < dictionaryEnd && pdf[index] == (byte)'<' && pdf[index + 1] == (byte)'<') {
                depth++;
                index += 2;
                continue;
            }
            if (index + 1 < dictionaryEnd && pdf[index] == (byte)'>' && pdf[index + 1] == (byte)'>') {
                depth--;
                index += 2;
                continue;
            }
            if (depth == 1 && MatchesAt(pdf, PdfLengthName, index, dictionaryEnd)) {
                int afterName = index + PdfLengthName.Length;
                if (afterName < dictionaryEnd && !IsPdfWhitespace(pdf[afterName])) {
                    index++;
                    continue;
                }
                int valueOffset = SkipPdfWhitespaceAndComments(pdf, afterName, dictionaryEnd);
                if (!TryReadNonNegativePdfInteger(pdf, ref valueOffset, dictionaryEnd, out length)) return false;
                int nextToken = SkipPdfWhitespaceAndComments(pdf, valueOffset, dictionaryEnd);
                if (nextToken >= dictionaryEnd || pdf[nextToken] < (byte)'0' || pdf[nextToken] > (byte)'9') {
                    return true;
                }

                int objectNumber = length;
                if (!TryReadNonNegativePdfInteger(pdf, ref nextToken, dictionaryEnd, out int generation)) return false;
                nextToken = SkipPdfWhitespaceAndComments(pdf, nextToken, dictionaryEnd);
                if (nextToken >= dictionaryEnd || pdf[nextToken] != (byte)'R') return false;
                int afterReference = nextToken + 1;
                if (afterReference < dictionaryEnd &&
                    !IsPdfWhitespace(pdf[afterReference]) &&
                    !IsPdfDelimiter(pdf[afterReference])) {
                    return false;
                }
                return indirectLengths != null &&
                    indirectLengths.TryGetValue((objectNumber, generation), out length);
            }
            index++;
        }
        return false;
    }

    private static bool TryReadNonNegativePdfInteger(byte[] pdf, ref int index, int endExclusive, out int value) {
        value = 0;
        int start = index;
        while (index < endExclusive && pdf[index] >= (byte)'0' && pdf[index] <= (byte)'9') {
            int digit = pdf[index] - (byte)'0';
            if (value > (int.MaxValue - digit) / 10) return false;
            value = (value * 10) + digit;
            index++;
        }
        return index > start;
    }

    private static int SkipPdfWhitespaceAndComments(byte[] pdf, int index, int endExclusive) {
        while (index < endExclusive) {
            while (index < endExclusive && IsPdfWhitespace(pdf[index])) index++;
            if (index >= endExclusive || pdf[index] != (byte)'%') return index;
            while (index < endExclusive && pdf[index] != (byte)'\r' && pdf[index] != (byte)'\n') index++;
        }
        return index;
    }

    private static bool SkipPdfLexicalValue(
        byte[] pdf,
        ref int index,
        int endExclusive,
        IReadOnlyDictionary<(int ObjectNumber, int Generation), int>? indirectLengths = null) {
        byte value = pdf[index];
        if (value == (byte)'%') {
            while (index < endExclusive && pdf[index] != (byte)'\r' && pdf[index] != (byte)'\n') index++;
            return true;
        }
        if (value == (byte)'(') {
            SkipPdfLiteralString(pdf, ref index, endExclusive);
            return true;
        }
        if (value == (byte)'<' && (index + 1 >= endExclusive || pdf[index + 1] != (byte)'<')) {
            index++;
            while (index < endExclusive && pdf[index] != (byte)'>') index++;
            if (index < endExclusive) index++;
            return true;
        }
        if (MatchesPdfKeyword(pdf, index, endExclusive, PdfStreamKeyword)) {
            int streamDataStart = index + PdfStreamKeyword.Length;
            if (streamDataStart < endExclusive && pdf[streamDataStart] == (byte)'\r') streamDataStart++;
            if (streamDataStart < endExclusive && pdf[streamDataStart] == (byte)'\n') streamDataStart++;
            index = TryFindPdfStreamBoundary(
                pdf,
                index,
                streamDataStart,
                endExclusive,
                indirectLengths,
                out _,
                out int endStream)
                ? endStream + PdfEndStreamKeyword.Length
                : endExclusive;
            return true;
        }
        return false;
    }

    private static bool TryMatchIndirectObjectHeader(byte[] pdf, int offset, int endExclusive) {
        if (offset > 0 && !IsPdfWhitespace(pdf[offset - 1])) return false;
        int index = offset;
        if (!TrySkipUnsignedInteger(pdf, ref index, endExclusive, requirePositive: true) ||
            !SkipRequiredPdfWhitespace(pdf, ref index, endExclusive) ||
            !TrySkipUnsignedInteger(pdf, ref index, endExclusive, requirePositive: false) ||
            !SkipRequiredPdfWhitespace(pdf, ref index, endExclusive)) {
            return false;
        }
        return MatchesPdfKeyword(pdf, index, endExclusive, PdfObjectKeyword);
    }

    private static bool TrySkipUnsignedInteger(byte[] pdf, ref int index, int endExclusive, bool requirePositive) {
        int start = index;
        bool nonZero = false;
        while (index < endExclusive && pdf[index] >= (byte)'0' && pdf[index] <= (byte)'9') {
            nonZero |= pdf[index] != (byte)'0';
            index++;
        }
        return index > start && (!requirePositive || nonZero);
    }

    private static bool SkipRequiredPdfWhitespace(byte[] pdf, ref int index, int endExclusive) {
        int start = index;
        while (index < endExclusive && IsPdfWhitespace(pdf[index])) index++;
        return index > start;
    }

    private static bool MatchesPdfKeyword(byte[] pdf, int offset, int endExclusive, byte[] keyword) {
        if (offset > 0 && pdf[offset - 1] == (byte)'/') return false;
        if (offset > 0 && !IsPdfWhitespace(pdf[offset - 1]) && !IsPdfDelimiter(pdf[offset - 1])) return false;
        if (!MatchesAt(pdf, keyword, offset, endExclusive)) return false;
        int after = offset + keyword.Length;
        return after >= endExclusive || IsPdfWhitespace(pdf[after]) || IsPdfDelimiter(pdf[after]);
    }

    private static int IndexOfPdfKeyword(byte[] pdf, byte[] keyword, int startOffset, int endExclusive) {
        int lastStart = Math.Min(endExclusive, pdf.Length) - keyword.Length;
        for (int index = Math.Max(0, startOffset); index <= lastStart; index++) {
            if (MatchesPdfKeyword(pdf, index, endExclusive, keyword)) return index;
        }
        return -1;
    }

    private readonly struct PdfIndirectObjectSpan {
        internal PdfIndirectObjectSpan(int start, int end) {
            Start = start;
            End = end;
        }

        internal int Start { get; }
        internal int End { get; }
    }

    private readonly struct PdfByteSpan {
        internal PdfByteSpan(int start, int end) {
            Start = start;
            End = end;
        }

        internal int Start { get; }
        internal int End { get; }
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
        return PdfIncrementalObjectWriter.Append(
            pdf,
            objects,
            security,
            trailerRaw,
            changedObjectNumbers,
            rawObjects);
    }
}
