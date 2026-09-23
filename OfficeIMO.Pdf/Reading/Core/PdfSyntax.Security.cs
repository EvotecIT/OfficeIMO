using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfSyntax {

    internal static PdfDocumentSecurityInfo ReadDocumentSecurityInfo(
        byte[] pdf,
        PdfLoadOptions? options = null,
        bool includeParsedDetails = true,
        CancellationToken cancellationToken = default) =>
        ReadDocumentSecurityInfo(pdf, options, includeParsedDetails, out _, cancellationToken);

    internal static PdfDocumentSecurityInfo ReadDocumentSecurityInfo(
        byte[] pdf,
        PdfLoadOptions? options,
        bool includeParsedDetails,
        out string decodedText,
        CancellationToken cancellationToken) {
        Guard.NotNull(pdf, nameof(pdf));
        cancellationToken.ThrowIfCancellationRequested();
        PdfReadLimits limits = options?.Limits ?? new PdfReadLimits();
        limits.Validate();
        if (pdf.LongLength > limits.MaxInputBytes) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.InputBytes, limits.MaxInputBytes, pdf.LongLength);
        }

        string text = PdfEncoding.Latin1GetStringCancellable(pdf, cancellationToken);
        decodedText = text;
        cancellationToken.ThrowIfCancellationRequested();
        int? encryptObjectNumber = TryReadLastReferenceObjectNumber(text, "Encrypt", cancellationToken);
        bool hasEncryption = encryptObjectNumber.HasValue;
        // The detailed path below already parses the object graph and derives signature
        // fields and values from it. Keep this initial fallback marker scan raw so a
        // cancellation-aware caller does not pay for a second, tokenless parse.
        RawSecurityMarkers markers = ScanRawSecurityMarkers(text, cancellationToken);
        bool hasSignatures = markers.HasSignatures;
        bool hasByteRange = markers.HasByteRange;
        IReadOnlyList<int> startXrefOffsets = ReadStartXrefOffsets(text, limits.MaxRevisions, cancellationToken);
        int startXrefCount = startXrefOffsets.Count;
        int? lastStartXrefOffset = startXrefOffsets.Count == 0 ? null : startXrefOffsets[startXrefOffsets.Count - 1];
        IReadOnlyList<int> previousXrefOffsets = ReadIntegerNameValues(text, "Prev", limits.MaxRevisions, cancellationToken);
        bool hasPreviousRevision = previousXrefOffsets.Count > 0;
        IReadOnlyList<PdfDocumentRevisionInfo> revisions = BuildRevisionInfo(startXrefOffsets, previousXrefOffsets);
        bool hasXrefStreams = markers.HasXrefStreams;
        bool hasObjectStreams = markers.HasObjectStreams;
        bool hasTrailerId = markers.HasTrailerId;

        PdfReference? rootReference = TryReadLastReference(text, "Root", cancellationToken);
        int? rootObjectNumber = rootReference?.ObjectNumber;
        int? rootObjectGeneration = rootReference?.Generation;
        PdfReference? infoReference = TryReadLastReference(text, "Info", cancellationToken);
        int? infoObjectNumber = infoReference?.ObjectNumber;
        int? infoObjectGeneration = infoReference?.Generation;
        string? encryptionFilter = null;
        string? encryptionSubFilter = null;
        int? encryptionVersion = null;
        int? encryptionRevision = null;
        int? encryptionLengthBits = null;
        int? encryptionPermissions = null;
        bool? encryptMetadata = null;
        PdfPasswordAuthenticationRole passwordAuthenticationRole = PdfPasswordAuthenticationRole.None;

        if (encryptObjectNumber.HasValue &&
            TryReadObjectDictionary(text, encryptObjectNumber.Value, out PdfDictionary? encryptionDictionary, cancellationToken) &&
            encryptionDictionary is not null) {
            encryptionFilter = TryReadName(encryptionDictionary, "Filter");
            encryptionSubFilter = TryReadName(encryptionDictionary, "SubFilter");
            encryptionVersion = TryReadInteger(encryptionDictionary, "V");
            encryptionRevision = TryReadInteger(encryptionDictionary, "R");
            encryptionLengthBits = TryReadInteger(encryptionDictionary, "Length");
            encryptionPermissions = TryReadPermissionMask(encryptionDictionary);
            encryptMetadata = TryReadBoolean(encryptionDictionary, "EncryptMetadata");
        }

        var signatureFieldObjectNumbers = new List<int>();
        var signatureFieldNames = new List<string>();
        var signatures = new List<PdfSignatureInfo>();
        var signatureFieldsByValue = new Dictionary<int, SignatureFieldState>();
        int signatureValueCount = 0;
        int byteRangeValueCount = 0;
        int? acroFormSignatureFlags = null;
        bool hasDocMDPPermissions = false;
        int? docMDPSignatureObjectNumber = null;
        string? docMDPTransformMethod = null;
        string? docMDPTransformVersion = null;
        int? docMDPPermissionLevel = null;
        bool hasUsageRights = false;
        var usageRightsObjectNumbers = new List<int>();
        PdfDocumentDssInfo documentSecurityStore = PdfDocumentDssInfo.Empty;

        if (includeParsedDetails) {
            try {
                var (objects, trailerRaw) = ParseObjects(
                    pdf,
                    options,
                    out PdfRepairReport repairReport,
                    out _,
                    cancellationToken);
                cancellationToken.ThrowIfCancellationRequested();
                rootReference = ReadTrailerReference(trailerRaw, "Root", limits, cancellationToken);
                if (rootReference is not null) {
                    rootObjectNumber = rootReference.ObjectNumber;
                    rootObjectGeneration = rootReference.Generation;
                }

                infoReference = ReadTrailerReference(trailerRaw, "Info", limits, cancellationToken);
                if (infoReference is not null) {
                    infoObjectNumber = infoReference.ObjectNumber;
                    infoObjectGeneration = infoReference.Generation;
                }

                PdfReference? encryptReference = ReadTrailerReference(trailerRaw, "Encrypt", limits, cancellationToken);
                encryptObjectNumber = encryptReference?.ObjectNumber;
                hasEncryption = encryptReference is not null;
                encryptionFilter = null;
                encryptionSubFilter = null;
                encryptionVersion = null;
                encryptionRevision = null;
                encryptionLengthBits = null;
                encryptionPermissions = null;
                encryptMetadata = null;
                if (encryptReference is not null &&
                    PdfObjectLookup.TryGet(objects, encryptReference, out PdfIndirectObject? encryptionObject) &&
                    encryptionObject.Value is PdfDictionary parsedEncryptionDictionary) {
                    encryptionFilter = TryReadName(parsedEncryptionDictionary, "Filter");
                    encryptionSubFilter = TryReadName(parsedEncryptionDictionary, "SubFilter");
                    encryptionVersion = TryReadInteger(parsedEncryptionDictionary, "V");
                    encryptionRevision = TryReadInteger(parsedEncryptionDictionary, "R");
                    encryptionLengthBits = TryReadInteger(parsedEncryptionDictionary, "Length");
                    encryptionPermissions = TryReadPermissionMask(parsedEncryptionDictionary);
                    encryptMetadata = TryReadBoolean(parsedEncryptionDictionary, "EncryptMetadata");
                    if (TryCreateDecryptor(objects, trailerRaw, options, out PdfStandardSecurityHandler? authenticatedHandler) &&
                        authenticatedHandler is not null) {
                        passwordAuthenticationRole = authenticatedHandler.AuthenticationRole;
                    }
                }

                PdfDictionary? catalog = FindCatalog(objects, trailerRaw, cancellationToken);
                if (catalog is not null) {
                    documentSecurityStore = ReadDocumentSecurityStoreInfo(objects, catalog, cancellationToken);
                    ReadCatalogSecurityState(
                        objects,
                        catalog,
                        out acroFormSignatureFlags,
                        out hasDocMDPPermissions,
                        out docMDPSignatureObjectNumber,
                        out docMDPTransformMethod,
                        out docMDPTransformVersion,
                        out docMDPPermissionLevel,
                        out hasUsageRights,
                        usageRightsObjectNumbers,
                        cancellationToken);
                }

                var orderedObjects = OrderSecurityObjects(objects, cancellationToken);
                var seenFieldNames = new HashSet<string>(StringComparer.Ordinal);
                foreach (var entry in orderedObjects) {
                    cancellationToken.ThrowIfCancellationRequested();
                    PdfDictionary? dictionary = entry.Value.Value switch {
                        PdfDictionary directDictionary => directDictionary,
                        PdfStream stream => stream.Dictionary,
                        _ => null
                    };

                    if (dictionary is null) {
                        continue;
                    }

                    if (TryReadName(objects, dictionary, "FT") == "Sig") {
                        signatureFieldObjectNumbers.Add(entry.Key);
                        string? fieldName = TryReadText(objects, dictionary, "T");
                        if (!string.IsNullOrEmpty(fieldName) && seenFieldNames.Add(fieldName!)) {
                            signatureFieldNames.Add(fieldName!);
                        }

                        if (dictionary.Items.TryGetValue("V", out PdfObject? valueObject) &&
                            valueObject is PdfReference valueReference) {
                            signatureFieldsByValue[valueReference.ObjectNumber] = new SignatureFieldState(
                                entry.Key,
                                fieldName,
                                ReadSignatureFieldLockInfo(objects, dictionary, cancellationToken),
                                ReadSignatureSeedValueInfo(objects, dictionary, cancellationToken));
                        }
                    }
                }

                foreach (var entry in orderedObjects) {
                    cancellationToken.ThrowIfCancellationRequested();
                    PdfDictionary? dictionary = entry.Value.Value switch {
                        PdfDictionary directDictionary => directDictionary,
                        PdfStream stream => stream.Dictionary,
                        _ => null
                    };

                    if (dictionary is null) {
                        continue;
                    }

                    bool isSignatureValue = TryReadName(objects, dictionary, "Type") == "Sig";
                    if (TryReadByteRangeValues(objects, dictionary, out IReadOnlyList<long> currentByteRangeValues, cancellationToken)) {
                        isSignatureValue = true;
                        byteRangeValueCount += currentByteRangeValues.Count;
                    }

                    if (isSignatureValue) {
                        signatureValueCount++;
                        signatureFieldsByValue.TryGetValue(entry.Key, out var field);
                        signatures.Add(ReadSignatureInfo(
                            objects,
                            entry.Key,
                            dictionary,
                            field,
                            currentByteRangeValues,
                            cancellationToken));
                    }
                }
                // Successful parsing supersedes raw fallback markers: opaque strings and
                // stream payloads are not signature dictionaries or byte-range arrays.
                hasSignatures = ContainsAnyDocumentPdfName(pdf, objects, repairReport, cancellationToken, "ByteRange", "SigFlags", "Sig");
                hasByteRange = hasSignatures && ContainsAnyDocumentPdfName(pdf, objects, repairReport, cancellationToken, "ByteRange");
            } catch (Exception ex) when (ex is not OperationCanceledException && ex is not OutOfMemoryException && ex is not StackOverflowException) {
                signatureValueCount = markers.ByteRangeCount;
                byteRangeValueCount = 0;
            }
        } else {
            cancellationToken.ThrowIfCancellationRequested();
            signatureValueCount = markers.ByteRangeCount;
            byteRangeValueCount = 0;
        }

        cancellationToken.ThrowIfCancellationRequested();
        return new PdfDocumentSecurityInfo(
            hasEncryption,
            encryptObjectNumber,
            encryptionFilter,
            encryptionSubFilter,
            encryptionVersion,
            encryptionRevision,
            encryptionLengthBits,
            encryptionPermissions,
            encryptMetadata,
            passwordAuthenticationRole,
            hasSignatures || signatureFieldObjectNumbers.Count > 0 || signatureValueCount > 0,
            signatureFieldObjectNumbers.Count == 0 ? Array.Empty<int>() : signatureFieldObjectNumbers.AsReadOnly(),
            signatureFieldNames.Count == 0 ? Array.Empty<string>() : signatureFieldNames.AsReadOnly(),
            signatures.Count == 0 ? Array.Empty<PdfSignatureInfo>() : signatures.AsReadOnly(),
            signatureValueCount,
            byteRangeValueCount > 0 || hasByteRange,
            byteRangeValueCount,
            acroFormSignatureFlags,
            hasDocMDPPermissions,
            docMDPSignatureObjectNumber,
            docMDPTransformMethod,
            docMDPTransformVersion,
            docMDPPermissionLevel,
            hasUsageRights,
            usageRightsObjectNumbers.Count == 0 ? Array.Empty<int>() : usageRightsObjectNumbers.AsReadOnly(),
            documentSecurityStore,
            rootObjectNumber,
            rootObjectGeneration,
            infoObjectNumber,
            infoObjectGeneration,
            hasTrailerId,
            startXrefCount,
            lastStartXrefOffset,
            startXrefOffsets,
            previousXrefOffsets,
            revisions,
            hasPreviousRevision,
            hasXrefStreams,
            hasObjectStreams);
    }

    private static void ReadCatalogSecurityState(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary catalog,
        out int? acroFormSignatureFlags,
        out bool hasDocMDPPermissions,
        out int? docMDPSignatureObjectNumber,
        out string? docMDPTransformMethod,
        out string? docMDPTransformVersion,
        out int? docMDPPermissionLevel,
        out bool hasUsageRights,
        List<int> usageRightsObjectNumbers,
        CancellationToken cancellationToken) {
        acroFormSignatureFlags = null;
        hasDocMDPPermissions = false;
        docMDPSignatureObjectNumber = null;
        docMDPTransformMethod = null;
        docMDPTransformVersion = null;
        docMDPPermissionLevel = null;
        hasUsageRights = false;

        if (catalog.Items.TryGetValue("AcroForm", out PdfObject? acroFormObject) &&
            ResolveObject(objects, acroFormObject) is PdfDictionary acroForm) {
            acroFormSignatureFlags = TryReadInteger(objects, acroForm, "SigFlags");
        }

        if (catalog.Items.TryGetValue("Perms", out PdfObject? permissionsObject) &&
            ResolveObject(objects, permissionsObject) is PdfDictionary permissions) {
            hasDocMDPPermissions = permissions.Items.ContainsKey("DocMDP");
            if (permissions.Items.TryGetValue("DocMDP", out PdfObject? docMDPObject)) {
                docMDPSignatureObjectNumber = docMDPObject is PdfReference docMDPReference ? docMDPReference.ObjectNumber : null;
                if (ResolveObject(objects, docMDPObject) is PdfDictionary docMDPSignature) {
                    ReadDocMDPTransformState(
                        objects,
                        docMDPSignature,
                        out docMDPTransformMethod,
                        out docMDPTransformVersion,
                        out docMDPPermissionLevel,
                        cancellationToken);
                }
            }

            ReadUsageRightsReference(permissions, "UR", usageRightsObjectNumbers);
            ReadUsageRightsReference(permissions, "UR3", usageRightsObjectNumbers);
            hasUsageRights = permissions.Items.ContainsKey("UR") || permissions.Items.ContainsKey("UR3");
        }
    }

    private static PdfSignatureInfo ReadSignatureInfo(
        Dictionary<int, PdfIndirectObject> objects,
        int objectNumber,
        PdfDictionary dictionary,
        SignatureFieldState? field,
        IReadOnlyList<long> byteRangeValues,
        CancellationToken cancellationToken) {
        bool hasByteRange = byteRangeValues.Count > 0;
        bool hasContents = dictionary.Items.ContainsKey("Contents");
        int? contentsSizeBytes = TryReadContentsSizeBytes(objects, dictionary);
        int? contentsEncodedSizeBytes = TryReadContentsEncodedSizeBytes(dictionary);
        int referenceCount = TryReadReferenceCount(objects, dictionary);
        PdfSignatureFieldLockInfo? fieldLock = ReadSignatureFieldMdpInfo(objects, dictionary, cancellationToken) ?? field?.FieldLock;

        return new PdfSignatureInfo(
            objectNumber,
            field?.FieldObjectNumber,
            field?.FieldName,
            fieldLock,
            field?.SeedValue,
            TryReadName(objects, dictionary, "Filter"),
            TryReadName(objects, dictionary, "SubFilter"),
            TryReadText(objects, dictionary, "Name"),
            TryReadText(objects, dictionary, "Location"),
            TryReadText(objects, dictionary, "Reason"),
            TryReadText(objects, dictionary, "ContactInfo"),
            TryReadText(objects, dictionary, "M"),
            hasByteRange,
            byteRangeValues,
            byteRangeValues.Count,
            hasContents,
            TryReadContentsBytes(objects, dictionary, cancellationToken),
            contentsSizeBytes,
            contentsEncodedSizeBytes,
            referenceCount);
    }

    private static PdfSignatureSeedValueInfo? ReadSignatureSeedValueInfo(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary signatureField,
        CancellationToken cancellationToken) {
        if (!signatureField.Items.TryGetValue("SV", out PdfObject? seedValueObject) ||
            ResolveObject(objects, seedValueObject) is not PdfDictionary seedValue) {
            return null;
        }

        string? filter = TryReadName(objects, seedValue, "Filter");
        IReadOnlyList<string> subFilters = ReadNameOrTextArray(objects, seedValue, "SubFilter", cancellationToken);
        IReadOnlyList<string> digestMethods = ReadNameOrTextArray(objects, seedValue, "DigestMethod", cancellationToken);
        IReadOnlyList<string> reasons = ReadNameOrTextArray(objects, seedValue, "Reasons", cancellationToken);
        int? flags = TryReadInteger(objects, seedValue, "Ff");
        bool? addRevInfo = TryReadBoolean(objects, seedValue, "AddRevInfo");
        int? mdpPermissionLevel = null;
        if (seedValue.Items.TryGetValue("MDP", out PdfObject? mdpObject) &&
            ResolveObject(objects, mdpObject) is PdfDictionary mdp) {
            mdpPermissionLevel = TryReadInteger(objects, mdp, "P");
        }

        return !string.IsNullOrEmpty(filter) ||
            subFilters.Count > 0 ||
            digestMethods.Count > 0 ||
            reasons.Count > 0 ||
            flags.HasValue ||
            addRevInfo.HasValue ||
            mdpPermissionLevel.HasValue
            ? new PdfSignatureSeedValueInfo(filter, subFilters, digestMethods, reasons, flags, addRevInfo, mdpPermissionLevel)
            : null;
    }

    private static void ReadDocMDPTransformState(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary signature,
        out string? transformMethod,
        out string? transformVersion,
        out int? permissionLevel,
        CancellationToken cancellationToken) {
        transformMethod = null;
        transformVersion = null;
        permissionLevel = null;

        if (!signature.Items.TryGetValue("Reference", out PdfObject? referenceObject) ||
            ResolveObject(objects, referenceObject) is not PdfArray references) {
            return;
        }

        for (int i = 0; i < references.Items.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (ResolveObject(objects, references.Items[i]) is not PdfDictionary reference) {
                continue;
            }

            string? currentMethod = TryReadName(objects, reference, "TransformMethod");
            if (!string.Equals(currentMethod, "DocMDP", StringComparison.Ordinal)) {
                continue;
            }

            transformMethod = currentMethod;
            if (reference.Items.TryGetValue("TransformParams", out PdfObject? transformParamsObject) &&
                ResolveObject(objects, transformParamsObject) is PdfDictionary transformParams) {
                transformVersion = TryReadNameOrText(objects, transformParams, "V");
                permissionLevel = TryReadInteger(objects, transformParams, "P");
            }

            return;
        }
    }

    private static void ReadUsageRightsReference(PdfDictionary permissions, string key, List<int> objectNumbers) {
        if (permissions.Items.TryGetValue(key, out PdfObject? value) &&
            value is PdfReference reference &&
            !objectNumbers.Contains(reference.ObjectNumber)) {
            objectNumbers.Add(reference.ObjectNumber);
        }
    }

    private static bool TryReadObjectDictionary(string text, int objectNumber, out PdfDictionary? dictionary, CancellationToken cancellationToken = default) {
        dictionary = null;
        int searchIndex = 0;
        while (searchIndex < text.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            int candidateIndex = IndexOfSecurityMarker(text, objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture), searchIndex, cancellationToken);
            if (candidateIndex < 0) {
                return false;
            }

            int lineStart = candidateIndex;
            while (lineStart > 0 && text[lineStart - 1] != '\n' && text[lineStart - 1] != '\r') {
                if ((lineStart & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                lineStart--;
            }

            int lineEnd = candidateIndex;
            while (lineEnd < text.Length && text[lineEnd] != '\n' && text[lineEnd] != '\r') {
                if ((lineEnd & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                lineEnd++;
            }

            if (IsSecurityObjectHeader(text, lineStart, lineEnd, objectNumber, cancellationToken)) {
                int dictStart = IndexOfSecurityMarker(text, "<<", lineEnd, cancellationToken);
                int objectEnd = IndexOfSecurityMarker(text, "endobj", lineEnd, cancellationToken);
                if (dictStart >= 0 && objectEnd > dictStart) {
                    int dictEnd = FindSecurityDictionaryEnd(text, dictStart, objectEnd, cancellationToken);
                    if (dictEnd > dictStart) {
                        string dictText = SafeSlice(text, dictStart + 2, dictEnd - (dictStart + 2), 1_000_000);
                        try {
                            dictionary = ParseDictionary(dictText);
                            return true;
                        } catch (Exception ex) when (ex is not OperationCanceledException && ex is not OutOfMemoryException && ex is not StackOverflowException) {
                            return false;
                        }
                    }
                }
            }

            // A line has only one possible object header. Do not rescan a long
            // malformed line for each later occurrence of the object number.
            searchIndex = lineEnd < text.Length ? lineEnd + 1 : text.Length;
        }

        return false;
    }

    private static bool IsSecurityObjectHeader(string text, int lineStart, int lineEnd, int objectNumber,
        CancellationToken cancellationToken) {
        int cursor = lineStart;
        while (cursor < lineEnd && char.IsWhiteSpace(text[cursor])) {
            if ((cursor & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            cursor++;
        }

        if (!TryReadNonNegativeInteger(text, ref cursor, out int foundObjectNumber, cancellationToken) ||
            cursor > lineEnd || foundObjectNumber != objectNumber ||
            !TrySkipRequiredWhitespace(text, ref cursor, cancellationToken) || cursor > lineEnd ||
            !TryReadNonNegativeInteger(text, ref cursor, out _, cancellationToken) || cursor > lineEnd ||
            !TrySkipRequiredWhitespace(text, ref cursor, cancellationToken) || cursor > lineEnd ||
            cursor + 3 > lineEnd || string.CompareOrdinal(text, cursor, "obj", 0, 3) != 0) {
            return false;
        }

        int afterObjectKeyword = cursor + 3;
        return afterObjectKeyword == lineEnd ||
            (!char.IsLetterOrDigit(text[afterObjectKeyword]) && text[afterObjectKeyword] != '_');
    }

    private static IReadOnlyList<int> ReadStartXrefOffsets(string text, int maxRevisions, CancellationToken cancellationToken = default) {
        var offsets = new List<int>();
        int cursor = 0;
        while (TryReadNextStartXrefOffset(text, ref cursor, out int offset, cancellationToken)) {
            offsets.Add(offset);
            if (offsets.Count > maxRevisions) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.Revisions, maxRevisions, offsets.Count);
            }
        }

        return offsets.Count == 0 ? Array.Empty<int>() : offsets.AsReadOnly();
    }

    /// <summary>Reads the same bounded decimal revision markers used by raw security inspection.</summary>
    private static bool TryReadNextStartXrefOffset(string text, ref int cursor, out int offset, CancellationToken cancellationToken = default) {
        const string marker = "startxref";
        while (cursor < text.Length) {
            int markerIndex = IndexOfSecurityMarker(text, marker, cursor, cancellationToken);
            if (markerIndex < 0) break;

            cursor = markerIndex + 1;
            int digitIndex = markerIndex + marker.Length;
            if (digitIndex >= text.Length || !char.IsWhiteSpace(text[digitIndex])) continue;
            do {
                if ((digitIndex & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                digitIndex++;
            } while (digitIndex < text.Length && char.IsWhiteSpace(text[digitIndex]));

            int value = 0;
            bool overflow = false;
            int end = digitIndex;
            while (end < text.Length && text[end] >= '0' && text[end] <= '9') {
                if ((end & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                int digit = text[end] - '0';
                if (!overflow) {
                    if (value > (int.MaxValue - digit) / 10) overflow = true;
                    else value = value * 10 + digit;
                }
                end++;
            }

            if (end == digitIndex) continue;
            cursor = end;
            if (overflow) continue;
            offset = value;
            return true;
        }

        cursor = text.Length;
        offset = 0;
        return false;
    }

    private static IReadOnlyList<int> ReadIntegerNameValues(string text, string key, int maxValues, CancellationToken cancellationToken = default) {
        var values = new List<int>();
        string token = "/" + key;
        int searchIndex = 0;
        while (TryFindIntegerNameValue(text, token, searchIndex, out int value, out int nextIndex, cancellationToken)) {
            values.Add(value);
            if (values.Count > maxValues) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.Revisions, maxValues, values.Count);
            }

            searchIndex = nextIndex;
        }

        return values.Count == 0 ? Array.Empty<int>() : values.AsReadOnly();
    }

    private static IReadOnlyList<PdfDocumentRevisionInfo> BuildRevisionInfo(IReadOnlyList<int> startXrefOffsets, IReadOnlyList<int> previousXrefOffsets) {
        if (startXrefOffsets.Count == 0) {
            return Array.Empty<PdfDocumentRevisionInfo>();
        }

        var revisions = new List<PdfDocumentRevisionInfo>(startXrefOffsets.Count);
        int firstPreviousRevisionIndex = Math.Max(0, startXrefOffsets.Count - previousXrefOffsets.Count);
        for (int i = 0; i < startXrefOffsets.Count; i++) {
            int? previousOffset = null;
            int index = i - firstPreviousRevisionIndex;
            if (index >= 0 && index < previousXrefOffsets.Count) {
                previousOffset = previousXrefOffsets[index];
            }

            revisions.Add(new PdfDocumentRevisionInfo(i + 1, startXrefOffsets[i], previousOffset));
        }

        return revisions.AsReadOnly();
    }

    private static int? TryReadLastReferenceObjectNumber(string text, string key, CancellationToken cancellationToken = default) {
        return TryReadLastReference(text, key, cancellationToken)?.ObjectNumber;
    }

    private static int? TryReadFirstReferenceObjectNumber(string text, string key) {
        return TryReadFirstReference(text, key)?.ObjectNumber;
    }

    internal static PdfReference? TryReadFirstReference(string text, string key) {
        return TryFindReference(text, "/" + key, 0, out PdfReference? reference, out _)
            ? reference
            : null;
    }

    private static PdfReference? TryReadLastReference(string text, string key, CancellationToken cancellationToken = default) {
        PdfReference? reference = null;
        string token = "/" + key;
        int searchIndex = 0;
        while (TryFindReference(text, token, searchIndex, out PdfReference? candidate, out int nextIndex, cancellationToken)) {
            reference = candidate;
            searchIndex = nextIndex;
        }

        return reference;
    }

    private static bool TryFindIntegerNameValue(
        string text,
        string token,
        int searchIndex,
        out int value,
        out int nextIndex,
        CancellationToken cancellationToken = default) {
        value = 0;
        nextIndex = text.Length;
        while (searchIndex < text.Length) {
            int tokenIndex = IndexOfSecurityMarker(text, token, searchIndex, cancellationToken);
            if (tokenIndex < 0) return false;

            int cursor = tokenIndex + token.Length;
            if (TrySkipRequiredWhitespace(text, ref cursor, cancellationToken) &&
                TryReadNonNegativeInteger(text, ref cursor, out value, cancellationToken)) {
                nextIndex = cursor;
                return true;
            }

            searchIndex = tokenIndex + token.Length;
        }

        return false;
    }

    private static bool TryFindReference(
        string text,
        string token,
        int searchIndex,
        out PdfReference? reference,
        out int nextIndex,
        CancellationToken cancellationToken = default) {
        reference = null;
        nextIndex = text.Length;
        while (searchIndex < text.Length) {
            int tokenIndex = IndexOfSecurityMarker(text, token, searchIndex, cancellationToken);
            if (tokenIndex < 0) return false;

            int cursor = tokenIndex + token.Length;
            if (TrySkipRequiredWhitespace(text, ref cursor, cancellationToken) &&
                TryReadNonNegativeInteger(text, ref cursor, out int objectNumber, cancellationToken) &&
                TrySkipRequiredWhitespace(text, ref cursor, cancellationToken) &&
                TryReadNonNegativeInteger(text, ref cursor, out int generation, cancellationToken) &&
                TrySkipRequiredWhitespace(text, ref cursor, cancellationToken) &&
                cursor < text.Length &&
                text[cursor] == 'R') {
                reference = new PdfReference(objectNumber, generation);
                nextIndex = cursor + 1;
                return true;
            }

            searchIndex = tokenIndex + token.Length;
        }

        return false;
    }

    private static int IndexOfSecurityMarker(string text, string marker, int startIndex, CancellationToken cancellationToken) {
        if (!cancellationToken.CanBeCanceled) return text.IndexOf(marker, startIndex, StringComparison.Ordinal);
        const int window = 65536;
        for (int index = startIndex; index < text.Length; index += window) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = Math.Min(text.Length - index, window + marker.Length - 1);
            int found = text.IndexOf(marker, index, count, StringComparison.Ordinal);
            if (found >= 0) return found;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return -1;
    }

    private static int FindSecurityDictionaryEnd(string text, int start, int limit, CancellationToken cancellationToken) {
        int depth = 0;
        for (int index = start; index + 1 < limit; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            char current = text[index];
            char next = text[index + 1];
            if (current == '<' && next == '<') { depth++; index++; continue; }
            if (current == '>' && next == '>') {
                depth--;
                index++;
                if (depth == 0) return index + 1;
            }
        }
        return -1;
    }

    private static KeyValuePair<int, PdfIndirectObject>[] OrderSecurityObjects(
        Dictionary<int, PdfIndirectObject> objects,
        CancellationToken cancellationToken) {
        var ordered = new KeyValuePair<int, PdfIndirectObject>[objects.Count];
        int index = 0;
        foreach (var item in objects) {
            cancellationToken.ThrowIfCancellationRequested();
            ordered[index++] = item;
        }
        try {
            Array.Sort(ordered, (left, right) => {
                cancellationToken.ThrowIfCancellationRequested();
                return left.Key.CompareTo(right.Key);
            });
        } catch (InvalidOperationException) when (cancellationToken.IsCancellationRequested) {
            cancellationToken.ThrowIfCancellationRequested();
            throw;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return ordered;
    }

    private static bool TrySkipRequiredWhitespace(string text, ref int index, CancellationToken cancellationToken = default) {
        int start = index;
        while (index < text.Length && char.IsWhiteSpace(text[index])) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            index++;
        }
        return index > start;
    }

    private static bool TryReadNonNegativeInteger(string text, ref int index, out int value, CancellationToken cancellationToken = default) {
        value = 0;
        int start = index;
        while (index < text.Length) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            int digit = text[index] - '0';
            if ((uint)digit > 9U) break;
            if (value > (int.MaxValue - digit) / 10) {
                while (index < text.Length && text[index] >= '0' && text[index] <= '9') {
                    if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                    index++;
                }
                return false;
            }

            value = (value * 10) + digit;
            index++;
        }

        return index > start;
    }

    private static string? TryReadText(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, string key) {
        return dictionary.Items.TryGetValue(key, out PdfObject? value) &&
            ResolveObject(objects, value) is PdfStringObj text &&
            !string.IsNullOrEmpty(text.Value)
            ? text.Value
            : null;
    }

    private static string? TryReadName(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, string key) {
        return dictionary.Items.TryGetValue(key, out PdfObject? value) &&
            ResolveObject(objects, value) is PdfName name &&
            !string.IsNullOrEmpty(name.Name)
            ? name.Name
            : null;
    }

    private static string? TryReadNameOrText(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, string key) {
        if (!dictionary.Items.TryGetValue(key, out PdfObject? value)) {
            return null;
        }

        return ResolveObject(objects, value) switch {
            PdfName name when !string.IsNullOrEmpty(name.Name) => name.Name,
            PdfStringObj text when !string.IsNullOrEmpty(text.Value) => text.Value,
            _ => null
        };
    }

    private static IReadOnlyList<string> ReadNameOrTextArray(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, string key, CancellationToken cancellationToken) {
        if (!dictionary.Items.TryGetValue(key, out PdfObject? value)) {
            return Array.Empty<string>();
        }

        PdfObject? resolved = ResolveObject(objects, value);
        if (resolved is PdfArray array) {
            var values = new List<string>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            for (int i = 0; i < array.Items.Count; i++) {
                cancellationToken.ThrowIfCancellationRequested();
                string? item = ReadNameOrText(objects, array.Items[i]);
                if (!string.IsNullOrEmpty(item) && seen.Add(item!)) {
                    values.Add(item!);
                }
            }

            return values.Count == 0 ? Array.Empty<string>() : values.AsReadOnly();
        }

        string? scalar = ReadNameOrText(objects, resolved);
        return string.IsNullOrEmpty(scalar) ? Array.Empty<string>() : new[] { scalar! };
    }

    private static string? ReadNameOrText(Dictionary<int, PdfIndirectObject> objects, PdfObject? value) {
        return ResolveObject(objects, value) switch {
            PdfName name when !string.IsNullOrEmpty(name.Name) => name.Name,
            PdfStringObj text when !string.IsNullOrEmpty(text.Value) => text.Value,
            _ => null
        };
    }

    private static string? TryReadName(PdfDictionary dictionary, string key) {
        return dictionary.Items.TryGetValue(key, out PdfObject? value) &&
            value is PdfName name &&
            !string.IsNullOrEmpty(name.Name)
            ? name.Name
            : null;
    }

    private static bool? TryReadBoolean(PdfDictionary dictionary, string key) {
        return dictionary.Items.TryGetValue(key, out PdfObject? value) && value is PdfBoolean boolean
            ? boolean.Value
            : null;
    }

    private static bool? TryReadBoolean(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, string key) {
        return dictionary.Items.TryGetValue(key, out PdfObject? value) &&
            ResolveObject(objects, value) is PdfBoolean boolean
            ? boolean.Value
            : null;
    }

    private static int? TryReadInteger(PdfDictionary dictionary, string key) {
        return dictionary.Items.TryGetValue(key, out PdfObject? value) && value is PdfNumber number
            ? ToInteger(number)
            : null;
    }

    private static int? TryReadPermissionMask(PdfDictionary dictionary) {
        if (!dictionary.Items.TryGetValue("P", out PdfObject? value) || value is not PdfNumber number) {
            return null;
        }

        if (Math.Truncate(number.Value) != number.Value || number.Value < int.MinValue || number.Value > uint.MaxValue) {
            return null;
        }

        return number.Value > int.MaxValue
            ? unchecked((int)(uint)number.Value)
            : (int)number.Value;
    }

    private static int? TryReadInteger(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, string key) {
        return dictionary.Items.TryGetValue(key, out PdfObject? value) &&
            ResolveObject(objects, value) is PdfNumber number
            ? ToInteger(number)
            : null;
    }

    private static int? ToInteger(PdfNumber number) {
        if (number.Value < int.MinValue ||
            number.Value > int.MaxValue ||
            Math.Truncate(number.Value) != number.Value) {
            return null;
        }

        return (int)number.Value;
    }

    private static bool TryReadByteRangeValues(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, out IReadOnlyList<long> values, CancellationToken cancellationToken) {
        values = Array.Empty<long>();
        if (!dictionary.Items.TryGetValue("ByteRange", out PdfObject? byteRangeObject) ||
            ResolveObject(objects, byteRangeObject) is not PdfArray byteRange) {
            return false;
        }

        var ranges = new List<long>(byteRange.Items.Count);
        for (int i = 0; i < byteRange.Items.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (ResolveObject(objects, byteRange.Items[i]) is PdfNumber number &&
                TryToInt64(number, out long value)) {
                ranges.Add(value);
            }
        }

        values = ranges.Count == 0 ? Array.Empty<long>() : ranges.AsReadOnly();
        return ranges.Count > 0;
    }

    private static bool TryToInt64(PdfNumber number, out long value) {
        value = 0;
        if (number.Value < long.MinValue ||
            number.Value > long.MaxValue ||
            Math.Truncate(number.Value) != number.Value) {
            return false;
        }

        value = (long)number.Value;
        return true;
    }

    private static int? TryReadContentsSizeBytes(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary) {
        return dictionary.Items.TryGetValue("Contents", out PdfObject? contentsObject) &&
            ResolveObject(objects, contentsObject) is PdfStringObj contents
            ? contents.RawBytes.Length
            : null;
    }

    private static byte[]? TryReadContentsBytes(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, CancellationToken cancellationToken) {
        if (!dictionary.Items.TryGetValue("Contents", out PdfObject? contentsObject) ||
            ResolveObject(objects, contentsObject) is not PdfStringObj contents) return null;
        byte[] source = contents.RawBytes;
        cancellationToken.ThrowIfCancellationRequested();
        var copy = new byte[source.Length];
        for (int offset = 0; offset < source.Length; offset += 65536) {
            cancellationToken.ThrowIfCancellationRequested();
            Buffer.BlockCopy(source, offset, copy, offset, Math.Min(65536, source.Length - offset));
        }
        cancellationToken.ThrowIfCancellationRequested();
        return copy;
    }

    private static int? TryReadContentsEncodedSizeBytes(PdfDictionary dictionary) =>
        dictionary.Items.TryGetValue("Contents", out PdfObject? contentsObject) &&
        contentsObject is PdfStringObj contents
            ? contents.EncodedTokenLength
            : null;

    private static int TryReadReferenceCount(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary) {
        if (!dictionary.Items.TryGetValue("Reference", out PdfObject? referenceObject) ||
            ResolveObject(objects, referenceObject) is not PdfArray references) {
            return 0;
        }

        return references.Items.Count;
    }

    private sealed class SignatureFieldState {
        public SignatureFieldState(
            int fieldObjectNumber,
            string? fieldName,
            PdfSignatureFieldLockInfo? fieldLock,
            PdfSignatureSeedValueInfo? seedValue) {
            FieldObjectNumber = fieldObjectNumber;
            FieldName = fieldName;
            FieldLock = fieldLock;
            SeedValue = seedValue;
        }

        public int FieldObjectNumber { get; }

        public string? FieldName { get; }

        public PdfSignatureFieldLockInfo? FieldLock { get; }

        public PdfSignatureSeedValueInfo? SeedValue { get; }
    }
}
