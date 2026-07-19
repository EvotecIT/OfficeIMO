using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

internal static partial class PdfSyntax {
#if NET8_0_OR_GREATER
    private static readonly Regex StartXrefRegex = new Regex(@"startxref\s+(\d+)", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex ObjectHeaderTemplateRegex = new Regex(@"^\s*(\d+)\s+(\d+)\s+obj\b", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
#else
    private static readonly Regex StartXrefRegex = new Regex(@"startxref\s+(\d+)", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex ObjectHeaderTemplateRegex = new Regex(@"^\s*(\d+)\s+(\d+)\s+obj\b", RegexOptions.Compiled, RegexTimeout);
#endif

    internal static PdfDocumentSecurityInfo ReadDocumentSecurityInfo(byte[] pdf, PdfReadOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        PdfReadLimits limits = options?.Limits ?? new PdfReadLimits();
        limits.Validate();
        if (pdf.LongLength > limits.MaxInputBytes) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.InputBytes, limits.MaxInputBytes, pdf.LongLength);
        }

        string text = PdfEncoding.Latin1GetString(pdf);
        int? encryptObjectNumber = TryReadLastReferenceObjectNumber(text, "Encrypt");
        bool hasEncryption = encryptObjectNumber.HasValue;
        bool hasSignatures = HasSignatureMarkers(pdf);
        IReadOnlyList<int> startXrefOffsets = ReadStartXrefOffsets(text, limits.MaxRevisions);
        int startXrefCount = startXrefOffsets.Count;
        int? lastStartXrefOffset = startXrefOffsets.Count == 0 ? null : startXrefOffsets[startXrefOffsets.Count - 1];
        IReadOnlyList<int> previousXrefOffsets = ReadIntegerNameValues(text, "Prev", limits.MaxRevisions);
        bool hasPreviousRevision = previousXrefOffsets.Count > 0;
        IReadOnlyList<PdfDocumentRevisionInfo> revisions = BuildRevisionInfo(startXrefOffsets, previousXrefOffsets);
        bool hasXrefStreams = ContainsPdfName(text, "XRef") && ContainsPdfName(text, "W");
        bool hasObjectStreams = ContainsPdfName(text, "ObjStm");
        bool hasTrailerId = ContainsPdfName(text, "ID");

        PdfReference? rootReference = TryReadLastReference(text, "Root");
        int? rootObjectNumber = rootReference?.ObjectNumber;
        int? rootObjectGeneration = rootReference?.Generation;
        PdfReference? infoReference = TryReadLastReference(text, "Info");
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
            TryReadObjectDictionary(text, encryptObjectNumber.Value, out PdfDictionary? encryptionDictionary) &&
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

        try {
            var (objects, trailerRaw) = ParseObjects(pdf, options);
            rootReference = TryReadLastReference(trailerRaw, "Root");
            if (rootReference is not null) {
                rootObjectNumber = rootReference.ObjectNumber;
                rootObjectGeneration = rootReference.Generation;
            }

            infoReference = TryReadLastReference(trailerRaw, "Info");
            if (infoReference is not null) {
                infoObjectNumber = infoReference.ObjectNumber;
                infoObjectGeneration = infoReference.Generation;
            }

            PdfReference? encryptReference = TryReadLastReference(trailerRaw, "Encrypt");
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

            PdfDictionary? catalog = FindCatalog(objects, trailerRaw);
            if (catalog is not null) {
                documentSecurityStore = ReadDocumentSecurityStoreInfo(objects, catalog);
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
                    usageRightsObjectNumbers);
            }

            foreach (var entry in objects.OrderBy(static item => item.Key)) {
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
                    if (!string.IsNullOrEmpty(fieldName) && !signatureFieldNames.Contains(fieldName!)) {
                        signatureFieldNames.Add(fieldName!);
                    }

                    if (dictionary.Items.TryGetValue("V", out PdfObject? valueObject) &&
                        valueObject is PdfReference valueReference) {
                        signatureFieldsByValue[valueReference.ObjectNumber] = new SignatureFieldState(
                            entry.Key,
                            fieldName,
                            ReadSignatureFieldLockInfo(objects, dictionary),
                            ReadSignatureSeedValueInfo(objects, dictionary));
                    }
                }
            }

            foreach (var entry in objects.OrderBy(static item => item.Key)) {
                PdfDictionary? dictionary = entry.Value.Value switch {
                    PdfDictionary directDictionary => directDictionary,
                    PdfStream stream => stream.Dictionary,
                    _ => null
                };

                if (dictionary is null) {
                    continue;
                }

                bool isSignatureValue = TryReadName(objects, dictionary, "Type") == "Sig";
                if (TryReadByteRangeValues(objects, dictionary, out IReadOnlyList<long> currentByteRangeValues)) {
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
                        currentByteRangeValues));
                }
            }
        } catch (Exception ex) when (ex is not OutOfMemoryException && ex is not StackOverflowException) {
            signatureValueCount = CountPdfNameOccurrences(text, "ByteRange");
            byteRangeValueCount = 0;
        }

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
            byteRangeValueCount > 0 || ContainsPdfName(text, "ByteRange"),
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
        List<int> usageRightsObjectNumbers) {
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
                        out docMDPPermissionLevel);
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
        IReadOnlyList<long> byteRangeValues) {
        bool hasByteRange = byteRangeValues.Count > 0;
        bool hasContents = dictionary.Items.ContainsKey("Contents");
        int? contentsSizeBytes = TryReadContentsSizeBytes(objects, dictionary);
        int? contentsEncodedSizeBytes = TryReadContentsEncodedSizeBytes(dictionary);
        int referenceCount = TryReadReferenceCount(objects, dictionary);

        return new PdfSignatureInfo(
            objectNumber,
            field?.FieldObjectNumber,
            field?.FieldName,
            field?.FieldLock,
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
            TryReadContentsBytes(objects, dictionary),
            contentsSizeBytes,
            contentsEncodedSizeBytes,
            referenceCount);
    }

    private static PdfSignatureFieldLockInfo? ReadSignatureFieldLockInfo(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary signatureField) {
        if (!signatureField.Items.TryGetValue("Lock", out PdfObject? lockObject) ||
            ResolveObject(objects, lockObject) is not PdfDictionary lockDictionary) {
            return null;
        }

        string? action = TryReadName(objects, lockDictionary, "Action");
        IReadOnlyList<string> fields = ReadNameOrTextArray(objects, lockDictionary, "Fields");
        return !string.IsNullOrEmpty(action) || fields.Count > 0
            ? new PdfSignatureFieldLockInfo(action, fields)
            : null;
    }

    private static PdfSignatureSeedValueInfo? ReadSignatureSeedValueInfo(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary signatureField) {
        if (!signatureField.Items.TryGetValue("SV", out PdfObject? seedValueObject) ||
            ResolveObject(objects, seedValueObject) is not PdfDictionary seedValue) {
            return null;
        }

        string? filter = TryReadName(objects, seedValue, "Filter");
        IReadOnlyList<string> subFilters = ReadNameOrTextArray(objects, seedValue, "SubFilter");
        IReadOnlyList<string> digestMethods = ReadNameOrTextArray(objects, seedValue, "DigestMethod");
        IReadOnlyList<string> reasons = ReadNameOrTextArray(objects, seedValue, "Reasons");
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
        out int? permissionLevel) {
        transformMethod = null;
        transformVersion = null;
        permissionLevel = null;

        if (!signature.Items.TryGetValue("Reference", out PdfObject? referenceObject) ||
            ResolveObject(objects, referenceObject) is not PdfArray references) {
            return;
        }

        for (int i = 0; i < references.Items.Count; i++) {
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

    private static bool TryReadObjectDictionary(string text, int objectNumber, out PdfDictionary? dictionary) {
        dictionary = null;
        int searchIndex = 0;
        while (searchIndex < text.Length) {
            int candidateIndex = text.IndexOf(objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture), searchIndex, StringComparison.Ordinal);
            if (candidateIndex < 0) {
                return false;
            }

            int lineStart = candidateIndex;
            while (lineStart > 0 && text[lineStart - 1] != '\n' && text[lineStart - 1] != '\r') {
                lineStart--;
            }

            int lineEnd = candidateIndex;
            while (lineEnd < text.Length && text[lineEnd] != '\n' && text[lineEnd] != '\r') {
                lineEnd++;
            }

            string headerLine = text.Substring(lineStart, lineEnd - lineStart);
            Match headerMatch = ObjectHeaderTemplateRegex.Match(headerLine);
            if (headerMatch.Success &&
                int.TryParse(headerMatch.Groups[1].Value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int foundObjectNumber) &&
                foundObjectNumber == objectNumber) {
                int dictStart = text.IndexOf("<<", lineEnd, StringComparison.Ordinal);
                int objectEnd = text.IndexOf("endobj", lineEnd, StringComparison.Ordinal);
                if (dictStart >= 0 && objectEnd > dictStart) {
                    int dictEnd = FindDictEnd(text, dictStart, objectEnd);
                    if (dictEnd > dictStart) {
                        string dictText = SafeSlice(text, dictStart + 2, dictEnd - (dictStart + 2), 1_000_000);
                        try {
                            dictionary = ParseDictionary(dictText);
                            return true;
                        } catch (Exception ex) when (ex is not OutOfMemoryException && ex is not StackOverflowException) {
                            return false;
                        }
                    }
                }
            }

            searchIndex = candidateIndex + 1;
        }

        return false;
    }

    private static IReadOnlyList<int> ReadStartXrefOffsets(string text, int maxRevisions) {
        var offsets = new List<int>();
        foreach (Match match in StartXrefRegex.Matches(text)) {
            if (int.TryParse(match.Groups[1].Value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int offset)) {
                offsets.Add(offset);
                if (offsets.Count > maxRevisions) {
                    throw PdfReadLimitException.Create(PdfReadLimitKind.Revisions, maxRevisions, offsets.Count);
                }
            }
        }

        return offsets.Count == 0 ? Array.Empty<int>() : offsets.AsReadOnly();
    }

    private static IReadOnlyList<int> ReadIntegerNameValues(string text, string key, int maxValues) {
#if NET8_0_OR_GREATER
        var regex = new Regex(@"/" + Regex.Escape(key) + @"\s+(\d+)", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
#else
        var regex = new Regex(@"/" + Regex.Escape(key) + @"\s+(\d+)", RegexOptions.Compiled, RegexTimeout);
#endif
        var values = new List<int>();
        foreach (Match match in regex.Matches(text)) {
            if (int.TryParse(match.Groups[1].Value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int value)) {
                values.Add(value);
                if (values.Count > maxValues) {
                    throw PdfReadLimitException.Create(PdfReadLimitKind.Revisions, maxValues, values.Count);
                }
            }
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

    private static int? TryReadLastReferenceObjectNumber(string text, string key) {
        return TryReadLastReference(text, key)?.ObjectNumber;
    }

    private static PdfReference? TryReadLastReference(string text, string key) {
#if NET8_0_OR_GREATER
        var regex = new Regex(@"/" + Regex.Escape(key) + @"\s+(\d+)\s+(\d+)\s+R", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
#else
        var regex = new Regex(@"/" + Regex.Escape(key) + @"\s+(\d+)\s+(\d+)\s+R", RegexOptions.Compiled, RegexTimeout);
#endif
        PdfReference? reference = null;
        foreach (Match match in regex.Matches(text)) {
            if (int.TryParse(match.Groups[1].Value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int objectNumber) &&
                int.TryParse(match.Groups[2].Value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int generation)) {
                reference = new PdfReference(objectNumber, generation);
            }
        }

        return reference;
    }

    private static int CountPdfNameOccurrences(string text, string name) {
        int count = 0;
        string token = "/" + name;
        int index = 0;
        while (index < text.Length) {
            index = text.IndexOf(token, index, StringComparison.Ordinal);
            if (index < 0) {
                return count;
            }

            int after = index + token.Length;
            if (after >= text.Length || IsPdfDelimiter(text[after]) || char.IsWhiteSpace(text[after])) {
                count++;
            }

            index = after;
        }

        return count;
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

    private static IReadOnlyList<string> ReadNameOrTextArray(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, string key) {
        if (!dictionary.Items.TryGetValue(key, out PdfObject? value)) {
            return Array.Empty<string>();
        }

        PdfObject? resolved = ResolveObject(objects, value);
        if (resolved is PdfArray array) {
            var values = new List<string>();
            for (int i = 0; i < array.Items.Count; i++) {
                string? item = ReadNameOrText(objects, array.Items[i]);
                if (!string.IsNullOrEmpty(item) && !values.Contains(item!)) {
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

    private static bool TryReadByteRangeValues(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, out IReadOnlyList<long> values) {
        values = Array.Empty<long>();
        if (!dictionary.Items.TryGetValue("ByteRange", out PdfObject? byteRangeObject) ||
            ResolveObject(objects, byteRangeObject) is not PdfArray byteRange) {
            return false;
        }

        var ranges = new List<long>(byteRange.Items.Count);
        for (int i = 0; i < byteRange.Items.Count; i++) {
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

    private static byte[]? TryReadContentsBytes(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary) {
        return dictionary.Items.TryGetValue("Contents", out PdfObject? contentsObject) &&
            ResolveObject(objects, contentsObject) is PdfStringObj contents
            ? (byte[])contents.RawBytes.Clone()
            : null;
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
