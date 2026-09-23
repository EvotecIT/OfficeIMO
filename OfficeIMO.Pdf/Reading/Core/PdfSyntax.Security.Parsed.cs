using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfSyntax {
    /// <summary>
    /// Builds security and revision evidence for a clear-text artifact emitted by
    /// <see cref="PdfFileAssembler"/>. Callers must only use this after the rewrite
    /// planner has accepted the source and the canonical assembler has produced the bytes.
    /// </summary>
    internal static PdfDocumentSecurityInfo ReadRewrittenOutputSecurityInfo(
        string decodedText,
        string trailerRaw,
        PdfLoadOptions options,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(decodedText, nameof(decodedText));
        Guard.NotNull(trailerRaw, nameof(trailerRaw));
        Guard.NotNull(options, nameof(options));
        cancellationToken.ThrowIfCancellationRequested();
        PdfReadLimits limits = options.Limits;
        var trailerReferences = ReadTrailerReferences(trailerRaw, "Encrypt", "Root", "Info", limits, cancellationToken);
        if (trailerReferences.First is not null) {
            throw new InvalidDataException("The canonical clear-text rewrite unexpectedly emitted an encryption reference.");
        }

        if (!TryGetLatestStartXrefOffset(decodedText, out int startXrefOffset, cancellationToken)) {
            throw new InvalidDataException("The canonical full rewrite did not contain a readable terminal cross-reference pointer.");
        }

        PdfReference? rootReference = trailerReferences.Second;
        PdfReference? infoReference = trailerReferences.Third;
        IReadOnlyList<int> startXrefOffsets = new[] { startXrefOffset };
        IReadOnlyList<int> previousXrefOffsets = Array.Empty<int>();
        IReadOnlyList<PdfDocumentRevisionInfo> revisions = BuildRevisionInfo(startXrefOffsets, previousXrefOffsets);
        cancellationToken.ThrowIfCancellationRequested();
        return new PdfDocumentSecurityInfo(
            hasEncryption: false,
            encryptObjectNumber: null,
            encryptionFilter: null,
            encryptionSubFilter: null,
            encryptionVersion: null,
            encryptionRevision: null,
            encryptionLengthBits: null,
            encryptionPermissions: null,
            encryptMetadata: null,
            PdfPasswordAuthenticationRole.None,
            hasSignatures: false,
            Array.Empty<int>(),
            Array.Empty<string>(),
            Array.Empty<PdfSignatureInfo>(),
            signatureValueCount: 0,
            hasByteRange: false,
            byteRangeValueCount: 0,
            acroFormSignatureFlags: null,
            hasDocMDPPermissions: false,
            docMDPSignatureObjectNumber: null,
            docMDPTransformMethod: null,
            docMDPTransformVersion: null,
            docMDPPermissionLevel: null,
            hasUsageRights: false,
            Array.Empty<int>(),
            PdfDocumentDssInfo.Empty,
            rootReference?.ObjectNumber,
            rootReference?.Generation,
            infoReference?.ObjectNumber,
            infoReference?.Generation,
            hasTrailerId: true,
            startXrefCount: 1,
            lastStartXrefOffset: startXrefOffset,
            startXrefOffsets,
            previousXrefOffsets,
            revisions,
            hasPreviousRevision: false,
            hasXrefStreams: false,
            hasObjectStreams: false);
    }

    internal static PdfDocumentSecurityInfo ReadDocumentSecurityInfo(
        byte[] pdf,
        Dictionary<int, PdfIndirectObject> objects,
        string trailerRaw,
        PdfDocumentSecurityInfo fallback,
        PdfRepairReport repairReport,
        PdfLoadOptions? options = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(objects, nameof(objects));
        Guard.NotNull(fallback, nameof(fallback));
        cancellationToken.ThrowIfCancellationRequested();

        PdfReadLimits limits = options?.Limits ?? new PdfReadLimits();
        var trailerReferences = ReadTrailerReferences(trailerRaw, "Encrypt", "Root", "Info", limits, cancellationToken);
        PdfReference? encryptReference = trailerReferences.First;
        int? encryptObjectNumber = encryptReference?.ObjectNumber;
        bool hasEncryption = encryptObjectNumber.HasValue;
        string? encryptionFilter = null;
        string? encryptionSubFilter = null;
        int? encryptionVersion = null;
        int? encryptionRevision = null;
        int? encryptionLengthBits = null;
        int? encryptionPermissions = null;
        bool? encryptMetadata = null;
        PdfPasswordAuthenticationRole passwordAuthenticationRole = fallback.PasswordAuthenticationRole;
        if (encryptObjectNumber.HasValue &&
            encryptReference is not null &&
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

        var signatureFieldObjectNumbers = new List<int>();
        var signatureFieldNames = new List<string>();
        var seenFieldNames = new HashSet<string>(StringComparer.Ordinal);
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
        KeyValuePair<int, PdfIndirectObject>[] orderedObjects = OrderSecurityObjects(objects, cancellationToken);
        bool hasParsedSignatureMarker = false;
        bool hasParsedByteRangeMarker = false;

        PdfDictionary? catalog = FindCatalog(objects, trailerRaw, cancellationToken);
        if (catalog is not null) {
            cancellationToken.ThrowIfCancellationRequested();
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

        foreach (var entry in orderedObjects) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!hasParsedSignatureMarker || !hasParsedByteRangeMarker) {
                CollectParsedSignatureMarkers(
                    entry.Value.Value,
                    ref hasParsedSignatureMarker,
                    ref hasParsedByteRangeMarker,
                    cancellationToken);
            }
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
                signatureFieldsByValue.TryGetValue(entry.Key, out SignatureFieldState? field);
                signatures.Add(ReadSignatureInfo(
                    objects,
                    entry.Key,
                    dictionary,
                    field,
                    currentByteRangeValues,
                    cancellationToken));
            }
        }

        PdfReference? rootReference = trailerReferences.Second;
        int? rootObjectNumber = rootReference?.ObjectNumber ?? fallback.RootObjectNumber;
        int? rootObjectGeneration = rootReference?.Generation ?? fallback.RootObjectGeneration;
        PdfReference? infoReference = trailerReferences.Third;
        int? infoObjectNumber = infoReference?.ObjectNumber ?? fallback.InfoObjectNumber;
        int? infoObjectGeneration = infoReference?.Generation ?? fallback.InfoObjectGeneration;
        string? rawFallback = null;
        bool hasSignatures = hasParsedSignatureMarker;
        if (!hasSignatures && repairReport.HasIncompleteObjectCoverage) {
            rawFallback = PdfEncoding.Latin1GetStringCancellable(pdf, cancellationToken);
            hasSignatures = ContainsAnyPdfName(rawFallback, cancellationToken, "ByteRange", "SigFlags", "Sig");
        }
        bool hasByteRange = byteRangeValueCount > 0 || hasParsedByteRangeMarker;
        if (!hasByteRange && hasSignatures && repairReport.HasIncompleteObjectCoverage) {
            rawFallback ??= PdfEncoding.Latin1GetStringCancellable(pdf, cancellationToken);
            hasByteRange = ContainsAnyPdfName(rawFallback, cancellationToken, "ByteRange");
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
            hasByteRange,
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
            fallback.HasTrailerId,
            fallback.StartXrefCount,
            fallback.LastStartXrefOffset,
            fallback.StartXrefOffsets,
            fallback.PreviousXrefOffsets,
            fallback.Revisions,
            fallback.HasPreviousRevision,
            fallback.HasXrefStreams,
            fallback.HasObjectStreams);
    }

    private static void CollectParsedSignatureMarkers(
        PdfObject value,
        ref bool hasSignatureMarker,
        ref bool hasByteRangeMarker,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        switch (value) {
            case PdfName name:
                CollectParsedSignatureMarker(name.Name, ref hasSignatureMarker, ref hasByteRangeMarker);
                return;
            case PdfDictionary dictionary:
                foreach (KeyValuePair<string, PdfObject> item in dictionary.Items) {
                    cancellationToken.ThrowIfCancellationRequested();
                    CollectParsedSignatureMarker(item.Key, ref hasSignatureMarker, ref hasByteRangeMarker);
                    if (hasSignatureMarker && hasByteRangeMarker) return;
                    CollectParsedSignatureMarkers(item.Value, ref hasSignatureMarker, ref hasByteRangeMarker, cancellationToken);
                    if (hasSignatureMarker && hasByteRangeMarker) return;
                }
                return;
            case PdfArray array:
                foreach (PdfObject item in array.Items) {
                    CollectParsedSignatureMarkers(item, ref hasSignatureMarker, ref hasByteRangeMarker, cancellationToken);
                    if (hasSignatureMarker && hasByteRangeMarker) return;
                }
                return;
            case PdfStream stream:
                CollectParsedSignatureMarkers(stream.Dictionary, ref hasSignatureMarker, ref hasByteRangeMarker, cancellationToken);
                return;
        }
    }

    private static void CollectParsedSignatureMarker(
        string name,
        ref bool hasSignatureMarker,
        ref bool hasByteRangeMarker) {
        if (name == "ByteRange") {
            hasByteRangeMarker = true;
            hasSignatureMarker = true;
        } else if (name == "SigFlags" || name == "Sig") {
            hasSignatureMarker = true;
        }
    }
}
