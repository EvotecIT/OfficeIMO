namespace OfficeIMO.Pdf;

internal static partial class PdfSyntax {
    internal static PdfDocumentSecurityInfo ReadDocumentSecurityInfo(
        byte[] pdf,
        Dictionary<int, PdfIndirectObject> objects,
        string trailerRaw,
        PdfDocumentSecurityInfo fallback) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(objects, nameof(objects));
        Guard.NotNull(fallback, nameof(fallback));

        string text = PdfEncoding.Latin1GetString(pdf);
        int? encryptObjectNumber = TryReadFirstReferenceObjectNumber(trailerRaw, "Encrypt");
        bool hasEncryption = encryptObjectNumber.HasValue;
        string? encryptionFilter = null;
        string? encryptionSubFilter = null;
        int? encryptionVersion = null;
        int? encryptionRevision = null;
        int? encryptionLengthBits = null;
        int? encryptionPermissions = null;
        bool? encryptMetadata = null;
        if (encryptObjectNumber.HasValue &&
            TryReadFirstReference(trailerRaw, "Encrypt") is PdfReference encryptReference &&
            PdfObjectLookup.TryGet(objects, encryptReference, out PdfIndirectObject? encryptionObject) &&
            encryptionObject.Value is PdfDictionary parsedEncryptionDictionary) {
            encryptionFilter = TryReadName(parsedEncryptionDictionary, "Filter");
            encryptionSubFilter = TryReadName(parsedEncryptionDictionary, "SubFilter");
            encryptionVersion = TryReadInteger(parsedEncryptionDictionary, "V");
            encryptionRevision = TryReadInteger(parsedEncryptionDictionary, "R");
            encryptionLengthBits = TryReadInteger(parsedEncryptionDictionary, "Length");
            encryptionPermissions = TryReadPermissionMask(parsedEncryptionDictionary);
            encryptMetadata = TryReadBoolean(parsedEncryptionDictionary, "EncryptMetadata");
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
                signatureFieldsByValue.TryGetValue(entry.Key, out SignatureFieldState? field);
                signatures.Add(ReadSignatureInfo(
                    objects,
                    entry.Key,
                    dictionary,
                    field,
                    currentByteRangeValues));
            }
        }

        int? rootObjectNumber = TryReadFirstReferenceObjectNumber(trailerRaw, "Root") ?? fallback.RootObjectNumber;
        int? infoObjectNumber = TryReadFirstReferenceObjectNumber(trailerRaw, "Info") ?? fallback.InfoObjectNumber;
        bool hasByteRange = byteRangeValueCount > 0 || ContainsPdfName(text, "ByteRange");

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
            fallback.PasswordAuthenticationRole,
            fallback.HasSignatures || signatureFieldObjectNumbers.Count > 0 || signatureValueCount > 0,
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
            fallback.RootObjectGeneration,
            infoObjectNumber,
            fallback.InfoObjectGeneration,
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
}
