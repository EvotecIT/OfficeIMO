namespace OfficeIMO.Pdf;

internal static partial class PdfSyntax {
    internal static bool TryCreateDecryptor(
        Dictionary<int, PdfIndirectObject> map,
        string trailerRaw,
        PdfLoadOptions? options,
        out PdfStandardSecurityHandler? decryptor,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        decryptor = null;
        PdfReference? encryptReference = FindEncryptReference(map, trailerRaw, options?.Limits, cancellationToken);
        if (encryptReference is null) {
            return false;
        }

        if (!PdfObjectLookup.TryGet(map, encryptReference, out PdfIndirectObject? encryptObject) ||
            encryptObject.Value is not PdfDictionary encryptionDictionary) {
            throw new PdfUnsupportedEncryptionException("PDF encryption dictionary could not be read.");
        }

        byte[] fileId = ReadFirstFileId(map, trailerRaw, cancellationToken);
        bool supplied = options != null && options.Password != null;
        decryptor = PdfStandardSecurityHandler.Create(
            ResolveEncryptionDictionaryReferences(map, encryptionDictionary, cancellationToken),
            fileId,
            options?.Password,
            supplied,
            options?.AesCryptographyProvider,
            options?.PublicKeyDecryption,
            cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return true;
    }

    /// <summary>
    /// Returns the encryption dictionary with indirect entry values, the /CF crypt filter dictionary and each crypt
    /// filter dictionary resolved. ISO 32000-2 7.3.10 allows any of them to be indirect objects, and the security
    /// handler reads them before the document's objects are decrypted.
    /// </summary>
    private static PdfDictionary ResolveEncryptionDictionaryReferences(
        Dictionary<int, PdfIndirectObject> map,
        PdfDictionary encryptionDictionary,
        CancellationToken cancellationToken) {
        if (!ContainsReference(encryptionDictionary) &&
            (encryptionDictionary.Get<PdfDictionary>("CF") is not PdfDictionary directFilters || !ContainsReference(directFilters))) {
            return encryptionDictionary;
        }

        PdfDictionary resolved = ResolveEntries(encryptionDictionary);
        if (resolved.Get<PdfDictionary>("CF") is PdfDictionary filters) {
            PdfDictionary resolvedFilters = ResolveEntries(filters);
            foreach (string name in resolvedFilters.Items.Keys.ToList()) {
                if (resolvedFilters.Items[name] is PdfDictionary filter && ContainsReference(filter)) {
                    resolvedFilters.Items[name] = ResolveEntries(filter);
                }
            }

            resolved.Items["CF"] = resolvedFilters;
        }

        return resolved;

        static bool ContainsReference(PdfDictionary dictionary) {
            foreach (PdfObject value in dictionary.Items.Values) {
                if (value is PdfReference) return true;
            }

            return false;
        }

        PdfDictionary ResolveEntries(PdfDictionary source) {
            var copy = new PdfDictionary(source.Items.Count);
            foreach (KeyValuePair<string, PdfObject> entry in source.Items) {
                cancellationToken.ThrowIfCancellationRequested();
                PdfObject value = entry.Value;
                if (value is PdfReference &&
                    PdfObjectLookup.ResolveChainCancellable(map, value, cancellationToken) is PdfObject target &&
                    target is not PdfStream) {
                    value = target;
                }

                copy.Items[entry.Key] = value;
            }

            return copy;
        }
    }

    /// <summary>
    /// The trailer's /Encrypt reference, or, when the trailer chain cannot be recovered, the one /Encrypt reference
    /// that the document's cross-reference streams agree on.
    /// </summary>
    internal static PdfReference? FindEncryptReference(
        Dictionary<int, PdfIndirectObject> map,
        string trailerRaw,
        PdfReadLimits? limits,
        CancellationToken cancellationToken) {
        PdfReference? reference = ReadTrailerReference(trailerRaw, "Encrypt", limits, cancellationToken);
        if (reference is not null || TrailerDeclaresKey(trailerRaw, "Encrypt", limits, cancellationToken)) return reference;
        PdfReference? found = null;
        foreach (PdfIndirectObject indirect in map.Values) {
            cancellationToken.ThrowIfCancellationRequested();
            if (IsUnbuiltWithOtherType(indirect, "XRef") ||
                indirect.Value is not PdfStream stream ||
                stream.Dictionary.Get<PdfName>("Type")?.Name != "XRef" ||
                stream.Dictionary.Get<PdfReference>("Encrypt") is not PdfReference candidate) {
                continue;
            }

            if (found is not null && (found.ObjectNumber != candidate.ObjectNumber || found.Generation != candidate.Generation)) {
                return null;
            }

            found = candidate;
        }

        return found;
    }

    private static byte[] ReadFirstFileId(Dictionary<int, PdfIndirectObject> map, string trailerRaw,
        CancellationToken cancellationToken) {
        // The active trailer (or the active cross-reference stream it was built from) owns the file identifier;
        // other cross-reference streams are a fallback for files whose trailer chain was not recovered.
        byte[] trailerId = ReadFirstFileId(trailerRaw, cancellationToken);
        if (trailerId.Length > 0) return trailerId;
        foreach (PdfIndirectObject indirect in map.Values) {
            cancellationToken.ThrowIfCancellationRequested();
            if (indirect.Value is PdfStream stream &&
                stream.Dictionary.Get<PdfName>("Type")?.Name == "XRef" &&
                TryReadFirstFileId(stream.Dictionary, out byte[] fileId)) {
                return fileId;
            }
        }

        return trailerId;
    }

    private static bool TryReadFirstFileId(PdfDictionary dictionary, out byte[] fileId) {
        fileId = Array.Empty<byte>();
        if (dictionary.Get<PdfArray>("ID") is PdfArray idArray &&
            idArray.Items.Count > 0 &&
            idArray.Items[0] is PdfStringObj firstId) {
            fileId = firstId.RawBytes;
            return true;
        }

        return false;
    }

    private static byte[] ReadFirstFileId(string trailerRaw, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        int dictStart = trailerRaw.IndexOf("<<", StringComparison.Ordinal);
        if (dictStart < 0) {
            return Array.Empty<byte>();
        }

        int dictEnd = FindDictEnd(trailerRaw, dictStart, trailerRaw.Length, cancellationToken);
        if (dictEnd <= dictStart) {
            return Array.Empty<byte>();
        }

        string dictText = SafeSliceCancellable(trailerRaw, dictStart + 2, dictEnd - (dictStart + 2), 1_000_000, cancellationToken);
        PdfDictionary trailer = ParseDictionary(dictText, cancellationToken: cancellationToken);
        if (trailer.Get<PdfArray>("ID") is PdfArray idArray &&
            idArray.Items.Count > 0 &&
            idArray.Items[0] is PdfStringObj firstId) {
            return firstId.RawBytes;
        }

        return Array.Empty<byte>();
    }

    private static void DecryptObjects(
        Dictionary<int, PdfIndirectObject> map,
        PdfStandardSecurityHandler decryptor,
        int encryptObjectNumber, CancellationToken cancellationToken) {
        var replacements = new List<PdfIndirectObject>();
        decryptor.ReferenceResolver = reference =>
            PdfObjectLookup.TryGet(map, reference, out PdfIndirectObject? target) ? target.Value : null;
        foreach (PdfIndirectObject indirect in map.Values) {
            cancellationToken.ThrowIfCancellationRequested();
            if (indirect.ObjectNumber == encryptObjectNumber) {
                continue;
            }

            if (indirect.Value is PdfStream stream &&
                stream.Dictionary.Get<PdfName>("Type")?.Name == "XRef") {
                continue;
            }

            PdfObject decrypted = decryptor.DecryptObject(indirect.ObjectNumber, indirect.Generation, indirect.Value, cancellationToken);
            replacements.Add(new PdfIndirectObject(indirect.ObjectNumber, indirect.Generation, decrypted));
        }

        for (int i = 0; i < replacements.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            map[replacements[i].ObjectNumber] = replacements[i];
        }
    }
}
