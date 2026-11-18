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
        PdfDictionary? encryptionDictionary;
        if (encryptReference is null) {
            encryptionDictionary = ReadDirectTrailerEncryptionDictionary(trailerRaw, options?.Limits, cancellationToken);
            if (encryptionDictionary is null) return false;
        } else if (!PdfObjectLookup.TryGet(map, encryptReference, out PdfIndirectObject? encryptObject) ||
            (encryptionDictionary = encryptObject.Value as PdfDictionary) is null) {
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
    /// Returns the encryption dictionary with indirect values resolved in its own entries, in /CF and in each crypt
    /// filter dictionary. ISO 32000-2 7.3.10 lets /CF, the crypt filter dictionaries and their entries be indirect
    /// objects; strings in the encryption dictionary must be direct, and a stream or an unresolvable reference is left
    /// as stored. The security handler reads these values before the document's objects are decrypted.
    /// </summary>
    private static PdfDictionary ResolveEncryptionDictionaryReferences(
        Dictionary<int, PdfIndirectObject> map,
        PdfDictionary encryptionDictionary,
        CancellationToken cancellationToken) {
        if (!NeedsResolution(encryptionDictionary)) return encryptionDictionary;

        // Each reference chain is walked once for the whole dictionary, and a crypt filter dictionary shared by
        // several names is copied once, so the work is linear in the entries and the objects they reach.
        var resolvedReferences = new Dictionary<(int ObjectNumber, int Generation), PdfObject?>();
        var normalizedFilters = new Dictionary<PdfDictionary, PdfDictionary>(); // PdfDictionary compares by identity.
        PdfDictionary resolved = ResolveEntries(encryptionDictionary);
        if (resolved.Get<PdfDictionary>("CF") is PdfDictionary filters) {
            PdfDictionary resolvedFilters = ResolveEntries(filters);
            foreach (string name in resolvedFilters.Items.Keys.ToList()) {
                cancellationToken.ThrowIfCancellationRequested();
                if (resolvedFilters.Items[name] is not PdfDictionary filter) continue;
                // A dictionary shared by several names is scanned and, when needed, copied once.
                if (!normalizedFilters.TryGetValue(filter, out PdfDictionary? normalized)) {
                    normalized = ContainsReference(filter) ? ResolveEntries(filter) : filter;
                    normalizedFilters.Add(filter, normalized);
                }

                resolvedFilters.Items[name] = normalized;
            }

            resolved.Items["CF"] = resolvedFilters;
        }

        return resolved;

        static bool NeedsResolution(PdfDictionary dictionary) {
            if (ContainsReference(dictionary)) return true;
            if (dictionary.Get<PdfDictionary>("CF") is not PdfDictionary filters) return false;
            if (ContainsReference(filters)) return true;
            foreach (PdfObject filter in filters.Items.Values) {
                if (filter is PdfDictionary filterDictionary && ContainsReference(filterDictionary)) return true;
            }

            return false;
        }

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
                if (value is PdfReference reference && Resolve(reference) is PdfObject target && target is not PdfStream) {
                    value = target;
                }

                copy.Items[entry.Key] = value;
            }

            return copy;
        }

        // Follows a reference chain to its first non-reference value. Every reference on the way records the result,
        // so a later entry that reaches the same chain stops at the first recorded reference. A cycle or a missing
        // object resolves to null.
        PdfObject? Resolve(PdfReference reference) {
            var walked = new List<(int ObjectNumber, int Generation)>();
            var onChain = new HashSet<(int ObjectNumber, int Generation)>();
            PdfObject? result = null;
            PdfReference current = reference;
            while (true) {
                cancellationToken.ThrowIfCancellationRequested();
                var key = (current.ObjectNumber, current.Generation);
                if (resolvedReferences.TryGetValue(key, out PdfObject? known)) { result = known; break; }
                if (!onChain.Add(key)) break;
                walked.Add(key);
                if (!PdfObjectLookup.TryGet(map, current, out PdfIndirectObject? indirect)) break;
                if (indirect.Value is PdfReference next) { current = next; continue; }
                result = indirect.Value;
                break;
            }

            foreach (var key in walked) resolvedReferences[key] = result;
            return result;
        }
    }

    /// <summary>
    /// An encryption dictionary written directly in the trailer instead of as an indirect object. Its strings are
    /// never encrypted, because only indirect objects are.
    /// </summary>
    internal static PdfDictionary? ReadDirectTrailerEncryptionDictionary(string? trailerRaw, PdfReadLimits? limits,
        CancellationToken cancellationToken) {
        if (string.IsNullOrEmpty(trailerRaw)) return null;
        int searchIndex = 0;
        PdfReadLimits effectiveLimits = limits ?? new PdfReadLimits();
        while (TryReadNextTrailerDictionary(trailerRaw!, ref searchIndex, effectiveLimits, out PdfDictionary dictionary, cancellationToken)) {
            if (dictionary.Items.TryGetValue("Encrypt", out PdfObject? value)) return value as PdfDictionary;
        }

        return null;
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
