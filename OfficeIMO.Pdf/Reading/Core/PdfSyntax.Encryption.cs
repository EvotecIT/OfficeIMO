namespace OfficeIMO.Pdf;

internal static partial class PdfSyntax {
    internal static bool TryCreateDecryptor(
        Dictionary<int, PdfIndirectObject> map,
        string trailerRaw,
        PdfReadOptions? options,
        out PdfStandardSecurityHandler? decryptor) {
        decryptor = null;
        PdfReference? encryptReference = TryReadFirstReference(trailerRaw, "Encrypt");
        if (encryptReference is null) {
            return false;
        }

        if (!PdfObjectLookup.TryGet(map, encryptReference, out PdfIndirectObject? encryptObject) ||
            encryptObject.Value is not PdfDictionary encryptionDictionary) {
            throw new PdfUnsupportedEncryptionException("PDF encryption dictionary could not be read.");
        }

        byte[] fileId = ReadFirstFileId(map, trailerRaw);
        bool supplied = options != null && options.Password != null;
        decryptor = PdfStandardSecurityHandler.Create(encryptionDictionary, fileId, options?.Password, supplied);
        return true;
    }

    private static byte[] ReadFirstFileId(Dictionary<int, PdfIndirectObject> map, string trailerRaw) {
        foreach (PdfIndirectObject indirect in map.Values) {
            if (indirect.Value is PdfStream stream &&
                stream.Dictionary.Get<PdfName>("Type")?.Name == "XRef" &&
                TryReadFirstFileId(stream.Dictionary, out byte[] fileId)) {
                return fileId;
            }
        }

        return ReadFirstFileId(trailerRaw);
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

    private static byte[] ReadFirstFileId(string trailerRaw) {
        int dictStart = trailerRaw.IndexOf("<<", StringComparison.Ordinal);
        if (dictStart < 0) {
            return Array.Empty<byte>();
        }

        int dictEnd = FindDictEnd(trailerRaw, dictStart, trailerRaw.Length);
        if (dictEnd <= dictStart) {
            return Array.Empty<byte>();
        }

        string dictText = SafeSlice(trailerRaw, dictStart + 2, dictEnd - (dictStart + 2), 1_000_000);
        PdfDictionary trailer = ParseDictionary(dictText);
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
        int encryptObjectNumber) {
        var replacements = new List<PdfIndirectObject>();
        foreach (PdfIndirectObject indirect in map.Values) {
            if (indirect.ObjectNumber == encryptObjectNumber) {
                continue;
            }

            if (indirect.Value is PdfStream stream &&
                stream.Dictionary.Get<PdfName>("Type")?.Name == "XRef") {
                continue;
            }

            PdfObject decrypted = decryptor.DecryptObject(indirect.ObjectNumber, indirect.Generation, indirect.Value);
            replacements.Add(new PdfIndirectObject(indirect.ObjectNumber, indirect.Generation, decrypted));
        }

        for (int i = 0; i < replacements.Count; i++) {
            map[replacements[i].ObjectNumber] = replacements[i];
        }
    }
}
