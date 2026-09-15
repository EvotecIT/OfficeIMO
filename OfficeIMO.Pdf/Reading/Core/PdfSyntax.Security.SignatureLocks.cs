namespace OfficeIMO.Pdf;

internal static partial class PdfSyntax {
    private static PdfSignatureFieldLockInfo? ReadSignatureFieldLockInfo(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary signatureField) {
        if (!signatureField.Items.TryGetValue("Lock", out PdfObject? lockObject) ||
            ResolveObject(objects, lockObject) is not PdfDictionary lockDictionary) {
            return null;
        }

        return ReadFieldLockTransformParameters(objects, lockDictionary);
    }

    private static PdfSignatureFieldLockInfo? ReadSignatureFieldMdpInfo(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary signature) {
        if (!signature.Items.TryGetValue("Reference", out PdfObject? referenceObject)) {
            return null;
        }

        if (ResolveObject(objects, referenceObject) is not PdfArray references) {
            return CreateFailClosedFieldLock();
        }

        PdfSignatureFieldLockInfo? fieldLock = null;
        for (int i = 0; i < references.Items.Count; i++) {
            if (ResolveObject(objects, references.Items[i]) is not PdfDictionary reference) {
                return CreateFailClosedFieldLock();
            }

            string? transformMethod = TryReadName(objects, reference, "TransformMethod");
            if (string.IsNullOrEmpty(transformMethod)) {
                return CreateFailClosedFieldLock();
            }

            if (!string.Equals(transformMethod, "FieldMDP", StringComparison.Ordinal)) {
                continue;
            }

            if (fieldLock is not null ||
                !reference.Items.TryGetValue("TransformParams", out PdfObject? transformParamsObject) ||
                ResolveObject(objects, transformParamsObject) is not PdfDictionary transformParams) {
                return CreateFailClosedFieldLock();
            }

            fieldLock = ReadFieldLockTransformParameters(objects, transformParams);
        }

        return fieldLock;
    }

    private static PdfSignatureFieldLockInfo ReadFieldLockTransformParameters(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary transformParameters) {
        string? action = TryReadName(objects, transformParameters, "Action");
        if (!string.Equals(action, "All", StringComparison.Ordinal) &&
            !string.Equals(action, "Include", StringComparison.Ordinal) &&
            !string.Equals(action, "Exclude", StringComparison.Ordinal)) {
            return CreateFailClosedFieldLock();
        }

        IReadOnlyList<string> fields = Array.Empty<string>();
        if (!string.Equals(action, "All", StringComparison.Ordinal) &&
            !TryReadFieldLockNames(objects, transformParameters, out fields)) {
            return CreateFailClosedFieldLock();
        }

        return new PdfSignatureFieldLockInfo(action, fields);
    }

    private static bool TryReadFieldLockNames(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary transformParameters,
        out IReadOnlyList<string> fields) {
        fields = Array.Empty<string>();
        if (!transformParameters.Items.TryGetValue("Fields", out PdfObject? fieldsObject) ||
            ResolveObject(objects, fieldsObject) is not PdfArray fieldsArray ||
            fieldsArray.Items.Count == 0) {
            return false;
        }

        var values = new List<string>();
        for (int i = 0; i < fieldsArray.Items.Count; i++) {
            string? field = ReadNameOrText(objects, fieldsArray.Items[i]);
            if (string.IsNullOrEmpty(field)) {
                return false;
            }

            if (!values.Contains(field!)) {
                values.Add(field!);
            }
        }

        fields = values.AsReadOnly();
        return true;
    }

    private static PdfSignatureFieldLockInfo CreateFailClosedFieldLock() =>
        new("All", Array.Empty<string>());
}
