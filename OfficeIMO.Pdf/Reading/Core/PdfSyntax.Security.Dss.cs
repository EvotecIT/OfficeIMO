namespace OfficeIMO.Pdf;

internal static partial class PdfSyntax {
    private static PdfDocumentDssInfo ReadDocumentSecurityStoreInfo(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary catalog) {
        if (!catalog.Items.TryGetValue("DSS", out PdfObject? dssObject)) {
            return PdfDocumentDssInfo.Empty;
        }

        int? objectNumber = dssObject is PdfReference reference ? reference.ObjectNumber : null;
        if (ResolveObject(objects, dssObject) is not PdfDictionary dss) {
            return new PdfDocumentDssInfo(
                true,
                objectNumber,
                Array.Empty<string>(),
                Array.Empty<int>(),
                Array.Empty<int>(),
                Array.Empty<int>(),
                Array.Empty<int>(),
                Array.Empty<int>(),
                Array.Empty<int>(),
                Array.Empty<int>());
        }

        var vriKeys = new List<string>();
        var vriCerts = new List<int>();
        var vriOcsps = new List<int>();
        var vriCrls = new List<int>();
        var timestamps = new List<int>();
        ReadVriEvidence(objects, dss, vriKeys, vriCerts, vriOcsps, vriCrls, timestamps);

        return new PdfDocumentDssInfo(
            true,
            objectNumber,
            ToReadOnly(vriKeys),
            ReadReferenceArrayObjectNumbers(objects, dss, "Certs"),
            ReadReferenceArrayObjectNumbers(objects, dss, "OCSPs"),
            ReadReferenceArrayObjectNumbers(objects, dss, "CRLs"),
            ToReadOnly(vriCerts),
            ToReadOnly(vriOcsps),
            ToReadOnly(vriCrls),
            ToReadOnly(timestamps));
    }

    private static void ReadVriEvidence(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary dss,
        List<string> vriKeys,
        List<int> certs,
        List<int> ocsps,
        List<int> crls,
        List<int> timestamps) {
        if (!dss.Items.TryGetValue("VRI", out PdfObject? vriObject) ||
            ResolveObject(objects, vriObject) is not PdfDictionary vri) {
            return;
        }

        foreach (var entry in vri.Items.OrderBy(static item => item.Key, StringComparer.Ordinal)) {
            if (!string.IsNullOrEmpty(entry.Key) && !vriKeys.Contains(entry.Key)) {
                vriKeys.Add(entry.Key);
            }

            if (ResolveObject(objects, entry.Value) is not PdfDictionary vriEntry) {
                continue;
            }

            AddReferenceArrayObjectNumbers(objects, vriEntry, "Cert", certs);
            AddReferenceArrayObjectNumbers(objects, vriEntry, "OCSP", ocsps);
            AddReferenceArrayObjectNumbers(objects, vriEntry, "CRL", crls);
            AddSingleReferenceObjectNumber(vriEntry, "TS", timestamps);
        }
    }

    private static IReadOnlyList<int> ReadReferenceArrayObjectNumbers(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary dictionary,
        string key) {
        var objectNumbers = new List<int>();
        AddReferenceArrayObjectNumbers(objects, dictionary, key, objectNumbers);
        return ToReadOnly(objectNumbers);
    }

    private static void AddReferenceArrayObjectNumbers(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary dictionary,
        string key,
        List<int> objectNumbers) {
        if (!dictionary.Items.TryGetValue(key, out PdfObject? value)) {
            return;
        }

        PdfObject? resolved = ResolveObject(objects, value);
        if (resolved is PdfArray array) {
            for (int i = 0; i < array.Items.Count; i++) {
                AddReferenceObjectNumber(array.Items[i], objectNumbers);
            }

            return;
        }

        AddReferenceObjectNumber(value, objectNumbers);
    }

    private static void AddSingleReferenceObjectNumber(PdfDictionary dictionary, string key, List<int> objectNumbers) {
        if (dictionary.Items.TryGetValue(key, out PdfObject? value)) {
            AddReferenceObjectNumber(value, objectNumbers);
        }
    }

    private static void AddReferenceObjectNumber(PdfObject? value, List<int> objectNumbers) {
        if (value is PdfReference reference && !objectNumbers.Contains(reference.ObjectNumber)) {
            objectNumbers.Add(reference.ObjectNumber);
        }
    }

    private static IReadOnlyList<T> ToReadOnly<T>(List<T> values) {
        return values.Count == 0 ? Array.Empty<T>() : values.AsReadOnly();
    }
}
