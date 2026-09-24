using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfSyntax {
    private static PdfDocumentDssInfo ReadDocumentSecurityStoreInfo(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary catalog,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
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
        ReadVriEvidence(objects, dss, vriKeys, vriCerts, vriOcsps, vriCrls, timestamps, cancellationToken);

        return new PdfDocumentDssInfo(
            true,
            objectNumber,
            ToReadOnly(vriKeys),
            ReadReferenceArrayObjectNumbers(objects, dss, "Certs", cancellationToken),
            ReadReferenceArrayObjectNumbers(objects, dss, "OCSPs", cancellationToken),
            ReadReferenceArrayObjectNumbers(objects, dss, "CRLs", cancellationToken),
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
        List<int> timestamps,
        CancellationToken cancellationToken) {
        if (!dss.Items.TryGetValue("VRI", out PdfObject? vriObject) ||
            ResolveObject(objects, vriObject) is not PdfDictionary vri) {
            return;
        }

        var ordered = new List<KeyValuePair<string, PdfObject>>(vri.Items.Count);
        foreach (var entry in vri.Items) {
            cancellationToken.ThrowIfCancellationRequested();
            ordered.Add(entry);
        }
        try {
            ordered.Sort((left, right) => {
                cancellationToken.ThrowIfCancellationRequested();
                return PdfStringComparison.CompareOrdinal(left.Key, right.Key, cancellationToken);
            });
        } catch (InvalidOperationException) when (cancellationToken.IsCancellationRequested) {
            cancellationToken.ThrowIfCancellationRequested();
            throw;
        }
        var seenCerts = new HashSet<int>();
        var seenOcsps = new HashSet<int>();
        var seenCrls = new HashSet<int>();
        var seenTimestamps = new HashSet<int>();
        foreach (var entry in ordered) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!string.IsNullOrEmpty(entry.Key)) {
                vriKeys.Add(entry.Key);
            }

            if (ResolveObject(objects, entry.Value) is not PdfDictionary vriEntry) {
                continue;
            }

            AddReferenceArrayObjectNumbers(objects, vriEntry, "Cert", certs, seenCerts, cancellationToken);
            AddReferenceArrayObjectNumbers(objects, vriEntry, "OCSP", ocsps, seenOcsps, cancellationToken);
            AddReferenceArrayObjectNumbers(objects, vriEntry, "CRL", crls, seenCrls, cancellationToken);
            AddSingleReferenceObjectNumber(vriEntry, "TS", timestamps, seenTimestamps);
        }
    }

    private static IReadOnlyList<int> ReadReferenceArrayObjectNumbers(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary dictionary,
        string key,
        CancellationToken cancellationToken) {
        var objectNumbers = new List<int>();
        AddReferenceArrayObjectNumbers(objects, dictionary, key, objectNumbers, new HashSet<int>(), cancellationToken);
        return ToReadOnly(objectNumbers);
    }

    private static void AddReferenceArrayObjectNumbers(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary dictionary,
        string key,
        List<int> objectNumbers,
        HashSet<int> seen,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!dictionary.Items.TryGetValue(key, out PdfObject? value)) {
            return;
        }

        PdfObject? resolved = ResolveObject(objects, value);
        if (resolved is PdfArray array) {
            for (int i = 0; i < array.Items.Count; i++) {
                cancellationToken.ThrowIfCancellationRequested();
                AddReferenceObjectNumber(array.Items[i], objectNumbers, seen);
            }

            return;
        }

        AddReferenceObjectNumber(value, objectNumbers, seen);
    }

    private static void AddSingleReferenceObjectNumber(PdfDictionary dictionary, string key, List<int> objectNumbers, HashSet<int> seen) {
        if (dictionary.Items.TryGetValue(key, out PdfObject? value)) {
            AddReferenceObjectNumber(value, objectNumbers, seen);
        }
    }

    private static void AddReferenceObjectNumber(PdfObject? value, List<int> objectNumbers, HashSet<int> seen) {
        if (value is PdfReference reference && seen.Add(reference.ObjectNumber)) {
            objectNumbers.Add(reference.ObjectNumber);
        }
    }

    private static IReadOnlyList<T> ToReadOnly<T>(List<T> values) {
        return values.Count == 0 ? Array.Empty<T>() : values.AsReadOnly();
    }
}
