using System.Security.Cryptography;
using OfficeIMO.Pdf.Filters;
using OfficeIMO.Security;

namespace OfficeIMO.Pdf;

/// <summary>
/// Decoded Document Security Store values (ISO 32000-2 section 12.8.4.3) and the proof-of-existence time each
/// revocation value gains from the earliest valid document time-stamp whose revision contains it.
/// </summary>
internal sealed class PdfDocumentValidationData {
    private readonly Dictionary<string, DateTimeOffset> _proofOfExistence;

    private PdfDocumentValidationData(
        IReadOnlyList<byte[]> certificates,
        IReadOnlyList<Value> revocation,
        Dictionary<string, DateTimeOffset> proofOfExistence) {
        Certificates = certificates;
        Revocation = revocation;
        _proofOfExistence = proofOfExistence;
    }

    internal static PdfDocumentValidationData Empty { get; } =
        new PdfDocumentValidationData(Array.Empty<byte[]>(), Array.Empty<Value>(), new Dictionary<string, DateTimeOffset>(StringComparer.Ordinal));

    internal IReadOnlyList<byte[]> Certificates { get; }

    internal IReadOnlyList<Value> Revocation { get; }

    internal IEnumerable<string> RevocationHashes => Revocation.Select(static value => value.Hash);

    internal IReadOnlyList<OfficeRevocationToken> CreateTokens() =>
        Revocation.Select(value => new OfficeRevocationToken(
            value.Encoded,
            value.IsOcsp,
            _proofOfExistence.TryGetValue(value.Hash, out DateTimeOffset time) ? time : null)).ToArray();

    internal PdfDocumentValidationData WithProofOfExistence(IReadOnlyDictionary<string, DateTimeOffset> proofOfExistence) {
        var times = new Dictionary<string, DateTimeOffset>(StringComparer.Ordinal);
        foreach (Value value in Revocation) {
            if (proofOfExistence.TryGetValue(value.Hash, out DateTimeOffset time)) times[value.Hash] = time;
        }
        return new PdfDocumentValidationData(Certificates, Revocation, times);
    }

    internal static PdfDocumentValidationData Read(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDocumentDssInfo store,
        int maxDecodedBytes,
        CancellationToken cancellationToken) {
        if (!store.HasValidationEvidence) return Empty;
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var certificates = new List<byte[]>();
        foreach (int objectNumber in store.CertificateObjectNumbers.Concat(store.VriCertificateObjectNumbers)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (TryDecode(objects, objectNumber, maxDecodedBytes, out byte[] value) && seen.Add(Hash(value))) certificates.Add(value);
        }

        var revocation = new List<Value>();
        AddRevocation(objects, store.OcspObjectNumbers.Concat(store.VriOcspObjectNumbers), isOcsp: true, maxDecodedBytes, seen, revocation, cancellationToken);
        AddRevocation(objects, store.CrlObjectNumbers.Concat(store.VriCrlObjectNumbers), isOcsp: false, maxDecodedBytes, seen, revocation, cancellationToken);
        return new PdfDocumentValidationData(certificates, revocation, new Dictionary<string, DateTimeOffset>(StringComparer.Ordinal));
    }

    private static void AddRevocation(
        Dictionary<int, PdfIndirectObject> objects,
        IEnumerable<int> objectNumbers,
        bool isOcsp,
        int maxDecodedBytes,
        HashSet<string> seen,
        List<Value> target,
        CancellationToken cancellationToken) {
        foreach (int objectNumber in objectNumbers) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!TryDecode(objects, objectNumber, maxDecodedBytes, out byte[] value)) continue;
            string hash = Hash(value);
            if (seen.Add(hash)) target.Add(new Value(value, isOcsp, hash));
        }
    }

    private static bool TryDecode(Dictionary<int, PdfIndirectObject> objects, int objectNumber, int maxDecodedBytes, out byte[] value) {
        value = Array.Empty<byte>();
        if (!objects.TryGetValue(objectNumber, out PdfIndirectObject? indirect) || indirect.Value is not PdfStream stream) return false;
        if (stream.DecodingFailed || !StreamDecoder.TryDecode(stream, maxDecodedBytes, out byte[] decoded, objects)) return false;
        if (decoded.Length == 0) return false;
        value = decoded;
        return true;
    }

    internal static string Hash(byte[] value) {
#if NET8_0_OR_GREATER
        return Convert.ToBase64String(SHA256.HashData(value));
#else
        using SHA256 sha256 = SHA256.Create();
        return Convert.ToBase64String(sha256.ComputeHash(value));
#endif
    }

    internal sealed class Value {
        internal Value(byte[] encoded, bool isOcsp, string hash) {
            Encoded = encoded;
            IsOcsp = isOcsp;
            Hash = hash;
        }

        internal byte[] Encoded { get; }
        internal bool IsOcsp { get; }
        internal string Hash { get; }
    }
}
