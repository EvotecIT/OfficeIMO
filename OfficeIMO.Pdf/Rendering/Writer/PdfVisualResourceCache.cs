using System.Security.Cryptography;

namespace OfficeIMO.Pdf;

/// <summary>
/// Document-scoped lookup of immutable visual resources. Streams the object store still holds in memory are
/// matched by their dictionary (which carries /Length) and then byte for byte, keeping only references the
/// store already holds; other streams are keyed by a content hash, so no payload is retained.
/// </summary>
internal sealed class PdfVisualResourceCache {
    private const int StreamCandidates = 8;
    private readonly Dictionary<string, int> _objects = new(StringComparer.Ordinal);
    private readonly Dictionary<string, List<KeyValuePair<byte[], int>>> _streams = new(StringComparer.Ordinal);
    private readonly Dictionary<string, KeyValuePair<byte[], int>> _overflowStreams = new(StringComparer.Ordinal);

    internal bool TryGet(string key, out int objectId) => _objects.TryGetValue(key, out objectId);
    internal void Add(string key, int objectId) => _objects.Add(key, objectId);

    internal bool TryGetStream(string dictionary, byte[] data, out int objectId) {
        if (_streams.TryGetValue(dictionary, out List<KeyValuePair<byte[], int>>? candidates)) {
            foreach (KeyValuePair<byte[], int> candidate in candidates) {
                if (candidate.Key.Length == data.Length && candidate.Key.SequenceEqual(data)) {
                    objectId = candidate.Value;
                    return true;
                }
            }
            if (candidates.Count == StreamCandidates &&
                _overflowStreams.TryGetValue(Hash(dictionary, data), out KeyValuePair<byte[], int> hashed) &&
                hashed.Key.SequenceEqual(data)) {
                objectId = hashed.Value;
                return true;
            }
        }
        objectId = 0;
        return false;
    }

    // A bounded candidate list keeps many same-length streams from turning every lookup into a scan;
    // streams past it are keyed by content hash instead.
    internal void AddStream(string dictionary, byte[] data, int objectId) {
        if (!_streams.TryGetValue(dictionary, out List<KeyValuePair<byte[], int>>? candidates)) {
            candidates = new List<KeyValuePair<byte[], int>>(1);
            _streams.Add(dictionary, candidates);
        }
        if (candidates.Count < StreamCandidates) candidates.Add(new KeyValuePair<byte[], int>(data, objectId));
        else _overflowStreams[Hash(dictionary, data)] = new KeyValuePair<byte[], int>(data, objectId);
    }

    internal static bool IsVisualDictionary(string body) =>
        body.StartsWith("<< /Type /ExtGState ", StringComparison.Ordinal) ||
        body.StartsWith("<< /ShadingType ", StringComparison.Ordinal) ||
        body.StartsWith("<< /FunctionType ", StringComparison.Ordinal);

    internal static bool ShouldCompressStream(string dictionary) =>
        !ContainsOrdinal(dictionary, "/Filter ") &&
        (dictionary.StartsWith("<< /Type /EmbeddedFile ", StringComparison.Ordinal) ||
         dictionary.StartsWith("<< /N ", StringComparison.Ordinal) ||
         ContainsOrdinal(dictionary, " /Subtype /Form "));

    internal static bool IsShareableStream(string dictionary) =>
        !ContainsOrdinal(dictionary, "/StructParent") &&
        (IsAnonymousContentStream(dictionary) ||
         dictionary.StartsWith("<< /N ", StringComparison.Ordinal) ||
         ContainsOrdinal(dictionary, " /Type /Pattern ") ||
         ContainsOrdinal(dictionary, " /FunctionType ") ||
         ContainsOrdinal(dictionary, " /Subtype /Image ") ||
         ContainsOrdinal(dictionary, " /Subtype /Form "));

    // Identical page programs can share storage while each invocation still resolves
    // names and marked-content ids in its own page resource/structure context.
    private static bool IsAnonymousContentStream(string dictionary) =>
        dictionary.StartsWith("<< /Length ", StringComparison.Ordinal) &&
        !ContainsOrdinal(dictionary, " /Type ") &&
        !ContainsOrdinal(dictionary, " /Subtype ") &&
        !ContainsOrdinal(dictionary, " /Length1 ");

    private static bool ContainsOrdinal(string value, string candidate) {
#if NETSTANDARD2_0 || NETFRAMEWORK
        return value.IndexOf(candidate, StringComparison.Ordinal) >= 0;
#else
        return value.Contains(candidate, StringComparison.Ordinal);
#endif
    }

    internal static string Hash(string dictionary, byte[]? data = null) {
        using SHA256 hash = SHA256.Create();
        byte[] prefix = PdfEncoding.Latin1GetBytes(dictionary + "\0");
        hash.TransformBlock(prefix, 0, prefix.Length, prefix, 0);
        byte[] payload = data ?? Array.Empty<byte>();
        hash.TransformFinalBlock(payload, 0, payload.Length);
        return Convert.ToBase64String(hash.Hash!);
    }
}
