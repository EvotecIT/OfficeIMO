using System.Security.Cryptography;

namespace OfficeIMO.Pdf;

/// <summary>Document-scoped content hashes for immutable visual resources; retains no stream payloads.</summary>
internal sealed class PdfVisualResourceCache {
    private readonly Dictionary<string, int> _objects = new(StringComparer.Ordinal);

    internal bool TryGet(string key, out int objectId) => _objects.TryGetValue(key, out objectId);
    internal void Add(string key, int objectId) => _objects.Add(key, objectId);

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
