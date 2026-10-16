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
        dictionary.IndexOf("/Filter ", StringComparison.Ordinal) < 0 &&
        (dictionary.StartsWith("<< /Type /EmbeddedFile ", StringComparison.Ordinal) ||
         dictionary.StartsWith("<< /N ", StringComparison.Ordinal) ||
         dictionary.Contains(" /Subtype /Form ", StringComparison.Ordinal));

    internal static bool IsShareableStream(string dictionary) =>
        dictionary.IndexOf("/StructParent", StringComparison.Ordinal) < 0 &&
        (IsAnonymousContentStream(dictionary) ||
         dictionary.StartsWith("<< /N ", StringComparison.Ordinal) ||
         dictionary.Contains(" /Type /Pattern ", StringComparison.Ordinal) ||
         dictionary.Contains(" /FunctionType ", StringComparison.Ordinal) ||
         dictionary.Contains(" /Subtype /Image ", StringComparison.Ordinal) ||
         dictionary.Contains(" /Subtype /Form ", StringComparison.Ordinal));

    // Identical page programs can share storage while each invocation still resolves
    // names and marked-content ids in its own page resource/structure context.
    private static bool IsAnonymousContentStream(string dictionary) =>
        dictionary.StartsWith("<< /Length ", StringComparison.Ordinal) &&
        !dictionary.Contains(" /Type ", StringComparison.Ordinal) &&
        !dictionary.Contains(" /Subtype ", StringComparison.Ordinal) &&
        !dictionary.Contains(" /Length1 ", StringComparison.Ordinal);

    internal static string Hash(string dictionary, byte[]? data = null) {
        using SHA256 hash = SHA256.Create();
        byte[] prefix = PdfEncoding.Latin1GetBytes(dictionary + "\0");
        hash.TransformBlock(prefix, 0, prefix.Length, prefix, 0);
        byte[] payload = data ?? Array.Empty<byte>();
        hash.TransformFinalBlock(payload, 0, payload.Length);
        return Convert.ToBase64String(hash.Hash!);
    }
}
