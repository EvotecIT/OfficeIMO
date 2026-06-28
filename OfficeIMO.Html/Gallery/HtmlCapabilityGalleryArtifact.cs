using System.Security.Cryptography;

namespace OfficeIMO.Html;

/// <summary>
/// Describes one artifact emitted by an HTML capability-gallery scenario.
/// </summary>
public sealed class HtmlCapabilityGalleryArtifact {
    /// <summary>
    /// Creates an artifact descriptor.
    /// </summary>
    /// <param name="id">Stable artifact identifier within the scenario.</param>
    /// <param name="kind">Artifact kind, such as <c>input-html</c>, <c>docx</c>, or <c>roundtrip-html</c>.</param>
    /// <param name="path">Artifact file path.</param>
    /// <param name="mediaType">Artifact media type.</param>
    /// <param name="length">Artifact length in bytes.</param>
    /// <param name="sha256">Lowercase hexadecimal SHA-256 hash of the artifact bytes.</param>
    public HtmlCapabilityGalleryArtifact(string id, string kind, string path, string mediaType, long length, string sha256) {
        Id = id ?? throw new ArgumentNullException(nameof(id));
        Kind = kind ?? throw new ArgumentNullException(nameof(kind));
        Path = path ?? throw new ArgumentNullException(nameof(path));
        MediaType = mediaType ?? throw new ArgumentNullException(nameof(mediaType));
        Length = length;
        Sha256 = sha256 ?? throw new ArgumentNullException(nameof(sha256));
    }

    /// <summary>
    /// Creates an artifact descriptor from an existing file.
    /// </summary>
    /// <param name="id">Stable artifact identifier within the scenario.</param>
    /// <param name="kind">Artifact kind, such as <c>input-html</c>, <c>docx</c>, or <c>roundtrip-html</c>.</param>
    /// <param name="path">Artifact file path.</param>
    /// <param name="mediaType">Artifact media type.</param>
    /// <returns>A descriptor containing file length and SHA-256 hash.</returns>
    public static HtmlCapabilityGalleryArtifact FromFile(string id, string kind, string path, string mediaType) {
        if (path == null) {
            throw new ArgumentNullException(nameof(path));
        }

        var info = new FileInfo(path);
        using FileStream stream = File.OpenRead(path);
        using SHA256 sha = SHA256.Create();
        byte[] hash = sha.ComputeHash(stream);
        return new HtmlCapabilityGalleryArtifact(id, kind, path, mediaType, info.Length, ToHex(hash));
    }

    /// <summary>
    /// Writes text content to an artifact file and returns its descriptor.
    /// </summary>
    /// <param name="id">Stable artifact identifier within the scenario.</param>
    /// <param name="kind">Artifact kind, such as <c>semantic-html</c> or <c>manifest-json</c>.</param>
    /// <param name="path">Artifact file path.</param>
    /// <param name="mediaType">Artifact media type.</param>
    /// <param name="content">Text content to write.</param>
    /// <returns>A descriptor containing file length and SHA-256 hash.</returns>
    public static HtmlCapabilityGalleryArtifact WriteTextFile(string id, string kind, string path, string mediaType, string content) {
        if (path == null) {
            throw new ArgumentNullException(nameof(path));
        }

        string? directory = System.IO.Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(directory)) {
            Directory.CreateDirectory(directory);
        }

        File.WriteAllText(path, content ?? string.Empty, Encoding.UTF8);
        return FromFile(id, kind, path, mediaType);
    }

    /// <summary>
    /// Stable artifact identifier within the scenario.
    /// </summary>
    public string Id { get; }

    /// <summary>
    /// Artifact kind, such as <c>input-html</c>, <c>docx</c>, or <c>roundtrip-html</c>.
    /// </summary>
    public string Kind { get; }

    /// <summary>
    /// Artifact file path.
    /// </summary>
    public string Path { get; }

    /// <summary>
    /// Artifact media type.
    /// </summary>
    public string MediaType { get; }

    /// <summary>
    /// Artifact length in bytes.
    /// </summary>
    public long Length { get; }

    /// <summary>
    /// Lowercase hexadecimal SHA-256 hash of the artifact bytes.
    /// </summary>
    public string Sha256 { get; }

    private static string ToHex(byte[] bytes) {
        var builder = new StringBuilder(bytes.Length * 2);
        foreach (byte value in bytes) {
            builder.Append(value.ToString("x2"));
        }

        return builder.ToString();
    }
}
