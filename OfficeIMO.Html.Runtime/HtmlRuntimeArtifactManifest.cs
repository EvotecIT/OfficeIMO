using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Html.Runtime;

/// <summary>One content-addressed entry in a deterministic runtime artifact manifest.</summary>
public sealed class HtmlRuntimeArtifactEntry {
    /// <summary>Stable logical entry name. Resource URLs are deliberately omitted.</summary>
    public string Name { get; init; } = string.Empty;
    /// <summary>Entry media type.</summary>
    public string ContentType { get; init; } = string.Empty;
    /// <summary>Exact entry byte count.</summary>
    public long ByteCount { get; init; }
    /// <summary>Lowercase SHA-256 content digest.</summary>
    public string Sha256 { get; init; } = string.Empty;
}

/// <summary>Deterministic content manifest for one captured runtime document.</summary>
public sealed class HtmlRuntimeArtifactManifest {
    /// <summary>Manifest schema version.</summary>
    public int SchemaVersion { get; init; } = 1;
    /// <summary>Stable artifact kind.</summary>
    public string Kind { get; init; } = "html-runtime-capture";
    /// <summary>Content-addressed artifact identity derived from ordered entry metadata.</summary>
    public string Id { get; init; } = string.Empty;
    /// <summary>Total content bytes represented by the manifest.</summary>
    public long ByteCount { get; init; }
    /// <summary>Ordered manifest entries.</summary>
    public IReadOnlyList<HtmlRuntimeArtifactEntry> Entries { get; init; } = Array.Empty<HtmlRuntimeArtifactEntry>();

    internal static HtmlRuntimeArtifactManifest Create(HtmlScriptCapture capture) {
        byte[] document = Encoding.UTF8.GetBytes(DocumentHtml(capture));
        var entries = new List<HtmlRuntimeArtifactEntry> { Entry("document.html", "text/html; charset=utf-8", document) };
        int index = 0;
        foreach (HtmlRuntimeResource resource in capture.Resources
                     .OrderBy(item => HtmlRuntimeResourcePolicy.Key(item.Url), StringComparer.Ordinal)) {
            entries.Add(Entry($"resource-{++index:D4}", resource.ContentType, resource.Content));
        }
        return new HtmlRuntimeArtifactManifest {
            Id = "sha256:" + ManifestHash(entries),
            ByteCount = entries.Sum(item => item.ByteCount),
            Entries = Array.AsReadOnly(entries.ToArray())
        };
    }

    internal static string DocumentHtml(HtmlScriptCapture capture) =>
        capture.Document.DocumentElement?.OuterHtml ?? string.Empty;

    private static HtmlRuntimeArtifactEntry Entry(string name, string contentType, byte[] content) => new() {
        Name = name,
        ContentType = contentType,
        ByteCount = content.LongLength,
        Sha256 = Hash(content)
    };

    private static string ManifestHash(IReadOnlyList<HtmlRuntimeArtifactEntry> entries) {
        using var stream = new MemoryStream();
        using (var writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true)) {
            writer.Write(1);
            writer.Write("html-runtime-capture");
            writer.Write(entries.Count);
            foreach (HtmlRuntimeArtifactEntry item in entries) {
                writer.Write(item.Name);
                writer.Write(item.ContentType);
                writer.Write(item.ByteCount);
                writer.Write(item.Sha256);
            }
        }
        stream.Position = 0;
        return Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
    }

    private static string Hash(byte[] content) => Convert.ToHexString(SHA256.HashData(content)).ToLowerInvariant();
}
