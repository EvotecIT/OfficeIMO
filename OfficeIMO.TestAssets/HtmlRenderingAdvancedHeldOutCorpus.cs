using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeIMO.Html;

namespace OfficeIMO.Tests;

/// <summary>Frozen, manifest-backed H4/advanced-held-out corpus authored before renderer evaluation.</summary>
internal sealed class HtmlRenderingAdvancedHeldOutCorpus {
    internal const string RelativeRoot = "OfficeIMO.TestAssets/Documents/Html/Qualification/H4/advanced-held-out";
    private const string ExpectedSchema = "officeimo.html.rendering-corpus";
    private const string ExpectedVersion = "2.0";

    private HtmlRenderingAdvancedHeldOutCorpus(
        string rootPath,
        string manifestPath,
        string manifestSha256,
        HtmlRenderingAdvancedHeldOutManifest manifest,
        IReadOnlyList<HtmlRenderingAdvancedHeldOutCase> cases) {
        RootPath = rootPath;
        ManifestPath = manifestPath;
        ManifestSha256 = manifestSha256;
        Manifest = manifest;
        Cases = cases;
    }

    internal string RootPath { get; }
    internal string ManifestPath { get; }
    internal string ManifestSha256 { get; }
    internal HtmlRenderingAdvancedHeldOutManifest Manifest { get; }
    internal IReadOnlyList<HtmlRenderingAdvancedHeldOutCase> Cases { get; }

    internal static HtmlRenderingAdvancedHeldOutCorpus Load(string? rootPath = null) {
        string root = Path.GetFullPath(rootPath ?? ResolveDefaultRoot());
        string manifestPath = Path.Combine(root, "manifest.json");
        byte[] manifestBytes = File.ReadAllBytes(manifestPath);
        HtmlRenderingAdvancedHeldOutManifest manifest = JsonSerializer.Deserialize<HtmlRenderingAdvancedHeldOutManifest>(
            manifestBytes,
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
            ?? throw new InvalidDataException("H4/advanced-held-out rendering corpus manifest is empty.");
        ValidateManifest(manifest);

        var cases = new List<HtmlRenderingAdvancedHeldOutCase>(manifest.Cases.Count);
        var declaredPaths = new HashSet<string>(StringComparer.Ordinal);
        foreach (HtmlRenderingAdvancedHeldOutCaseManifest entry in manifest.Cases) {
            string relativePath = NormalizePath(entry.Path);
            if (!declaredPaths.Add(relativePath)) throw new InvalidDataException("Duplicate H4/advanced-held-out input path: " + relativePath + ".");
            string fullPath = ResolveContainedPath(root, relativePath);
            byte[] bytes = File.ReadAllBytes(fullPath);
            if (bytes.LongLength != entry.Length) {
                throw new InvalidDataException($"H4/advanced-held-out input {relativePath} has {bytes.LongLength} bytes; expected {entry.Length}.");
            }
            string sha256 = Hash(bytes);
            if (!string.Equals(sha256, entry.Sha256, StringComparison.Ordinal)) {
                throw new InvalidDataException($"H4/advanced-held-out input {relativePath} has SHA-256 {sha256}; expected {entry.Sha256}.");
            }
            Encoding encoding = ResolveEncoding(entry.Encoding);
            cases.Add(new HtmlRenderingAdvancedHeldOutCase(entry, bytes, encoding.GetString(bytes)));
        }

        string[] actualInputs = Directory.GetFiles(root, "*.html", SearchOption.AllDirectories)
            .Select(path => NormalizePath(GetRelativePath(root, path)))
            .OrderBy(path => path, StringComparer.Ordinal)
            .ToArray();
        string[] declaredInputs = declaredPaths.OrderBy(path => path, StringComparer.Ordinal).ToArray();
        if (!actualInputs.SequenceEqual(declaredInputs, StringComparer.Ordinal)) {
            throw new InvalidDataException(
                "H4/advanced-held-out input set differs from the frozen manifest. Actual: " + string.Join(", ", actualInputs)
                + "; expected: " + string.Join(", ", declaredInputs) + ".");
        }

        return new HtmlRenderingAdvancedHeldOutCorpus(root, manifestPath, Hash(manifestBytes), manifest, cases.AsReadOnly());
    }

    private static string ResolveDefaultRoot() => Path.Combine(
        AppContext.BaseDirectory, "Documents", "Html", "Qualification", "H4", "advanced-held-out");

    private static void ValidateManifest(HtmlRenderingAdvancedHeldOutManifest manifest) {
        if (!string.Equals(manifest.Schema, ExpectedSchema, StringComparison.Ordinal))
            throw new InvalidDataException("Unsupported H4/advanced-held-out corpus schema: " + manifest.Schema + ".");
        if (!string.Equals(manifest.SchemaVersion, ExpectedVersion, StringComparison.Ordinal))
            throw new InvalidDataException("Unsupported H4/advanced-held-out corpus version: " + manifest.SchemaVersion + ".");
        if (string.IsNullOrWhiteSpace(manifest.CorpusId) || manifest.Cases.Count == 0)
            throw new InvalidDataException("H4/advanced-held-out corpus id and cases are required.");
        if (manifest.RenderIntents.Count != 3 || manifest.OutputFamilies.Count != 3)
            throw new InvalidDataException("H4/advanced-held-out must declare the three frozen render intents and output families.");
        if (manifest.Cases.Select(item => item.Id).Distinct(StringComparer.Ordinal).Count() != manifest.Cases.Count)
            throw new InvalidDataException("H4/advanced-held-out case ids must be unique.");
        foreach (HtmlRenderingAdvancedHeldOutCaseManifest item in manifest.Cases) {
            if (!IsSafeId(item.Id) || item.TextMarkers.Count == 0 || item.Capabilities.Count == 0)
                throw new InvalidDataException("Every H4/advanced-held-out case requires a safe id, text markers, and capabilities.");
            if (item.Length <= 0 || item.Sha256.Length != 64 || item.Sha256.Any(character => !Uri.IsHexDigit(character)))
                throw new InvalidDataException("Invalid H4/advanced-held-out length or SHA-256 for " + item.Id + ".");
            if (item.ExpectedPrintPageCount <= 0)
                throw new InvalidDataException("Expected print page count must be positive for " + item.Id + ".");
        }
    }

    private static Encoding ResolveEncoding(string name) => name.ToLowerInvariant() switch {
        "utf-8" => new UTF8Encoding(false, true),
        "windows-1252" => ResolveWindows1252(),
        _ => throw new InvalidDataException("Unsupported H4/advanced-held-out source encoding: " + name + ".")
    };

    private static Encoding ResolveWindows1252() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        return Encoding.GetEncoding(1252, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
    }

    private static string ResolveContainedPath(string root, string relativePath) {
        string fullPath = Path.GetFullPath(Path.Combine(root, relativePath.Replace('/', Path.DirectorySeparatorChar)));
        string prefix = root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
        if (!fullPath.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) || !File.Exists(fullPath))
            throw new InvalidDataException("H4/advanced-held-out input is missing or outside the corpus root: " + relativePath + ".");
        return fullPath;
    }

    private static string NormalizePath(string path) {
        string value = (path ?? string.Empty).Replace('\\', '/').TrimStart('/');
        if (value.Length == 0 || value.Split('/').Any(part => part.Length == 0 || part is "." or ".."))
            throw new InvalidDataException("Invalid H4/advanced-held-out input path: " + path + ".");
        return value;
    }

    private static string GetRelativePath(string root, string path) {
        var rootUri = new Uri(root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            + Path.DirectorySeparatorChar);
        return Uri.UnescapeDataString(rootUri.MakeRelativeUri(new Uri(path)).ToString())
            .Replace('/', Path.DirectorySeparatorChar);
    }

    private static bool IsSafeId(string value) => value.Length > 0 && value.Length <= 64
        && value.All(character => (character >= 'a' && character <= 'z')
            || (character >= 'A' && character <= 'Z')
            || (character >= '0' && character <= '9')
            || character == '-');

    private static string Hash(byte[] bytes) {
        using SHA256 sha = SHA256.Create();
        return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", string.Empty).ToLowerInvariant();
    }
}

internal sealed class HtmlRenderingAdvancedHeldOutManifest {
    [JsonPropertyName("$schema")]
    public string Schema { get; set; } = string.Empty;
    public string SchemaVersion { get; set; } = string.Empty;
    public string CorpusId { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string SourceKind { get; set; } = string.Empty;
    public string License { get; set; } = string.Empty;
    public DateTimeOffset FrozenAtUtc { get; set; }
    public List<string> RenderIntents { get; set; } = new();
    public List<string> OutputFamilies { get; set; } = new();
    public List<HtmlRenderingAdvancedHeldOutCaseManifest> Cases { get; set; } = new();
}

internal sealed class HtmlRenderingAdvancedHeldOutCaseManifest {
    public string Id { get; set; } = string.Empty;
    public string Path { get; set; } = string.Empty;
    public string Encoding { get; set; } = string.Empty;
    public long Length { get; set; }
    public string Sha256 { get; set; } = string.Empty;
    public int ExpectedPrintPageCount { get; set; }
    public List<string> Capabilities { get; set; } = new();
    public List<string> TextMarkers { get; set; } = new();
}

internal sealed class HtmlRenderingAdvancedHeldOutCase {
    private readonly byte[] _sourceBytes;

    internal HtmlRenderingAdvancedHeldOutCase(HtmlRenderingAdvancedHeldOutCaseManifest manifest, byte[] sourceBytes, string html) {
        Manifest = manifest;
        _sourceBytes = (byte[])sourceBytes.Clone();
        Html = html;
    }

    internal HtmlRenderingAdvancedHeldOutCaseManifest Manifest { get; }
    internal string Id => Manifest.Id;
    internal string SourceRelativePath => HtmlRenderingAdvancedHeldOutCorpus.RelativeRoot + "/" + Manifest.Path;
    internal byte[] SourceBytes => (byte[])_sourceBytes.Clone();
    internal string Html { get; }

    internal HtmlConversionDocument LoadDocument() {
        using var stream = new MemoryStream(_sourceBytes, writable: false);
        return HtmlConversionDocument.Load(stream);
    }
}
