using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeIMO.Html.Qualification;

internal sealed class HtmlQualificationCorpusManifest {
    [JsonPropertyName("$schema")]
    public string Schema { get; set; } = string.Empty;
    public string SchemaVersion { get; set; } = string.Empty;
    public string CorpusId { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string EntryPath { get; set; } = string.Empty;
    public string BaseUri { get; set; } = string.Empty;
    public string SourceKind { get; set; } = string.Empty;
    public string License { get; set; } = string.Empty;
    public List<HtmlQualificationInputFile> Files { get; set; } = new List<HtmlQualificationInputFile>();
    public List<HtmlQualificationSelectorExpectation> SelectorExpectations { get; set; } = new List<HtmlQualificationSelectorExpectation>();
    public List<string> TextMarkers { get; set; } = new List<string>();
    public List<HtmlQualificationProfile> Profiles { get; set; } = new List<HtmlQualificationProfile>();
}

internal sealed class HtmlQualificationInputFile {
    public string Path { get; set; } = string.Empty;
    public string Role { get; set; } = string.Empty;
    public string MediaType { get; set; } = string.Empty;
    public long Length { get; set; }
    public string Sha256 { get; set; } = string.Empty;
}

internal sealed class HtmlQualificationSelectorExpectation {
    public string Selector { get; set; } = string.Empty;
    public int Count { get; set; }
}

internal sealed class HtmlQualificationStyleExpectation {
    public string Selector { get; set; } = string.Empty;
    public string Property { get; set; } = string.Empty;
    public string Contains { get; set; } = string.Empty;
}

internal sealed class HtmlQualificationProfile {
    public string Id { get; set; } = string.Empty;
    public string Mode { get; set; } = string.Empty;
    public string Media { get; set; } = string.Empty;
    public double ViewportWidth { get; set; }
    public double ViewportHeight { get; set; }
    public int ExpectedPageCount { get; set; }
    public List<HtmlQualificationStyleExpectation> StyleExpectations { get; set; } = new List<HtmlQualificationStyleExpectation>();
}

internal sealed class HtmlQualificationCorpus {
    private const string ExpectedSchema = "officeimo.html.qualification-corpus";
    private const string ExpectedSchemaVersion = "1.0";
    private readonly IReadOnlyDictionary<string, byte[]> _bytesByPath;
    private readonly IReadOnlyDictionary<string, HtmlQualificationInputFile> _filesByResourceUri;

    private HtmlQualificationCorpus(
        string rootPath,
        string manifestPath,
        string manifestSha256,
        HtmlQualificationCorpusManifest manifest,
        IReadOnlyDictionary<string, byte[]> bytesByPath,
        IReadOnlyDictionary<string, HtmlQualificationInputFile> filesByResourceUri) {
        RootPath = rootPath;
        ManifestPath = manifestPath;
        ManifestSha256 = manifestSha256;
        Manifest = manifest;
        _bytesByPath = bytesByPath;
        _filesByResourceUri = filesByResourceUri;
    }

    internal string RootPath { get; }
    internal string ManifestPath { get; }
    internal string ManifestSha256 { get; }
    internal HtmlQualificationCorpusManifest Manifest { get; }
    internal Uri BaseUri => new Uri(Manifest.BaseUri, UriKind.Absolute);
    internal string EntryHtml => Encoding.UTF8.GetString(_bytesByPath[NormalizePath(Manifest.EntryPath)]);

    internal static string ResolveDefaultRoot() => Path.Combine(
        AppContext.BaseDirectory, "Documents", "Html", "Qualification", "H0", "v1");

    internal static HtmlQualificationCorpus Load(string? rootPath = null) {
        string root = Path.GetFullPath(rootPath ?? ResolveDefaultRoot());
        if (!Directory.Exists(root)) throw new DirectoryNotFoundException("HTML qualification corpus was not found at " + root + ".");
        string manifestPath = Path.Combine(root, "manifest.json");
        byte[] manifestBytes = File.ReadAllBytes(manifestPath);
        HtmlQualificationCorpusManifest manifest = JsonSerializer.Deserialize<HtmlQualificationCorpusManifest>(
            manifestBytes,
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
            ?? throw new InvalidDataException("HTML qualification corpus manifest is empty.");
        ValidateManifestShape(manifest);

        var bytesByPath = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        var declaredPaths = new HashSet<string>(StringComparer.Ordinal);
        foreach (HtmlQualificationInputFile file in manifest.Files) {
            string normalized = NormalizePath(file.Path);
            if (!declaredPaths.Add(normalized)) throw new InvalidDataException("Duplicate qualification input path: " + normalized + ".");
            string fullPath = ResolveContainedPath(root, normalized);
            byte[] bytes = File.ReadAllBytes(fullPath);
            if (bytes.LongLength != file.Length) {
                throw new InvalidDataException($"Qualification input {normalized} has {bytes.LongLength} bytes; expected {file.Length}.");
            }
            string hash = Hash(bytes);
            if (!string.Equals(hash, file.Sha256, StringComparison.Ordinal)) {
                throw new InvalidDataException($"Qualification input {normalized} has SHA-256 {hash}; expected {file.Sha256}.");
            }
            bytesByPath.Add(normalized, bytes);
        }

        string[] actualPaths = Directory.GetFiles(root, "*", SearchOption.AllDirectories)
            .Select(path => NormalizePath(GetRelativePath(root, path)))
            .Where(path => !string.Equals(path, "manifest.json", StringComparison.Ordinal))
            .OrderBy(path => path, StringComparer.Ordinal)
            .ToArray();
        string[] expectedPaths = declaredPaths.OrderBy(path => path, StringComparer.Ordinal).ToArray();
        if (!actualPaths.SequenceEqual(expectedPaths, StringComparer.Ordinal)) {
            throw new InvalidDataException(
                "Qualification input set differs from the manifest. Actual: " + string.Join(", ", actualPaths) +
                "; expected: " + string.Join(", ", expectedPaths) + ".");
        }

        string entryPath = NormalizePath(manifest.EntryPath);
        if (!bytesByPath.ContainsKey(entryPath)) throw new InvalidDataException("Qualification entryPath is not declared as an input file.");
        var resources = new Dictionary<string, HtmlQualificationInputFile>(StringComparer.Ordinal);
        foreach (HtmlQualificationInputFile file in manifest.Files.Where(IsRenderResource)) {
            string resourceUri = ResolveResourceUri(manifest, file.Path).AbsoluteUri;
            if (resources.ContainsKey(resourceUri)) throw new InvalidDataException("Duplicate qualification resource URI: " + resourceUri + ".");
            resources.Add(resourceUri, file);
        }

        return new HtmlQualificationCorpus(
            root,
            manifestPath,
            Hash(manifestBytes),
            manifest,
            bytesByPath,
            resources);
    }

    internal bool TryResolve(Uri uri, out HtmlQualificationInputFile file, out byte[] bytes) {
        if (_filesByResourceUri.TryGetValue(uri.AbsoluteUri, out HtmlQualificationInputFile? found)) {
            file = found;
            bytes = (byte[])_bytesByPath[NormalizePath(found.Path)].Clone();
            return true;
        }
        file = null!;
        bytes = Array.Empty<byte>();
        return false;
    }

    internal byte[] ReadBytes(HtmlQualificationInputFile file) =>
        (byte[])_bytesByPath[NormalizePath(file.Path)].Clone();

    private static void ValidateManifestShape(HtmlQualificationCorpusManifest manifest) {
        if (!string.Equals(manifest.Schema, ExpectedSchema, StringComparison.Ordinal))
            throw new InvalidDataException("Unsupported HTML qualification corpus schema: " + manifest.Schema + ".");
        if (!string.Equals(manifest.SchemaVersion, ExpectedSchemaVersion, StringComparison.Ordinal))
            throw new InvalidDataException("Unsupported HTML qualification corpus version: " + manifest.SchemaVersion + ".");
        if (string.IsNullOrWhiteSpace(manifest.CorpusId)) throw new InvalidDataException("Qualification corpusId is required.");
        if (string.IsNullOrWhiteSpace(manifest.EntryPath)) throw new InvalidDataException("Qualification entryPath is required.");
        if (!Uri.TryCreate(manifest.BaseUri, UriKind.Absolute, out Uri? baseUri) || baseUri.Scheme is not "http" and not "https")
            throw new InvalidDataException("Qualification baseUri must be an absolute HTTP or HTTPS URI.");
        if (manifest.Files.Count == 0 || manifest.Profiles.Count == 0 || manifest.TextMarkers.Count == 0)
            throw new InvalidDataException("Qualification files, profiles, and textMarkers must be non-empty.");
        if (manifest.Profiles.Select(profile => profile.Id).Distinct(StringComparer.OrdinalIgnoreCase).Count() != manifest.Profiles.Count)
            throw new InvalidDataException("Qualification profile ids must be unique.");
        foreach (HtmlQualificationProfile profile in manifest.Profiles) {
            if (!IsSafeIdentifier(profile.Id))
                throw new InvalidDataException("Qualification profile id must be one safe path segment: " + profile.Id + ".");
        }
        foreach (HtmlQualificationInputFile file in manifest.Files) {
            if (string.IsNullOrWhiteSpace(file.Path) || string.IsNullOrWhiteSpace(file.Role) || string.IsNullOrWhiteSpace(file.MediaType))
                throw new InvalidDataException("Every qualification file requires path, role, and mediaType.");
            if (file.Length <= 0 || file.Sha256.Length != 64 || file.Sha256.Any(character => !Uri.IsHexDigit(character)))
                throw new InvalidDataException("Qualification file length and SHA-256 are invalid for " + file.Path + ".");
        }
    }

    private static bool IsRenderResource(HtmlQualificationInputFile file) =>
        file.Role is "stylesheet" or "font" or "image";

    private static bool IsSafeIdentifier(string value) {
        if (string.IsNullOrWhiteSpace(value) || value.Length > 64 || value is "." or "..") return false;
        for (int index = 0; index < value.Length; index++) {
            char character = value[index];
            if (!char.IsLetterOrDigit(character) && character is not '-' and not '_' and not '.') return false;
        }
        return true;
    }

    private static Uri ResolveResourceUri(HtmlQualificationCorpusManifest manifest, string path) {
        var virtualEntry = new Uri("file:///qualification/" + NormalizePath(manifest.EntryPath));
        var virtualResource = new Uri("file:///qualification/" + NormalizePath(path));
        string relative = Uri.UnescapeDataString(virtualEntry.MakeRelativeUri(virtualResource).ToString());
        return new Uri(new Uri(manifest.BaseUri, UriKind.Absolute), relative);
    }

    private static string ResolveContainedPath(string root, string normalizedPath) {
        string fullPath = Path.GetFullPath(Path.Combine(root, normalizedPath.Replace('/', Path.DirectorySeparatorChar)));
        string prefix = root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
        if (!fullPath.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) || !File.Exists(fullPath))
            throw new InvalidDataException("Qualification input is missing or outside the corpus root: " + normalizedPath + ".");
        return fullPath;
    }

    private static string NormalizePath(string path) {
        string normalized = (path ?? string.Empty).Replace('\\', '/').TrimStart('/');
        if (normalized.Length == 0 || normalized.Split('/').Any(part => part.Length == 0 || part is "." or ".."))
            throw new InvalidDataException("Qualification input path is invalid: " + path + ".");
        return normalized;
    }

    private static string GetRelativePath(string root, string path) {
        var rootUri = new Uri(root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar);
        return Uri.UnescapeDataString(rootUri.MakeRelativeUri(new Uri(path)).ToString()).Replace('/', Path.DirectorySeparatorChar);
    }

    private static string Hash(byte[] bytes) {
        using SHA256 sha = SHA256.Create();
        return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", string.Empty).ToLowerInvariant();
    }
}
