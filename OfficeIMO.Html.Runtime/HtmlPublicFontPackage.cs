using System.Security.Cryptography;
using System.Text.Json;

namespace OfficeIMO.Html.Runtime;

internal sealed class HtmlPublicFontPackageIdentity {
    public string Id { get; set; } = string.Empty;
    public string ManifestSha256 { get; set; } = string.Empty;
    public string FilesSha256 { get; set; } = string.Empty;
    public string ManifestJson { get; set; } = string.Empty;
    public HtmlPublicFontFileIdentity[] Fonts { get; set; } = Array.Empty<HtmlPublicFontFileIdentity>();
    public HtmlPublicFontFileIdentity[] Licenses { get; set; } = Array.Empty<HtmlPublicFontFileIdentity>();
}

internal sealed class HtmlPublicFontFileIdentity {
    public string Name { get; set; } = string.Empty;
    public long ByteCount { get; set; }
    public string Sha256 { get; set; } = string.Empty;
}

internal static class HtmlPublicFontPackage {
    internal const string DirectoryName = "font-pack";
    internal const string ManifestName = "font-pack.json";

    internal static HtmlPublicFontPackageIdentity Load(string rendererDirectory) {
        ArgumentException.ThrowIfNullOrWhiteSpace(rendererDirectory);
        string directory = Path.Combine(Path.GetFullPath(rendererDirectory), DirectoryName);
        string manifestPath = Path.Combine(directory, ManifestName);
        if (!File.Exists(manifestPath))
            throw new HtmlScriptRuntimeException("The published renderer font-package manifest is missing.");
        byte[] manifestBytes = File.ReadAllBytes(manifestPath);
        if (manifestBytes.Length is 0 or > 128 * 1024)
            throw new HtmlScriptRuntimeException("The published renderer font-package manifest is invalid.");
        string manifestJson = System.Text.Encoding.UTF8.GetString(manifestBytes);
        using JsonDocument manifest = JsonDocument.Parse(manifestBytes, new JsonDocumentOptions {
            CommentHandling = JsonCommentHandling.Disallow,
            AllowTrailingCommas = false,
            MaxDepth = 32
        });
        JsonElement root = manifest.RootElement;
        string id = RequiredText(root, "id", 128);
        var declaredLicenses = new Dictionary<string, string>(StringComparer.Ordinal);
        if (!root.TryGetProperty("licenses", out JsonElement licenseEntries) || licenseEntries.ValueKind != JsonValueKind.Array)
            throw new HtmlScriptRuntimeException("The published renderer font-package manifest has no licenses array.");
        foreach (JsonElement license in licenseEntries.EnumerateArray()) {
            string name = RequiredFileName(license, "name");
            string sha256 = RequiredSha256(license, "sha256");
            if (!name.StartsWith("OFL-", StringComparison.Ordinal) || !name.EndsWith(".txt", StringComparison.Ordinal)
                || !declaredLicenses.TryAdd(name, sha256))
                throw new HtmlScriptRuntimeException("The font-package manifest contains an invalid license file.");
        }
        if (declaredLicenses.Count == 0 || declaredLicenses.Count > 16)
            throw new HtmlScriptRuntimeException("The published renderer font-package license count is invalid.");
        var declared = new Dictionary<string, string>(StringComparer.Ordinal);
        var referencedLicenses = new HashSet<string>(StringComparer.Ordinal);
        if (!root.TryGetProperty("fonts", out JsonElement fonts) || fonts.ValueKind != JsonValueKind.Array)
            throw new HtmlScriptRuntimeException("The published renderer font-package manifest has no fonts array.");
        foreach (JsonElement font in fonts.EnumerateArray()) {
            _ = RequiredText(font, "license", 128);
            string licenseFile = RequiredFileName(font, "licenseFile");
            if (!declaredLicenses.ContainsKey(licenseFile))
                throw new HtmlScriptRuntimeException("A font-package entry references an undeclared license file.");
            referencedLicenses.Add(licenseFile);
            if (!font.TryGetProperty("files", out JsonElement files) || files.ValueKind != JsonValueKind.Array)
                throw new HtmlScriptRuntimeException("A font-package entry has no files array.");
            foreach (JsonElement file in files.EnumerateArray()) {
                string name = RequiredFileName(file, "name");
                string sha256 = RequiredSha256(file, "sha256");
                if (!declared.TryAdd(name, sha256))
                    throw new HtmlScriptRuntimeException("The font-package manifest contains a duplicate file name.");
            }
        }
        if (declared.Count == 0 || declared.Count > 64)
            throw new HtmlScriptRuntimeException("The published renderer font-package file count is invalid.");
        if (!referencedLicenses.SetEquals(declaredLicenses.Keys))
            throw new HtmlScriptRuntimeException("The font-package manifest contains an unreferenced license file.");

        HtmlPublicFontFileIdentity[] fontFiles = declared.OrderBy(item => item.Key, StringComparer.Ordinal)
            .Select(item => VerifyFile(directory, item.Key, item.Value)).ToArray();
        string[] actualFonts = Directory.EnumerateFiles(directory)
            .Where(path => Path.GetExtension(path) is ".ttf" or ".otf" or ".woff" or ".woff2")
            .Select(Path.GetFileName).OrderBy(name => name, StringComparer.Ordinal).ToArray()!;
        if (!actualFonts.SequenceEqual(fontFiles.Select(file => file.Name), StringComparer.Ordinal))
            throw new HtmlScriptRuntimeException("The published renderer font-package files do not match its manifest.");

        HtmlPublicFontFileIdentity[] licenses = declaredLicenses.OrderBy(item => item.Key, StringComparer.Ordinal)
            .Select(item => VerifyFile(directory, item.Key, item.Value)).ToArray();
        string[] actualLicenses = Directory.EnumerateFiles(directory, "*.txt").Select(Path.GetFileName)
            .OrderBy(name => name, StringComparer.Ordinal).ToArray()!;
        if (!actualLicenses.SequenceEqual(licenses.Select(file => file.Name), StringComparer.Ordinal))
            throw new HtmlScriptRuntimeException("The published renderer font-package license files do not match its manifest.");
        string[] expectedFiles = new[] { ManifestName }.Concat(fontFiles.Select(file => file.Name))
            .Concat(licenses.Select(file => file.Name)).OrderBy(name => name, StringComparer.Ordinal).ToArray();
        string[] actualFiles = Directory.EnumerateFiles(directory).Select(Path.GetFileName)
            .OrderBy(name => name, StringComparer.Ordinal).ToArray()!;
        if (Directory.EnumerateDirectories(directory).Any()
            || !actualFiles.SequenceEqual(expectedFiles, StringComparer.Ordinal))
            throw new HtmlScriptRuntimeException("The published renderer font-package contains undeclared files.");
        return new HtmlPublicFontPackageIdentity {
            Id = id,
            ManifestSha256 = Digest(manifestBytes),
            FilesSha256 = HtmlPublicArtifactDigest.DirectorySha256(directory),
            ManifestJson = manifestJson,
            Fonts = fontFiles,
            Licenses = licenses
        };
    }

    internal static bool Matches(HtmlPublicFontPackageIdentity expected, HtmlPublicFontPackageIdentity? actual) =>
        actual != null && expected.Id == actual.Id && expected.ManifestSha256 == actual.ManifestSha256
        && expected.FilesSha256 == actual.FilesSha256 && expected.ManifestJson == actual.ManifestJson
        && FilesMatch(expected.Fonts, actual.Fonts) && FilesMatch(expected.Licenses, actual.Licenses);

    private static bool FilesMatch(IReadOnlyList<HtmlPublicFontFileIdentity> expected,
        IReadOnlyList<HtmlPublicFontFileIdentity>? actual) => actual != null && expected.Count == actual.Count
        && expected.Zip(actual).All(pair => pair.First.Name == pair.Second.Name
            && pair.First.ByteCount == pair.Second.ByteCount && pair.First.Sha256 == pair.Second.Sha256);

    private static HtmlPublicFontFileIdentity VerifyFile(string directory, string name, string? expectedSha256) {
        string path = Path.Combine(directory, name);
        if (!File.Exists(path)) throw new HtmlScriptRuntimeException("A published renderer font-package file is missing: " + name);
        byte[] bytes = File.ReadAllBytes(path);
        string sha256 = Digest(bytes);
        if (expectedSha256 != null && !sha256.Equals(expectedSha256, StringComparison.OrdinalIgnoreCase))
            throw new HtmlScriptRuntimeException("A published renderer font-package file does not match its manifest: " + name);
        return new HtmlPublicFontFileIdentity { Name = name, ByteCount = bytes.LongLength, Sha256 = sha256 };
    }

    private static string RequiredText(JsonElement value, string property, int maximumLength) {
        if (!value.TryGetProperty(property, out JsonElement element) || element.ValueKind != JsonValueKind.String
            || string.IsNullOrWhiteSpace(element.GetString()) || element.GetString()!.Length > maximumLength)
            throw new HtmlScriptRuntimeException("The font-package manifest property '" + property + "' is invalid.");
        return element.GetString()!;
    }

    private static string RequiredFileName(JsonElement value, string property) {
        string name = RequiredText(value, property, 128);
        if (name != Path.GetFileName(name) || name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            throw new HtmlScriptRuntimeException("The font-package manifest contains an invalid file name.");
        return name;
    }

    private static string RequiredSha256(JsonElement value, string property) {
        string digest = RequiredText(value, property, 64).ToLowerInvariant();
        if (digest.Length != 64 || digest.Any(character => !Uri.IsHexDigit(character)))
            throw new HtmlScriptRuntimeException("The font-package manifest contains an invalid SHA-256 digest.");
        return digest;
    }

    private static string Digest(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
}
