using System.Net.Http;
using System.Security.Cryptography;

namespace OfficeIMO.Reader.Ocr.Tesseract;

/// <summary>Controls opt-in provisioning of immutable, checksum-verified Tesseract language data.</summary>
public sealed class TesseractLanguageDataOptions {
    /// <summary>Optional cache root. A versioned per-user OfficeIMO cache is used by default.</summary>
    public string? CacheDirectory { get; set; }

    /// <summary>Optional caller-owned HTTP client, useful for proxy configuration and deterministic tests.</summary>
    public HttpClient? HttpClient { get; set; }

    /// <summary>Maximum downloaded bytes for one trained-data file. Defaults to 32 MiB.</summary>
    public long MaxBytesPerLanguage { get; set; } = 32L * 1024L * 1024L;
}

/// <summary>One verified local trained-data file.</summary>
public sealed class TesseractLanguageDataFile {
    internal TesseractLanguageDataFile(string language, string path, string sha256, long byteCount, bool downloaded) {
        Language = language;
        Path = path;
        Sha256 = sha256;
        ByteCount = byteCount;
        Downloaded = downloaded;
    }

    /// <summary>Tesseract language identifier.</summary>
    public string Language { get; }
    /// <summary>Absolute local file path.</summary>
    public string Path { get; }
    /// <summary>Verified lowercase SHA-256 digest.</summary>
    public string Sha256 { get; }
    /// <summary>Verified file length.</summary>
    public long ByteCount { get; }
    /// <summary>True when this call downloaded the file; false when a valid cached file was reused.</summary>
    public bool Downloaded { get; }
}

/// <summary>Verified language-data directory ready for <see cref="TesseractOcrEngineOptions.TessdataDirectory"/>.</summary>
public sealed class TesseractLanguageDataResult {
    internal TesseractLanguageDataResult(string directory, IReadOnlyList<TesseractLanguageDataFile> files) {
        Directory = directory;
        Files = files;
    }

    /// <summary>Absolute directory containing the verified trained-data files.</summary>
    public string Directory { get; }
    /// <summary>Requested files in language-expression order.</summary>
    public IReadOnlyList<TesseractLanguageDataFile> Files { get; }
    /// <summary>True when at least one file was downloaded.</summary>
    public bool Downloaded => Files.Any(static file => file.Downloaded);
}

/// <summary>Curated, immutable Tesseract fast-language provisioning for the default OfficeIMO languages.</summary>
public static class TesseractLanguageData {
    private const string CatalogCommit = "87416418657359cb625c412a48b6e1d6d41c29bd";
    private static readonly HttpClient SharedClient = new HttpClient { Timeout = TimeSpan.FromMinutes(3) };
    private static readonly IReadOnlyDictionary<string, Package> Catalog = new Dictionary<string, Package>(StringComparer.Ordinal) {
        ["eng"] = new Package("eng", "7d4322bd2a7749724879683fc3912cb542f19906c83bcc1a52132556427170b2", 4_113_088L),
        ["pol"] = new Package("pol", "c4476cdbc0e33d898d32345122b7be1cbf85ace15f920f06c7714756e1ef79b2", 4_765_518L),
        ["osd"] = new Package("osd", "9cf5d576fcc47564f11265841e5ca839001e7e6f38ff7f7aacf46d15a96b00ff", 10_562_727L)
    };

    /// <summary>Immutable tessdata_fast commit used by the built-in catalog.</summary>
    public static string Version => CatalogCommit;

    /// <summary>Language identifiers available through the checksum-pinned built-in catalog.</summary>
    public static IReadOnlyList<string> SupportedLanguages => Catalog.Keys.OrderBy(static language => language, StringComparer.Ordinal).ToArray();

    /// <summary>
    /// Ensures every language in a Tesseract expression such as <c>eng+pol</c> exists in a private versioned cache.
    /// Network access occurs only for missing or corrupt files and every payload must match the built-in SHA-256 digest.
    /// </summary>
    public static Task<TesseractLanguageDataResult> EnsureAsync(
        string languageExpression = "eng",
        TesseractLanguageDataOptions? options = null,
        CancellationToken cancellationToken = default) {
        string[] languages = ParseLanguageExpression(languageExpression);
        return EnsurePackagesAsync(languages.Select(language => ResolvePackage(language)).ToArray(), options, cancellationToken);
    }

    internal static async Task<TesseractLanguageDataResult> EnsurePackagesAsync(
        IReadOnlyList<Package> packages,
        TesseractLanguageDataOptions? options,
        CancellationToken cancellationToken) {
        if (packages == null) throw new ArgumentNullException(nameof(packages));
        TesseractLanguageDataOptions effective = options ?? new TesseractLanguageDataOptions();
        if (effective.MaxBytesPerLanguage < 1L) throw new ArgumentOutOfRangeException(nameof(options), "Maximum language-data bytes must be positive.");
        string directory = ResolveCacheDirectory(effective.CacheDirectory);
        Directory.CreateDirectory(directory);
        HttpClient client = effective.HttpClient ?? SharedClient;
        var files = new List<TesseractLanguageDataFile>(packages.Count);
        for (int i = 0; i < packages.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            Package package = packages[i];
            if (package.ByteCount > effective.MaxBytesPerLanguage) {
                throw new InvalidOperationException("Pinned Tesseract language data exceeds the configured per-language download limit.");
            }
            files.Add(await EnsurePackageAsync(directory, package, client, effective.MaxBytesPerLanguage, cancellationToken).ConfigureAwait(false));
        }
        return new TesseractLanguageDataResult(directory, files.AsReadOnly());
    }

    private static async Task<TesseractLanguageDataFile> EnsurePackageAsync(
        string directory,
        Package package,
        HttpClient client,
        long maximumBytes,
        CancellationToken cancellationToken) {
        string destination = Path.Combine(directory, package.Language + ".traineddata");
        if (TryVerify(destination, package, out string? cachedHash)) {
            return new TesseractLanguageDataFile(package.Language, destination, cachedHash!, package.ByteCount, downloaded: false);
        }
        RejectReparsePoint(destination);
        string staging = destination + ".download-" + Guid.NewGuid().ToString("N");
        try {
            using (var request = new HttpRequestMessage(HttpMethod.Get, package.Uri))
            using (HttpResponseMessage response = await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false)) {
                response.EnsureSuccessStatusCode();
                if (response.Content.Headers.ContentLength.HasValue && response.Content.Headers.ContentLength.Value > maximumBytes) {
                    throw new InvalidDataException("Tesseract language-data response exceeded the configured byte limit.");
                }
                using Stream source = await response.Content.ReadAsStreamAsync().ConfigureAwait(false);
                using var destinationStream = new FileStream(staging, FileMode.CreateNew, FileAccess.Write, FileShare.None, 81920, useAsync: true);
                using HashAlgorithm sha256 = SHA256.Create();
                byte[] buffer = new byte[81920];
                long total = 0L;
                while (true) {
                    int read = await source.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false);
                    if (read == 0) break;
                    total = checked(total + read);
                    if (total > maximumBytes) throw new InvalidDataException("Tesseract language-data response exceeded the configured byte limit.");
                    sha256.TransformBlock(buffer, 0, read, null, 0);
                    await destinationStream.WriteAsync(buffer, 0, read, cancellationToken).ConfigureAwait(false);
                }
                sha256.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
                string actualHash = ToHex(sha256.Hash!);
                if (total != package.ByteCount || !string.Equals(actualHash, package.Sha256, StringComparison.Ordinal)) {
                    throw new InvalidDataException("Tesseract language-data payload did not match the pinned length and SHA-256 digest.");
                }
            }

            if (File.Exists(destination)) File.Delete(destination);
            try {
                File.Move(staging, destination);
            } catch (IOException) when (TryVerify(destination, package, out _)) {
                File.Delete(staging);
            }
            if (!TryVerify(destination, package, out string? installedHash)) {
                throw new InvalidDataException("Installed Tesseract language data failed post-write verification.");
            }
            return new TesseractLanguageDataFile(package.Language, destination, installedHash!, package.ByteCount, downloaded: true);
        } finally {
            if (File.Exists(staging)) File.Delete(staging);
        }
    }

    private static bool TryVerify(string path, Package package, out string? actualHash) {
        actualHash = null;
        if (!File.Exists(path)) return false;
        RejectReparsePoint(path);
        var info = new FileInfo(path);
        if (info.Length != package.ByteCount) return false;
        using FileStream stream = File.OpenRead(path);
        using HashAlgorithm sha256 = SHA256.Create();
        actualHash = ToHex(sha256.ComputeHash(stream));
        return string.Equals(actualHash, package.Sha256, StringComparison.Ordinal);
    }

    private static void RejectReparsePoint(string path) {
        if (!File.Exists(path)) return;
        if ((File.GetAttributes(path) & FileAttributes.ReparsePoint) != 0) {
            throw new IOException("Tesseract language-data cache entries cannot be symbolic links or reparse points.");
        }
    }

    private static Package ResolvePackage(string language) {
        if (Catalog.TryGetValue(language, out Package? package)) return package;
        throw new NotSupportedException(
            "Automatic Tesseract language provisioning currently supports " + string.Join(", ", SupportedLanguages) +
            ". Install other languages with the host package manager or set TessdataDirectory explicitly.");
    }

    private static string[] ParseLanguageExpression(string expression) {
        if (string.IsNullOrWhiteSpace(expression)) throw new ArgumentException("Tesseract language expression cannot be empty.", nameof(expression));
        string[] values = expression.Split('+').Select(static value => value.Trim()).Where(static value => value.Length > 0).Distinct(StringComparer.Ordinal).ToArray();
        if (values.Length == 0 || values.Any(static value => value.Any(character => !(character >= 'a' && character <= 'z') && !(character >= 'A' && character <= 'Z') && character != '_' && character != '-'))) {
            throw new ArgumentException("Tesseract language identifiers may contain only letters, underscores, and hyphens.", nameof(expression));
        }
        return values;
    }

    private static string ResolveCacheDirectory(string? configured) {
        string root = string.IsNullOrWhiteSpace(configured)
            ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeIMO", "Ocr", "Tesseract", "tessdata-fast", CatalogCommit)
            : configured!;
        if (string.IsNullOrWhiteSpace(root)) throw new InvalidOperationException("A Tesseract language-data cache directory could not be resolved.");
        return Path.GetFullPath(Environment.ExpandEnvironmentVariables(root));
    }

    private static string ToHex(byte[] bytes) => string.Concat(bytes.Select(static value => value.ToString("x2", System.Globalization.CultureInfo.InvariantCulture)));

    internal sealed class Package {
        internal Package(string language, string sha256, long byteCount, Uri? uri = null) {
            Language = language;
            Sha256 = sha256;
            ByteCount = byteCount;
            Uri = uri ?? new Uri("https://raw.githubusercontent.com/tesseract-ocr/tessdata_fast/" + CatalogCommit + "/" + language + ".traineddata", UriKind.Absolute);
        }
        internal string Language { get; }
        internal string Sha256 { get; }
        internal long ByteCount { get; }
        internal Uri Uri { get; }
    }
}
