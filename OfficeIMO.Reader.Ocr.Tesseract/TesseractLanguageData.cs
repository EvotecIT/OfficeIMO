using System.Collections.Concurrent;
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
    private static readonly ConcurrentDictionary<string, SemaphoreSlim> PackageLocks = new(StringComparer.Ordinal);
    private static readonly IReadOnlyDictionary<string, Package> Catalog = new Dictionary<string, Package>(StringComparer.Ordinal) {
        ["ara"] = new Package("ara", "e3206d3dc87fd50c24a0fb9f01838615911d25168f4e64415244b67d2bb3e729", 1_432_056L),
        ["ces"] = new Package("ces", "934bcaf97ef3348413263331131c9fa7f55f30db333c711929c124fb635f7e1b", 3_795_684L),
        ["chi_sim"] = new Package("chi_sim", "a5fcb6f0db1e1d6d8522f39db4e848f05984669172e584e8d76b6b3141e1f730", 2_469_156L),
        ["chi_tra"] = new Package("chi_tra", "529c5b5797d64b126065cd55f2bb4c7fd7b15790798091b1ff259941a829330b", 2_366_642L),
        ["dan"] = new Package("dan", "acb1fd074487a31d1294fcdfd7d7c673467ffd8aeacb2ccd61ebcbf04eb4e2fa", 2_580_059L),
        ["deu"] = new Package("deu", "19d219bbb6672c869d20a9636c6816a81eb9a71796cb93ebe0cb1530e2cdb22d", 1_525_436L),
        ["ell"] = new Package("ell", "4fba8a0b461038d51f1c20d043d4f2ac38c4e778f1b90830847f7bd8fa3ba726", 1_419_514L),
        ["eng"] = new Package("eng", "7d4322bd2a7749724879683fc3912cb542f19906c83bcc1a52132556427170b2", 4_113_088L),
        ["fin"] = new Package("fin", "61a04cd62b507c3d9ae0e1cda399e6715ebf49dea9df47897c8acdcd3bd3e13c", 7_865_732L),
        ["fra"] = new Package("fra", "ced037562e8c80c13122dece28dd477d399af80911a28791a66a63ac1e3445ca", 1_130_365L),
        ["heb"] = new Package("heb", "11f9e43ab227f786352a50f75c94c2e9906f1baba86d93276da19da7ce0904db", 961_404L),
        ["hin"] = new Package("hin", "4c73ffc59d497c186b19d1e90f5d721d678ea6b2e277b719bee4e2af12271825", 1_122_751L),
        ["hun"] = new Package("hun", "35067e7cfe102dcdc953f9a758fdfaa6296b17a1ee6d874ee780fa306430b9fb", 5_296_273L),
        ["ita"] = new Package("ita", "b8f89e1e785118dac4d51ae042c029a64edb5c3ee42ef73027a6d412748d8827", 2_701_314L),
        ["jpn"] = new Package("jpn", "1f5de9236d2e85f5fdf4b3c500f2d4926f8d9449f28f5394472d9e8d83b91b4d", 2_471_260L),
        ["kor"] = new Package("kor", "6b85e11d9bbf07863b97b3523b1b112844c43e713df8b66418a081fd1060b3b2", 1_677_415L),
        ["nld"] = new Package("nld", "ced0e5e046a84c908a6aa7accbef9a232c4a5d9a8276691b81c6ee64d02963f6", 6_050_296L),
        ["nor"] = new Package("nor", "0451eb4f8049ae78196806bf878a389a2f40f1386fe038568cf4441226ba6ef2", 3_610_079L),
        ["pol"] = new Package("pol", "c4476cdbc0e33d898d32345122b7be1cbf85ace15f920f06c7714756e1ef79b2", 4_765_518L),
        ["por"] = new Package("por", "c4932b937207a9514b7514d518b931a99938c02a28a5a5a553f8599ed58b7deb", 1_982_756L),
        ["ron"] = new Package("ron", "9adfde6b51ba4b97efd10ea37c3070fd3fc2bad7815e81f5c3c198cd96216cc9", 2_376_323L),
        ["rus"] = new Package("rus", "e16e5e036cce1d9ec2b00063cf8b54472625b9e14d893a169e2b0dedeb4df225", 3_861_738L),
        ["slk"] = new Package("slk", "fbcc400a9c74c6a13d922fcb1211b655d1b165387b675ed75cd2dbd756b974a5", 4_427_661L),
        ["spa"] = new Package("spa", "6f2e04d02774a18f01bed44b1111f2cd7f3ba7ac9dc4373cd3f898a40ea6b464", 2_294_433L),
        ["swe"] = new Package("swe", "f7304988d41f833efebcc2d529df54b1903ecebbc3da1faabd19a0fddd4fe586", 4_167_034L),
        ["tur"] = new Package("tur", "7393381111e1152420fc4092cb44eef4237580d21b92bf30d7d221aad192c6b7", 4_550_554L),
        ["ukr"] = new Package("ukr", "d59e53e2bded32f4445f124b4b00240fcac7e8044c003ab822ccb94f0b3db59b", 3_825_102L),
        ["vie"] = new Package("vie", "79df64caf7bcfb2a27df5042ecb6121e196eada34da774956995747636d5bfa1", 531_275L),
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
        SemaphoreSlim packageLock = PackageLocks.GetOrAdd(destination, static _ => new SemaphoreSlim(1, 1));
        await packageLock.WaitAsync(cancellationToken).ConfigureAwait(false);
        try {
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

                PublishVerifiedFile(staging, destination, package);
                if (!TryVerify(destination, package, out string? installedHash)) {
                    throw new InvalidDataException("Installed Tesseract language data failed post-write verification.");
                }
                return new TesseractLanguageDataFile(package.Language, destination, installedHash!, package.ByteCount, downloaded: true);
            } finally {
                if (File.Exists(staging)) File.Delete(staging);
            }
        } finally {
            packageLock.Release();
        }
    }

    private static void PublishVerifiedFile(string staging, string destination, Package package) {
        for (int attempt = 0; attempt < 3; attempt++) {
            if (!File.Exists(destination)) {
                try {
                    File.Move(staging, destination);
                    return;
                } catch (IOException) when (TryVerify(destination, package, out _)) {
                    return;
                } catch (IOException) when (attempt < 2) {
                    continue;
                }
            }

            RejectReparsePoint(destination);
            try {
                File.Replace(staging, destination, destinationBackupFileName: null);
                return;
            } catch (FileNotFoundException) when (attempt < 2) {
                // Another process moved the destination between the existence check and replacement.
            } catch (IOException) when (TryVerify(destination, package, out _)) {
                return;
            }
        }

        throw new IOException("Could not atomically publish verified Tesseract language data.");
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
