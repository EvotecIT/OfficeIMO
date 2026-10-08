#if OFFICEIMO_BENCHMARK_NEW_APIS
using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using ExcelReader.Core.Reader;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks;

/// <summary>Separates the pinned small ciphertext from the generated upstream large input.</summary>
public enum EncryptedWorkbookInput { OriginalSmallXlsx, GeneratedLargeXlsx }

internal sealed record EncryptedWorkbookFixture(byte[] Plain, byte[] Encrypted, string PlainPath,
    string EncryptedPath, int Rows, int Columns, bool FileStreams) {
    internal const string Password = "hunter2";
    private const string UpstreamCommit = "ca5b50f99e8ef57ab476f0a2bc8043558d58b28d";
    private const string PlainHash = "85FBDE3D5CC6C936BE8D9C8A5F9658E4CDED25145A2E6449C5D1FCD54F64E602";
    private const string EncryptedHash = "13F9532B916634CB112D3593E7AB859F67B2DA35A3A47E574F1C1E26C5DA370D";
    private const string Crypto = "Agile AES-256-CBC/SHA-512/spin100000/dataIntegrity";

    internal static int LargeRows {
        get {
            string? value = Environment.GetEnvironmentVariable("EXCELREADER_LARGE_ENCRYPTED_ROWS");
            if (string.IsNullOrWhiteSpace(value)) return 300_000;
            return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int rows)
                && rows is > 0 and <= 1_048_576 ? rows
                : throw new ArgumentException("EXCELREADER_LARGE_ENCRYPTED_ROWS must be 1..1048576.");
        }
    }

    internal static EncryptedWorkbookFixture Load(EncryptedWorkbookInput input) {
        if (input == EncryptedWorkbookInput.OriginalSmallXlsx) {
            string root = ExistingDirectory("OFFICEIMO_ENCRYPTED_BENCHMARK_CORPUS");
            string plain = Path.Combine(root, "agile-aes256-sha512.plain.xlsx");
            string encrypted = Path.Combine(root, "agile-aes256-sha512.xlsx");
            return Describe(new(ReadChecked(plain, PlainHash, 8915), ReadChecked(encrypted, EncryptedHash, 15360),
                plain, encrypted, 2, 2, FileStreams: true), input);
        }
        string cache = ExistingDirectory("OFFICEIMO_ENCRYPTED_BENCHMARK_CACHE");
        string stem = $"upstream-encrypted-xlsx-{LargeRows}";
        string manifestPath = Path.Combine(cache, stem + ".json");
        if (!File.Exists(manifestPath)) throw new FileNotFoundException(
            "Prepare the immutable generated input with --prepare-encrypted before measurements.", manifestPath);
        CacheManifest manifest = JsonSerializer.Deserialize<CacheManifest>(File.ReadAllText(manifestPath))
            ?? throw new InvalidDataException("Invalid encrypted fixture manifest.");
        if (manifest.Upstream != UpstreamCommit || manifest.Rows != LargeRows || manifest.Crypto != Crypto
            || manifest.Producer != "ExcelReader.NET 6.0.0 / WorkbookGenerator.BuildAsync / Excel.EncryptPackage")
            throw new InvalidDataException("Generated input provenance differs from the pinned producer.");
        string plainPath = Path.Combine(cache, stem + ".plain.xlsx");
        string encryptedPath = Path.Combine(cache, stem + ".xlsx");
        return Describe(new(ReadChecked(plainPath, manifest.PlainSha256, manifest.PlainBytes),
            ReadChecked(encryptedPath, manifest.EncryptedSha256, manifest.EncryptedBytes), plainPath,
            encryptedPath, LargeRows, 4, FileStreams: false), input);
    }

    internal static async Task PrepareLargeAsync() {
        string cache = ExistingDirectory("OFFICEIMO_ENCRYPTED_BENCHMARK_CACHE");
        string stem = $"upstream-encrypted-xlsx-{LargeRows}";
        string manifestPath = Path.Combine(cache, stem + ".json");
        if (File.Exists(manifestPath)) { Load(EncryptedWorkbookInput.GeneratedLargeXlsx); return; }
        string plainPath = Path.Combine(cache, stem + ".plain.xlsx");
        string encryptedPath = Path.Combine(cache, stem + ".xlsx");
        if (File.Exists(plainPath) || File.Exists(encryptedPath))
            throw new IOException("An incomplete generated cache exists; inspect it before preparing a new input.");
        byte[] plain = await RawWorkbookFixture.CreateAsync(LargeRows, RawWorkbookFormat.Xlsx);
        using var plainStream = new MemoryStream(plain, writable: false);
        using var encryptedStream = new MemoryStream();
        ExcelReaderApi.EncryptPackage(plainStream, encryptedStream, new ExcelPassword(Password));
        byte[] encrypted = encryptedStream.ToArray();
        var fixture = new EncryptedWorkbookFixture(plain, encrypted, plainPath, encryptedPath, LargeRows, 4, false);
        // Authenticate and compare every field before publishing a usable cache manifest.
        await new EncryptedReadWorkload().SetupAsync(fixture);
        WriteNew(plainPath, plain);
        WriteNew(encryptedPath, encrypted);
        var manifest = new CacheManifest(UpstreamCommit,
            "ExcelReader.NET 6.0.0 / WorkbookGenerator.BuildAsync / Excel.EncryptPackage", Crypto,
            LargeRows, plain.Length, encrypted.Length, Hash(plain), Hash(encrypted));
        WriteNew(manifestPath, System.Text.Encoding.UTF8.GetBytes(JsonSerializer.Serialize(manifest,
            new JsonSerializerOptions { WriteIndented = true })));
        Console.WriteLine($"Prepared immutable encrypted input: {manifestPath}");
    }

    internal Stream OpenPlainStream() => FileStreams
        ? new FileStream(PlainPath, FileMode.Open, FileAccess.Read) : new MemoryStream(Plain, writable: false);
    internal Stream OpenEncryptedStream() => FileStreams
        ? new FileStream(EncryptedPath, FileMode.Open, FileAccess.Read) : new MemoryStream(Encrypted, writable: false);
    internal static ExcelReaderOptions PeerOptions(bool verify = true) => new() {
        Password = new ExcelPassword(Password), VerifyEncryptedIntegrity = verify,
    };
    internal static ExcelReadOptions OfficeOptions() => new() { HasHeaderRow = false, SheetIndex = 0 };
    private static string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes));
    private static byte[] ReadChecked(string path, string hash, int bytes) {
        byte[] value = File.ReadAllBytes(path);
        if (value.Length != bytes || Hash(value) != hash) throw new InvalidDataException($"Input identity differs: {path}");
        return value;
    }
    private static string ExistingDirectory(string variable) {
        string? value = Environment.GetEnvironmentVariable(variable);
        if (string.IsNullOrWhiteSpace(value) || !Path.IsPathFullyQualified(value) || !Directory.Exists(value))
            throw new ArgumentException($"{variable} must name an existing absolute task-owned directory.");
        return value;
    }
    private static void WriteNew(string path, byte[] bytes) {
        using var stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        stream.Write(bytes);
    }
    private static EncryptedWorkbookFixture Describe(EncryptedWorkbookFixture fixture, EncryptedWorkbookInput input) {
        Console.WriteLine($"Encrypted input={input}; rows={fixture.Rows}; columns={fixture.Columns}; "
            + $"plainBytes={fixture.Plain.Length}; encryptedBytes={fixture.Encrypted.Length}; "
            + $"plainSHA256={Hash(fixture.Plain)}; encryptedSHA256={Hash(fixture.Encrypted)}; "
            + $"stream={(fixture.FileStreams ? "FileStream" : "MemoryStream")}; producerCommit={UpstreamCommit}; "
            + $"password=hunter2; generatedCrypto={Crypto}. Small encryption producer is unpinned; "
            + "its authenticated plaintext oracle comes from independent msoffcrypto-tool decryption.");
        return fixture;
    }
    private sealed record CacheManifest(string Upstream, string Producer, string Crypto, int Rows,
        int PlainBytes, int EncryptedBytes, string PlainSha256, string EncryptedSha256);
}
#endif
