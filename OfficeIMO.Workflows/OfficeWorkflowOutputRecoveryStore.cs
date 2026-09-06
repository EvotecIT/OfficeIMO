using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Core.Internal;
using OfficeIMO.Internal;

namespace OfficeIMO.Workflows;

/// <summary>Stores bounded, private recovery artifacts before non-atomic provider publication.</summary>
public sealed class OfficeWorkflowOutputRecoveryStore {
    private const string Prefix = "output-";
    private const int MaximumRecords = 100;
    private const int MaximumMetadataBytes = 65536;
    private static readonly HashSet<string> Extensions = new(StringComparer.OrdinalIgnoreCase) { ".pdf", ".html", ".docx", ".xlsx", ".pptx" };

    /// <summary>Creates a store under a host-selected local directory. Admission is bounded across processes.</summary>
    public OfficeWorkflowOutputRecoveryStore(string directory, long maximumRetainedBytes = 1024L * 1024 * 1024) {
        DirectoryPath = OfficeStorageIdentity.GetLocalPath(directory) ?? throw new ArgumentException("Recovery requires a local directory.", nameof(directory));
        if (maximumRetainedBytes < 1) throw new ArgumentOutOfRangeException(nameof(maximumRetainedBytes));
        MaximumRetainedBytes = maximumRetainedBytes;
    }

    /// <summary>Gets the local storage root.</summary>
    public string DirectoryPath { get; }
    /// <summary>Gets the aggregate artifact admission limit.</summary>
    public long MaximumRetainedBytes { get; }

    /// <summary>Lists available records, excluding active publications and malformed records. Contents are verified on demand.</summary>
    public IReadOnlyList<OfficeWorkflowOutputRecovery> GetRecoveries() {
        if (!Directory.Exists(DirectoryPath)) return Array.Empty<OfficeWorkflowOutputRecovery>();
        var records = new List<OfficeWorkflowOutputRecovery>();
        foreach (string directory in Directory.EnumerateDirectories(DirectoryPath, Prefix + "*", SearchOption.TopDirectoryOnly).Where(IsRecordDirectory).Take(MaximumRecords + 1)) {
            try {
                EnsureRegularDirectory(directory);
                if (!File.Exists(Path.Combine(directory, "record.json"))) continue;
                using FileStream lease = OpenLease(Path.Combine(directory, ".lease"));
                if (ReadRecord(directory) is { } record) records.Add(record);
            } catch (Exception error) when (error is IOException or UnauthorizedAccessException or JsonException or ArgumentException) { }
        }
        return records.OrderByDescending(record => record.CreatedUtc).ToArray();
    }

    /// <summary>Verifies the current local copy against its stored length and fingerprint before a host opens it.</summary>
    public async Task VerifyAsync(OfficeWorkflowOutputRecovery recovery, CancellationToken token = default) {
        string directory = ValidateRecordDirectory(recovery);
        using FileStream lease = OpenLease(Path.Combine(directory, ".lease"));
        OfficeWorkflowOutputRecovery current = ReadRecord(directory) ?? throw new IOException("The recovery record is no longer available.");
        if (current.Sha256 != recovery.Sha256 || current.Length != recovery.Length || current.FilePath != recovery.FilePath) throw new IOException("The recovery record changed.");
        string fingerprint = await OfficeStreamPublication.ReadFingerprintAsync(_ => Task.FromResult<Stream>(OpenArtifact(current.FilePath)),
            MaximumRetainedBytes, token).ConfigureAwait(false);
        if (!string.Equals(fingerprint, current.Sha256, StringComparison.OrdinalIgnoreCase)) throw new IOException("The recovery copy no longer matches the prepared artifact.");
    }

    /// <summary>Deletes a recovery copy after the user explicitly discards it. The provider destination is never modified.</summary>
    public void Discard(OfficeWorkflowOutputRecovery recovery) {
        string directory = ValidateRecordDirectory(recovery);
        using (FileStream lease = OpenLease(Path.Combine(directory, ".lease"))) {
            OfficeWorkflowOutputRecovery current = ReadRecord(directory) ?? throw new IOException("The recovery record is no longer available.");
            if (current.Sha256 != recovery.Sha256) throw new IOException("The recovery record changed.");
            File.Delete(current.FilePath);
            File.Delete(Path.Combine(directory, "record.json"));
        }
        File.Delete(Path.Combine(directory, ".lease"));
        Directory.Delete(directory, recursive: false);
    }

    internal async Task<RecoveryLease> CreateAsync(byte[] bytes, string destination, string name, CancellationToken token) {
        string extension = ValidateExtension(name);
        if (destination.Length > 4096 || name.Length > 4096) throw new ArgumentException("The provider reference exceeds the supported size.");
        token.ThrowIfCancellationRequested();
        Directory.CreateDirectory(DirectoryPath);
        using FileStream admission = await AcquireAdmissionAsync(token).ConfigureAwait(false);
        string[] directories = Directory.EnumerateDirectories(DirectoryPath, Prefix + "*", SearchOption.TopDirectoryOnly).Where(IsRecordDirectory).Take(MaximumRecords + 1).ToArray();
        if (directories.Length >= MaximumRecords) throw new IOException("The workflow recovery store is full. Recover or discard existing copies before publishing more outputs.");
        long retained = 0;
        foreach (string existing in directories) {
            EnsureRegularDirectory(existing);
            foreach (string file in Directory.EnumerateFiles(existing, "*", SearchOption.TopDirectoryOnly).Take(16)) {
                EnsureRegularFile(file);
                retained = checked(retained + new FileInfo(file).Length);
            }
        }
        if (bytes.LongLength > MaximumRetainedBytes - retained) throw new IOException("The workflow recovery byte limit has been reached. Recover or discard existing copies first.");
        string directory = OfficeTemporaryDirectory.Create(Prefix, DirectoryPath);
        string path = Path.Combine(directory, "output" + extension);
        FileStream? lease = null;
        try {
            lease = OpenLease(Path.Combine(directory, ".lease"));
            string hash = Convert.ToHexString(SHA256.HashData(bytes));
            string id = Path.GetFileName(directory)[Prefix.Length..];
            var record = new Metadata(1, id, name, destination, bytes.LongLength, hash, DateTimeOffset.UtcNow);
            byte[] metadata = JsonSerializer.SerializeToUtf8Bytes(record);
            if (metadata.Length > MaximumMetadataBytes) throw new InvalidDataException("The recovery metadata exceeds its size limit.");
            OfficeFileCommit.WriteAllBytes(path, bytes, OfficeFileCommit.UnixFileAccessPolicy.OwnerOnly);
            OfficeFileCommit.WriteAllBytes(Path.Combine(directory, "record.json"), metadata, OfficeFileCommit.UnixFileAccessPolicy.OwnerOnly);
            await OfficeStreamPublication.VerifyFingerprintAsync(_ => Task.FromResult<Stream>(OpenArtifact(path)),
                hash, MaximumRetainedBytes, token).ConfigureAwait(false);
            token.ThrowIfCancellationRequested();
            var recovery = new OfficeWorkflowOutputRecovery(id, name, destination, path, bytes.LongLength, hash, record.CreatedUtc);
            var result = new RecoveryLease(recovery, lease);
            lease = null;
            return result;
        } catch (Exception failure) {
            lease?.Dispose();
            try {
                File.Delete(path); File.Delete(Path.Combine(directory, "record.json")); File.Delete(Path.Combine(directory, ".lease"));
                Directory.Delete(directory, recursive: false);
            } catch (Exception error) when (error is IOException or UnauthorizedAccessException) {
                failure.Data["OfficeIMO.Storage.OutputRecoveryCleanupFailed"] = directory;
            }
            throw;
        }
    }

    private async Task<FileStream> AcquireAdmissionAsync(CancellationToken token) {
        DateTime deadline = DateTime.UtcNow.AddSeconds(10);
        while (true) {
            token.ThrowIfCancellationRequested();
            try { return OpenLease(Path.Combine(DirectoryPath, ".admission.lock")); }
            catch (IOException) when (DateTime.UtcNow < deadline) { await Task.Delay(25, token).ConfigureAwait(false); }
        }
    }

    private OfficeWorkflowOutputRecovery? ReadRecord(string directory) {
        string id = Path.GetFileName(directory)[Prefix.Length..];
        if (!Guid.TryParseExact(id, "N", out _)) return null;
        string metadataPath = Path.Combine(directory, "record.json");
        if (!File.Exists(metadataPath)) return null;
        using FileStream input = OpenArtifact(metadataPath);
        Metadata? record = JsonSerializer.Deserialize<Metadata>(OfficeStreamReader.ReadAllBytes(input, MaximumMetadataBytes));
        if (record is null || record.Version != 1 || record.Id != id || record.Name is null || record.Name.Length > 4096 ||
            record.Destination is null || record.Destination.Length > 4096 || record.Length < 0 || record.Length > MaximumRetainedBytes ||
            record.Sha256 is null || record.Sha256.Length != 64 || record.Sha256.Any(character => !Uri.IsHexDigit(character))) return null;
        string extension = ValidateExtension(record.Name);
        string path = Path.Combine(directory, "output" + extension);
        return new(id, record.Name, record.Destination, path, record.Length, record.Sha256, record.CreatedUtc);
    }

    private static bool IsRecordDirectory(string directory) {
        string name = Path.GetFileName(directory);
        return name.StartsWith(Prefix, StringComparison.Ordinal) && Guid.TryParseExact(name[Prefix.Length..], "N", out _);
    }

    private string ValidateRecordDirectory(OfficeWorkflowOutputRecovery record) {
        ArgumentNullException.ThrowIfNull(record);
        if (!Guid.TryParseExact(record.Id, "N", out _)) throw new ArgumentException("The recovery identifier is invalid.", nameof(record));
        string directory = Path.Combine(DirectoryPath, Prefix + record.Id);
        EnsureRegularDirectory(directory);
        if (!string.Equals(Path.GetFullPath(record.FilePath), Path.Combine(directory, "output" + ValidateExtension(record.Name)), StringComparison.Ordinal)) {
            throw new ArgumentException("The recovery record belongs to another store.", nameof(record));
        }
        return directory;
    }

    internal static string ValidateExtension(string name) {
        string extension = Path.GetExtension(name).ToLowerInvariant();
        if (!Extensions.Contains(extension)) throw new ArgumentException("Choose a supported document output filename.", nameof(name));
        return extension;
    }
    private static void EnsureRegularDirectory(string path) {
        FileAttributes attributes = File.GetAttributes(path);
        if ((attributes & FileAttributes.ReparsePoint) != 0 || (attributes & FileAttributes.Directory) == 0) throw new IOException("Recovery directories cannot be links.");
    }
    private static void EnsureRegularFile(string path) {
        if ((File.GetAttributes(path) & (FileAttributes.ReparsePoint | FileAttributes.Directory)) != 0) throw new IOException("Recovery files cannot be links or directories.");
    }
    private static FileStream OpenArtifact(string path) {
        EnsureRegularFile(path);
        return OfficePathIdentity.OpenRegularFileForRead(path, OfficePathIdentity.ResolvePhysicalPath(Path.GetDirectoryName(path)!), 81920);
    }
    private static FileStream OpenLease(string path) {
        if (File.Exists(path)) EnsureRegularFile(path);
        var options = new FileStreamOptions { Mode = FileMode.OpenOrCreate, Access = FileAccess.ReadWrite, Share = FileShare.None, BufferSize = 1 };
        if (!OperatingSystem.IsWindows()) options.UnixCreateMode = UnixFileMode.UserRead | UnixFileMode.UserWrite;
        return new FileStream(path, options);
    }

    private sealed record Metadata(int Version, string Id, string Name, string Destination, long Length, string Sha256, DateTimeOffset CreatedUtc);

    internal sealed class RecoveryLease(OfficeWorkflowOutputRecovery recovery, FileStream lease) : IDisposable {
        internal OfficeWorkflowOutputRecovery Recovery { get; } = recovery;
        public void Dispose() => lease.Dispose();
    }
}
