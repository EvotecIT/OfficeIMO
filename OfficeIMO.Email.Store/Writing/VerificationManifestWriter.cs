using OfficeIMO.Email;
using OfficeIMO.Drawing.Internal;
using System.Security.Cryptography;

namespace OfficeIMO.Email.Store;

internal sealed class VerificationManifestWriter : IDisposable {
    private readonly string _destinationPath;
    private readonly string _temporaryPath;
    private readonly bool _overwrite;
    private readonly StreamWriter _writer;
    private readonly HashAlgorithm _aggregate = SHA256.Create();
    private readonly HMACSHA256 _differencePathDigest;
    private bool _completed;

    private VerificationManifestWriter(string destinationPath, bool overwrite,
        byte[] differencePathKey) {
        _destinationPath = Path.GetFullPath(destinationPath);
        _overwrite = overwrite;
        _differencePathDigest = new HMACSHA256(differencePathKey);
        string directory = Path.GetDirectoryName(_destinationPath) ?? Path.GetFullPath(".");
        Directory.CreateDirectory(directory);
        _temporaryPath = Path.Combine(directory, string.Concat(".", Path.GetFileName(_destinationPath),
            ".", Guid.NewGuid().ToString("N"), ".tmp"));
        _writer = new StreamWriter(new FileStream(_temporaryPath, FileMode.CreateNew,
            FileAccess.Write, FileShare.Read, 64 * 1024), new UTF8Encoding(false));
        _writer.NewLine = "\n";
        _writer.WriteLine("officeimo_email_store_verification_v2");
        _writer.WriteLine(string.Concat("semantic_schema\t",
            EmailSemanticComparer.CurrentSchemaVersion.ToString(CultureInfo.InvariantCulture)));
        _writer.WriteLine("digest_algorithm\tHMAC-SHA-256");
        _writer.WriteLine("difference_token_algorithm\tHMAC-SHA-256");
        _writer.WriteLine("ordinal\tassociated\torphaned\tstatus\tsource_digest\tdestination_digest\t" +
            "difference_count\tdifference_path_tokens");
    }

    internal static VerificationManifestWriter? TryCreate(string? destinationPath,
        bool overwrite, EmailSemanticComparisonOptions semanticOptions) {
        if (destinationPath == null) return null;
        if (semanticOptions == null) throw new ArgumentNullException(nameof(semanticOptions));
        byte[]? semanticKey = semanticOptions.CopyDigestKey();
        if (semanticKey == null) {
            throw new InvalidOperationException(
                "A persisted verification manifest requires keyed semantic fingerprints.");
        }
        string fullPath = Path.GetFullPath(destinationPath);
        if (File.Exists(fullPath) && !overwrite) {
            Array.Clear(semanticKey, 0, semanticKey.Length);
            throw new IOException(
                "The verification manifest already exists and overwriteExisting is false.");
        }
        byte[]? pathKey = null;
        try {
            using (var derivation = new HMACSHA256(semanticKey)) {
                pathKey = derivation.ComputeHash(Encoding.UTF8.GetBytes(
                    "OfficeIMO.Email.Store.VerificationManifest.DifferencePath.v2"));
            }
            return new VerificationManifestWriter(fullPath, overwrite, pathKey);
        } finally {
            Array.Clear(semanticKey, 0, semanticKey.Length);
            if (pathKey != null) Array.Clear(pathKey, 0, pathKey.Length);
        }
    }

    internal void Write(int ordinal, bool associated, bool orphaned, string status,
        EmailSemanticComparisonReport? comparison,
        IReadOnlyList<EmailSemanticDifference> differences) {
        string line = string.Join("\t", new[] {
            ordinal.ToString(CultureInfo.InvariantCulture),
            associated ? "1" : "0",
            orphaned ? "1" : "0",
            status,
            comparison?.Source.HexDigest ?? string.Empty,
            comparison?.Destination.HexDigest ?? string.Empty,
            differences.Count.ToString(CultureInfo.InvariantCulture),
            string.Join(",", differences.Select(CreateDifferencePathToken))
        });
        _writer.WriteLine(line);
        byte[] bytes = Encoding.UTF8.GetBytes(string.Concat(line, "\n"));
        _aggregate.TransformBlock(bytes, 0, bytes.Length, bytes, 0);
    }

    internal string Complete(int attempted, int matched, int mismatched, int failed) {
        if (_completed) return _destinationPath;
        _aggregate.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
        string aggregateDigest = BitConverter.ToString(_aggregate.Hash ?? Array.Empty<byte>())
            .Replace("-", string.Empty);
        _writer.WriteLine(string.Join("\t", new[] {
            "summary",
            attempted.ToString(CultureInfo.InvariantCulture),
            matched.ToString(CultureInfo.InvariantCulture),
            mismatched.ToString(CultureInfo.InvariantCulture),
            failed.ToString(CultureInfo.InvariantCulture),
            aggregateDigest
        }));
        _writer.Flush();
        _writer.Dispose();
        OfficeFileCommit.CommitTemporaryFile(_temporaryPath, _destinationPath,
            _overwrite
                ? OfficeFileCommit.ConflictPolicy.Replace
                : OfficeFileCommit.ConflictPolicy.FailIfExists);
        _completed = true;
        return _destinationPath;
    }

    public void Dispose() {
        _writer.Dispose();
        _aggregate.Dispose();
        _differencePathDigest.Dispose();
        if (!_completed) {
            try { if (File.Exists(_temporaryPath)) File.Delete(_temporaryPath); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    private string CreateDifferencePathToken(EmailSemanticDifference difference) {
        byte[] path = Encoding.UTF8.GetBytes(difference.Path);
        try {
            byte[] digest = _differencePathDigest.ComputeHash(path);
            return string.Concat(
                ((int)difference.Kind).ToString(CultureInfo.InvariantCulture),
                ":", BitConverter.ToString(digest).Replace("-", string.Empty));
        } finally {
            Array.Clear(path, 0, path.Length);
        }
    }
}
