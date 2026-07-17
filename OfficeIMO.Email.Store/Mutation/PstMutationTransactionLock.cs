using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Email.Store;

/// <summary>Coordinates OfficeIMO mutation transactions for one PST path across processes.</summary>
internal sealed class PstMutationTransactionLock : IDisposable {
    private static readonly ConcurrentDictionary<string, byte> ProcessLocks =
        new ConcurrentDictionary<string, byte>(StringComparer.Ordinal);
    private readonly FileStream _lockStream;
    private readonly string _identity;
    private bool _disposed;

    private PstMutationTransactionLock(FileStream lockStream, string identity) {
        _lockStream = lockStream;
        _identity = identity;
    }

    internal static PstMutationTransactionLock Acquire(string sourcePath) {
        string identity = Path.GetFullPath(sourcePath);
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) identity = identity.ToUpperInvariant();
        if (!ProcessLocks.TryAdd(identity, 0)) {
            throw new IOException("Another OfficeIMO mutation transaction already owns this PST path.");
        }
        byte[] digest;
        using (SHA256 sha = SHA256.Create()) digest = sha.ComputeHash(Encoding.UTF8.GetBytes(identity));
        string lockName = BitConverter.ToString(digest).Replace("-", string.Empty) + ".lock";
        Array.Clear(digest, 0, digest.Length);

        try {
            string lockDirectory = Path.Combine(Path.GetTempPath(), "OfficeIMO", "PstMutationLocks");
            Directory.CreateDirectory(lockDirectory);
            string lockPath = Path.Combine(lockDirectory, lockName);
            var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite,
                FileShare.None, 1, FileOptions.RandomAccess);
            return new PstMutationTransactionLock(lockStream, identity);
        } catch (Exception exception) when (exception is IOException || exception is UnauthorizedAccessException) {
            ProcessLocks.TryRemove(identity, out _);
            throw new IOException("Another OfficeIMO mutation transaction already owns this PST path.", exception);
        }
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        try {
            _lockStream.Dispose();
        } finally {
            ProcessLocks.TryRemove(_identity, out _);
        }
    }
}
