using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Win32.SafeHandles;

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

    internal string Identity => _identity;

    internal static PstMutationTransactionLock Acquire(string sourcePath, SafeFileHandle sourceHandle) {
        string identity = EmailStorePathIdentity.GetPhysicalIdentityKey(sourcePath, sourceHandle);
        if (!ProcessLocks.TryAdd(identity, 0)) {
            throw new IOException("Another OfficeIMO mutation transaction already owns this PST path.");
        }
        byte[] digest;
        using (SHA256 sha = SHA256.Create()) digest = sha.ComputeHash(Encoding.UTF8.GetBytes(identity));
        string lockName = ".officeimo-pst-" +
            BitConverter.ToString(digest).Replace("-", string.Empty) + ".lock";
        Array.Clear(digest, 0, digest.Length);

        try {
            string lockRoot = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            if (string.IsNullOrWhiteSpace(lockRoot)) lockRoot = Path.GetTempPath();
            string lockDirectory = Path.Combine(lockRoot, "OfficeIMO", "PstMutationLocks");
            Directory.CreateDirectory(lockDirectory);
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows) && ChangeMode(lockDirectory, 0x1c0U) != 0) {
                throw new IOException("The per-user OfficeIMO PST mutation lock directory could not be secured " +
                    "(OS error " + Marshal.GetLastWin32Error() + ").");
            }
            string lockPath = Path.Combine(lockDirectory, lockName);
            var lockStream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite,
                FileShare.None, 1, FileOptions.RandomAccess);
            string currentIdentity = EmailStorePathIdentity.GetPhysicalIdentityKey(sourcePath);
            if (!string.Equals(identity, currentIdentity, StringComparison.Ordinal)) {
                lockStream.Dispose();
                throw new IOException(
                    "The source PST path changed while its mutation lock was being acquired.");
            }
            return new PstMutationTransactionLock(lockStream, identity);
        } catch (UnauthorizedAccessException exception) {
            ProcessLocks.TryRemove(identity, out _);
            throw new IOException(
                "The adjacent OfficeIMO PST mutation lock could not be created.", exception);
        } catch (IOException exception) {
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

    [DllImport("libc", EntryPoint = "chmod", CharSet = CharSet.Ansi, SetLastError = true)]
    private static extern int ChangeMode(string path, uint mode);
}
