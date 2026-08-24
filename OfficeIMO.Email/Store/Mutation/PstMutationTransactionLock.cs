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
    private readonly FileStream _pathLockStream;
    private readonly FileStream _identityLockStream;
    private readonly string _identity;
    private bool _disposed;

    private PstMutationTransactionLock(FileStream pathLockStream, FileStream identityLockStream, string identity) {
        _pathLockStream = pathLockStream;
        _identityLockStream = identityLockStream;
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

        FileStream? pathLockStream = null;
        FileStream? identityLockStream = null;
        try {
            string physicalPath = EmailStorePathIdentity.ResolvePhysicalPath(sourcePath);
            string lockDirectory = Path.GetDirectoryName(physicalPath) ?? Directory.GetCurrentDirectory();
            string lockPath = Path.Combine(lockDirectory, lockName);
            pathLockStream = OpenLockStream(lockPath, sharedAcrossUsers: true);

            string identityLockRoot = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            if (string.IsNullOrWhiteSpace(identityLockRoot)) identityLockRoot = Path.GetTempPath();
            string identityLockDirectory = Path.Combine(identityLockRoot, "OfficeIMO", "PstMutationLocks");
            Directory.CreateDirectory(identityLockDirectory);
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows) &&
                ChangeMode(identityLockDirectory, 0x1c0U) != 0) {
                throw new IOException("The per-user OfficeIMO PST mutation lock directory could not be secured " +
                    "(OS error " + Marshal.GetLastWin32Error() + ").");
            }
            identityLockStream = OpenLockStream(Path.Combine(identityLockDirectory, lockName),
                sharedAcrossUsers: false);

            string currentIdentity = EmailStorePathIdentity.GetPhysicalIdentityKey(sourcePath);
            if (!string.Equals(identity, currentIdentity, StringComparison.Ordinal)) {
                throw new IOException(
                    "The source PST path changed while its mutation lock was being acquired.");
            }
            return new PstMutationTransactionLock(pathLockStream, identityLockStream, identity);
        } catch (UnauthorizedAccessException exception) {
            identityLockStream?.Dispose();
            pathLockStream?.Dispose();
            ProcessLocks.TryRemove(identity, out _);
            throw new IOException(
                "The adjacent OfficeIMO PST mutation lock could not be created.", exception);
        } catch (IOException exception) {
            identityLockStream?.Dispose();
            pathLockStream?.Dispose();
            ProcessLocks.TryRemove(identity, out _);
            throw new IOException("Another OfficeIMO mutation transaction already owns this PST path.", exception);
        }
    }

    private static FileStream OpenLockStream(string lockPath, bool sharedAcrossUsers) {
        var stream = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite,
            FileShare.None, 1, FileOptions.RandomAccess);
        if (!sharedAcrossUsers || RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return stream;
        if (ChangeMode(lockPath, 0x1b6U) == 0) return stream;
        int error = Marshal.GetLastWin32Error();
        stream.Dispose();
        throw new IOException("The adjacent OfficeIMO PST mutation lock could not be shared " +
            "(OS error " + error + ").");
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        try {
            try {
                _identityLockStream.Dispose();
            } finally {
                _pathLockStream.Dispose();
            }
        } finally {
            ProcessLocks.TryRemove(_identity, out _);
        }
    }

    [DllImport("libc", EntryPoint = "chmod", CharSet = CharSet.Ansi, SetLastError = true)]
    private static extern int ChangeMode(string path, uint mode);
}
