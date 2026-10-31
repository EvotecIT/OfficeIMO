using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Win32.SafeHandles;
using Xunit;
using CoreCommitClaim = OfficeIMO.Core.Internal.IO.OfficeFileCommitClaim;
using ReaderCommitClaim = OfficeIMO.Reader.Internal.IO.OfficeFileCommitClaim;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderWindowsFileCommitClaimTests {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SharedClaimAcquisitionPreservesExistingMarkedWindowsFile(bool coreClaim) {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-stale-claim-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        string destination = Path.Combine(folder, "publication.bin");
        string claimPath = CoreCommitClaim.CreatePath(destination);
        try {
            byte[] prefix = System.Text.Encoding.ASCII.GetBytes("OfficeIMO-commit-claim-v1:");
            byte[] marker = new byte[prefix.Length + 16];
            Buffer.BlockCopy(prefix, 0, marker, 0, prefix.Length);
            Guid.NewGuid().ToByteArray().CopyTo(marker, prefix.Length);
            File.WriteAllBytes(claimPath, marker);

            Exception? error = Record.Exception(() => { using (Acquire(coreClaim, destination)) { } });
            Assert.True(error is IOException or UnauthorizedAccessException, error?.ToString());
            Assert.True(coreClaim
                ? CoreCommitClaim.IsContention(error!, destination)
                : ReaderCommitClaim.IsContention(error!, destination));
            Assert.Equal(marker, File.ReadAllBytes(claimPath));
            Assert.False(File.Exists(destination));
        } finally {
            Directory.Delete(folder, recursive: true);
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SharedClaimAcquisitionRecoversWhenWindowsCompletesPendingDeletion(bool coreClaim) {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-pending-claim-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        string destination = Path.Combine(folder, "publication.bin");
        string claimPath = CoreCommitClaim.CreatePath(destination);
        SafeFileHandle? pendingDeletion = null;
        int acquiringThread = Thread.CurrentThread.ManagedThreadId;
        int pendingDeletionFailures = 0;
        EventHandler<FirstChanceExceptionEventArgs> releasePendingDeletion = (_, args) => {
            if (Thread.CurrentThread.ManagedThreadId != acquiringThread ||
                args.Exception is not UnauthorizedAccessException ||
                args.Exception.HResult != unchecked((int)0x80070005)) return;

            // Release only after CreateNew has encountered the real Windows pending-delete
            // state. The original exception still reaches Acquire, so its retry is required.
            SafeFileHandle? handle = Interlocked.Exchange(ref pendingDeletion, null);
            if (handle == null) return;
            Interlocked.Increment(ref pendingDeletionFailures);
            handle.Dispose();
        };

        try {
            pendingDeletion = CreatePendingDeletion(claimPath);
            UnauthorizedAccessException pendingError = Assert.Throws<UnauthorizedAccessException>(() => {
                using (new FileStream(claimPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None)) { }
            });
            Assert.Equal(unchecked((int)0x80070005), pendingError.HResult);
            AppDomain.CurrentDomain.FirstChanceException += releasePendingDeletion;
            using (Acquire(coreClaim, destination)) {
                Assert.Equal(1, pendingDeletionFailures);
                Assert.True(File.Exists(claimPath));
                Assert.False(File.Exists(destination));
            }
            Assert.Empty(Directory.GetFileSystemEntries(folder));
        } finally {
            AppDomain.CurrentDomain.FirstChanceException -= releasePendingDeletion;
            pendingDeletion?.Dispose();
            Directory.Delete(folder, recursive: true);
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SharedClaimAcquisitionPreservesPersistentWindowsAccessFailures(bool coreClaim) {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-denied-claim-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        string destination = Path.Combine(folder, "publication.bin");
        string claimPath = CoreCommitClaim.CreatePath(destination);
        Directory.CreateDirectory(claimPath);
        try {
            UnauthorizedAccessException error = Assert.Throws<UnauthorizedAccessException>(() => Acquire(coreClaim, destination));

            Assert.Equal(unchecked((int)0x80070005), error.HResult);
            Assert.False(coreClaim
                ? CoreCommitClaim.IsContention(error, destination)
                : ReaderCommitClaim.IsContention(error, destination));
            Assert.True(Directory.Exists(claimPath));
            Assert.False(File.Exists(destination));
        } finally {
            Directory.Delete(folder, recursive: true);
        }
    }

    private static IDisposable Acquire(bool coreClaim, string destination) => coreClaim
        ? (IDisposable)CoreCommitClaim.Acquire(destination)
        : ReaderCommitClaim.Acquire(destination);

    private static SafeFileHandle CreatePendingDeletion(string path) {
        const uint deleteAccess = 0x00010000;
        const uint createNew = 1;
        const uint fileAttributeNormal = 0x00000080;
        const int fileDispositionInfo = 4;
        SafeFileHandle handle = CreateFile(path, deleteAccess, FileShare.ReadWrite | FileShare.Delete,
            IntPtr.Zero, createNew, fileAttributeNormal, IntPtr.Zero);
        if (handle.IsInvalid) {
            int error = Marshal.GetLastWin32Error();
            handle.Dispose();
            throw new Win32Exception(error, "Unable to create the pending-deletion claim fixture.");
        }

        try {
            // Request classic disposition explicitly: DeleteFile can use POSIX semantics
            // and unlink the name immediately even while another handle remains open.
            var disposition = new FileDispositionInfo { DeleteFile = 1 };
            if (!SetFileInformationByHandle(handle, fileDispositionInfo, ref disposition,
                    (uint)Marshal.SizeOf(typeof(FileDispositionInfo)))) {
                int error = Marshal.GetLastWin32Error();
                throw new Win32Exception(error, "Unable to mark the claim fixture for pending deletion.");
            }
            return handle;
        } catch {
            handle.Dispose();
            throw;
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct FileDispositionInfo {
        // Native BOOLEAN is one byte, unlike a Win32 BOOL.
        internal byte DeleteFile;
    }

    [DllImport("kernel32.dll", EntryPoint = "CreateFileW", CharSet = CharSet.Unicode,
        ExactSpelling = true, SetLastError = true)]
    private static extern SafeFileHandle CreateFile(string fileName, uint desiredAccess,
        FileShare shareMode, IntPtr securityAttributes, uint creationDisposition,
        uint flagsAndAttributes, IntPtr templateFile);

    [DllImport("kernel32.dll", ExactSpelling = true, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetFileInformationByHandle(SafeFileHandle file,
        int fileInformationClass, ref FileDispositionInfo fileInformation, uint bufferSize);
}
