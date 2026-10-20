using System;
using System.IO;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Threading;
using Xunit;
using CoreCommitClaim = OfficeIMO.Core.Internal.IO.OfficeFileCommitClaim;
using ReaderCommitClaim = OfficeIMO.Reader.Internal.IO.OfficeFileCommitClaim;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderWindowsFileCommitClaimTests {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SharedClaimAcquisitionRecoversWhenWindowsCompletesPendingDeletion(bool coreClaim) {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-pending-claim-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        string destination = Path.Combine(folder, "publication.bin");
        string claimPath = CoreCommitClaim.CreatePath(destination);
        FileStream? pendingDeletion = null;
        int acquiringThread = Thread.CurrentThread.ManagedThreadId;
        int pendingDeletionFailures = 0;
        EventHandler<FirstChanceExceptionEventArgs> releasePendingDeletion = (_, args) => {
            if (Thread.CurrentThread.ManagedThreadId != acquiringThread ||
                args.Exception is not UnauthorizedAccessException ||
                args.Exception.HResult != unchecked((int)0x80070005)) return;

            // Release only after CreateNew has encountered the real Windows pending-delete
            // state. The original exception still reaches Acquire, so its retry is required.
            FileStream? handle = Interlocked.Exchange(ref pendingDeletion, null);
            if (handle == null) return;
            Interlocked.Increment(ref pendingDeletionFailures);
            handle.Dispose();
        };

        try {
            pendingDeletion = new FileStream(claimPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.Delete);
            File.Delete(claimPath);
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
}
