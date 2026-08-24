using OfficeIMO.Internal;
using System.Runtime.InteropServices;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public class PathIdentityContracts {
    [Fact]
    public void OrdinaryFilesHaveOnePhysicalLink() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-path-identity-" + Guid.NewGuid().ToString("N") + ".txt");
        try {
            File.WriteAllText(path, "identity");
            Assert.False(OfficePathIdentity.HasMultipleLinks(path));
            using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            OfficeFileMetadata byPath = OfficePathIdentity.GetMetadata(path);
            OfficeFileMetadata byHandle = OfficePathIdentity.GetMetadata(path, stream.SafeFileHandle);
            Assert.Equal(1UL, byPath.LinkCount);
            Assert.Equal(1UL, byHandle.LinkCount);
            Assert.Equal(byPath.Identity, byHandle.Identity);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void HardLinksSharePhysicalIdentity() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-path-identity-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string source = Path.Combine(directory, "source.txt");
        string alias = Path.Combine(directory, "alias.txt");
        try {
            File.WriteAllText(source, "identity");
            CreateHardLink(alias, source);
            Assert.True(OfficePathIdentity.AreEquivalent(source, alias));
            Assert.True(OfficePathIdentity.HasMultipleLinks(source));
            Assert.True(OfficePathIdentity.HasMultipleLinks(alias));
        } finally {
            if (File.Exists(alias)) File.Delete(alias);
            if (File.Exists(source)) File.Delete(source);
            if (Directory.Exists(directory)) Directory.Delete(directory);
        }
    }

    [Fact]
    public void WindowsLocalAndUncAliasesShareIdentityWhenAvailable() {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-path-identity-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string local = Path.Combine(directory, "future", "report.json");
            string root = Path.GetPathRoot(local)!;
            if (root.Length < 2 || root[1] != ':') return;
            string unc = "\\\\localhost\\" + char.ToLowerInvariant(root[0]) + "$" + local.Substring(2);
            string uncDirectory = Path.GetDirectoryName(unc)!;
            string uncExisting = Path.GetDirectoryName(uncDirectory)!;
            if (!Directory.Exists(uncExisting)) return;
            Assert.True(OfficePathIdentity.AreEquivalent(local, unc));
        } finally {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void DanglingLinksNeverCompareAsIndependentMissingPaths() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-path-identity-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string target = Path.Combine(directory, "future-target");
        string link = Path.Combine(directory, "future-link");
        try {
            if (!TryCreateSymbolicLink(link, target)) return;
            string linkedReport = Path.Combine(link, "report.json");
            string targetReport = Path.Combine(target, "report.json");
            try {
                Assert.True(OfficePathIdentity.AreEquivalent(linkedReport, targetReport));
            } catch (IOException) {
                // Older runtimes may not expose a safe target decoder. Indeterminate is fail-closed.
            }
        } finally {
            try { File.Delete(link); } catch (IOException) { }
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void MacMissingTailsHonorUnicodeNormalization() {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) return;
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-path-identity-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string composed = Path.Combine(directory, "caf\u00e9", "report.json");
            string decomposed = Path.Combine(directory, "cafe\u0301", "report.json");
            Assert.True(OfficePathIdentity.AreEquivalent(composed, decomposed));
        } finally {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void LegacyAndExtendedWindowsIdentitiesUseTheirSharedFileIndex() {
        var extended = OfficePhysicalFileIdentity.CreateWindowsExtended(string.Empty, 7, 101, 202, 11);
        var legacy = OfficePhysicalFileIdentity.CreateWindowsLegacy(string.Empty, 7, 11);
        var other = OfficePhysicalFileIdentity.CreateWindowsLegacy(string.Empty, 7, 12);

        Assert.Equal(extended, legacy);
        Assert.True(extended.HasSameNumericIdentity(legacy));
        Assert.Equal(extended.ToStableKey(), legacy.ToStableKey());
        Assert.NotEqual(extended, other);
    }

    [Fact]
    public void LinuxFilesystemClassificationDoesNotGuessForNtfsOrUnknownDrivers() {
        Assert.True(OfficePathIdentity.TryClassifyLinuxFileSystemCaseBehavior(0x2011bab0, out bool exFat));
        Assert.True(exFat);
        Assert.True(OfficePathIdentity.TryClassifyLinuxFileSystemCaseBehavior(0x0000ef53, out bool ext));
        Assert.False(ext);

        Assert.False(OfficePathIdentity.TryClassifyLinuxFileSystemCaseBehavior(0x5346544e, out bool ntfs));
        Assert.False(ntfs);
        Assert.False(OfficePathIdentity.TryClassifyLinuxFileSystemCaseBehavior(0x12345678, out bool unknown));
        Assert.False(unknown);
    }

    [Fact]
    public void LinuxStatxFallbackCoversPolicyAndCapabilityFailures() {
        Assert.True(OfficePathIdentity.ShouldFallbackFromLinuxStatx(1));
        Assert.True(OfficePathIdentity.ShouldFallbackFromLinuxStatx(13));
        Assert.True(OfficePathIdentity.ShouldFallbackFromLinuxStatx(22));
        Assert.True(OfficePathIdentity.ShouldFallbackFromLinuxStatx(38));
        Assert.True(OfficePathIdentity.ShouldFallbackFromLinuxStatx(95));

        Assert.False(OfficePathIdentity.ShouldFallbackFromLinuxStatx(2));
        Assert.False(OfficePathIdentity.ShouldFallbackFromLinuxStatx(5));

        Assert.True(OfficePathIdentity.HasRequiredLinuxStatxMetadata(0x00000105));
        Assert.False(OfficePathIdentity.HasRequiredLinuxStatxMetadata(0x00000101));
    }

    private static void CreateHardLink(string alias, string source) {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            if (!CreateHardLinkWindows(alias, source, IntPtr.Zero)) {
                throw new IOException("Unable to create a Windows hard link for the path identity contract.");
            }
        } else if (Link(source, alias) != 0) {
            throw new IOException("Unable to create a POSIX hard link for the path identity contract.");
        }
    }

    private static bool TryCreateSymbolicLink(string link, string target) {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            const int allowUnprivilegedCreate = 0x2;
            return CreateSymbolicLinkWindows(link, target, allowUnprivilegedCreate);
        }
        return CreateSymbolicLinkUnix(target, link) == 0;
    }

    [DllImport("kernel32.dll", EntryPoint = "CreateHardLinkW", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CreateHardLinkWindows(string fileName, string existingFileName, IntPtr securityAttributes);

    [DllImport("kernel32.dll", EntryPoint = "CreateSymbolicLinkW", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CreateSymbolicLinkWindows(string symlinkFileName, string targetFileName, int flags);

    [DllImport("libc", EntryPoint = "link", CharSet = CharSet.Ansi, SetLastError = true)]
    private static extern int Link(string existingPath, string newPath);

    [DllImport("libc", EntryPoint = "symlink", CharSet = CharSet.Ansi, SetLastError = true)]
    private static extern int CreateSymbolicLinkUnix(string target, string linkPath);
}
