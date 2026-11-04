using System.Management.Automation;
using System.Management.Automation.Runspaces;

namespace PSWriteOffice.Tests;

/// <summary>Creates task-owned directory aliases without requiring Windows symbolic-link privileges.</summary>
internal static class PhysicalPathTestSupport {
    internal static IDisposable CreateDirectoryAlias(Runspace runspace, string alias, string target) {
        if (!OperatingSystem.IsWindows()) {
            Directory.CreateSymbolicLink(alias, target);
            return new DirectoryAlias(alias);
        }
        runspace.SessionStateProxy.SetVariable("aliasDirectory", alias);
        runspace.SessionStateProxy.SetVariable("targetDirectory", target);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("New-Item -ItemType Junction -Path $aliasDirectory -Target $targetDirectory -ErrorAction Stop | Out-Null");
        shell.Invoke();
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
        Assert.True(Directory.Exists(alias));
        return new DirectoryAlias(alias);
    }

    // Remove the link before recursive fixture cleanup removes its target.
    private sealed class DirectoryAlias(string path) : IDisposable {
        public void Dispose() {
            Assert.True((File.GetAttributes(path) & FileAttributes.ReparsePoint) != 0);
            Directory.Delete(path);
        }
    }
}
