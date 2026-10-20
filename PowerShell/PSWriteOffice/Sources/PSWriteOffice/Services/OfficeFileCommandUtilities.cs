using System;
using System.Management.Automation;

namespace PSWriteOffice.Services;

/// <summary>Resolves literal filesystem input at the PowerShell provider boundary.</summary>
internal static class OfficeFileCommandUtilities {
    internal static string ResolvePath(PSCmdlet cmdlet, string path) {
        if (string.IsNullOrWhiteSpace(path)) throw new PSArgumentException("A filesystem path is required.", nameof(path));
        var result = cmdlet.SessionState.Path.GetUnresolvedProviderPathFromPSPath(path, out var provider, out _);
        if (!string.Equals(provider.Name, "FileSystem", StringComparison.OrdinalIgnoreCase))
            throw new PSArgumentException("Only filesystem paths are supported.", nameof(path));
        return result;
    }
}
