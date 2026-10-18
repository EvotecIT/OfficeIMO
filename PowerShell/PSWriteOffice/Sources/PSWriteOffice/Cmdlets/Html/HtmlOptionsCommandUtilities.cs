using System;
using System.IO;
using System.Management.Automation;

namespace PSWriteOffice.Cmdlets.Html;

internal static class HtmlOptionsCommandUtilities {
    internal static Uri NormalizeBaseUri(SessionState sessionState, string value) {
        if (Uri.TryCreate(value, UriKind.Absolute, out Uri? absoluteUri) &&
            !absoluteUri.IsFile &&
            absoluteUri.Scheme.Length > 1) {
            return absoluteUri;
        }

        string providerPath = absoluteUri?.IsFile == true
            ? absoluteUri.LocalPath
            : sessionState.Path.GetUnresolvedProviderPathFromPSPath(value);
        string fullPath = Path.GetFullPath(providerPath);
        bool isDirectory = Directory.Exists(fullPath) ||
            value.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal) ||
            value.EndsWith(Path.AltDirectorySeparatorChar.ToString(), StringComparison.Ordinal);
        if (isDirectory) {
            fullPath = fullPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
        }

        return new Uri(fullPath, UriKind.Absolute);
    }
}
