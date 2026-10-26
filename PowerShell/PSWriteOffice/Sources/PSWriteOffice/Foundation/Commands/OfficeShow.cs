using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

namespace PSWriteOffice.Foundation;

/// <summary>Opens a saved result using the platform's registered default application.</summary>
public static class OfficeShow {
    /// <summary>Opens one existing file without composing a shell command from its path.</summary>
    public static void Open(string path) {
        using var process = Process.Start(CreateStartInfo(path));
    }
    internal static ProcessStartInfo CreateStartInfo(string path) {
        string full = Path.GetFullPath(path);
        if (!File.Exists(full)) throw new FileNotFoundException("The saved output does not exist.", full);
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return new ProcessStartInfo(full) { UseShellExecute = true };
        var start = new ProcessStartInfo(RuntimeInformation.IsOSPlatform(OSPlatform.OSX) ? "/usr/bin/open" : "xdg-open") { UseShellExecute = false };
#if FRAMEWORK
        start.Arguments = "\"" + full.Replace("\"", "\\\"") + "\"";
#else
        start.ArgumentList.Add(full);
#endif
        return start;
    }
}
