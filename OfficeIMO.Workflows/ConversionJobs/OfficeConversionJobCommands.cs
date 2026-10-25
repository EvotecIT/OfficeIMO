namespace OfficeIMO.Workflows;

/// <summary>Copyable commands for a saved job. These helpers format text and never launch a process.</summary>
public static class OfficeConversionJobCommands {
    /// <summary>Returns a CLI command quoted for PowerShell on Windows or a POSIX shell on macOS/Linux.</summary>
    public static string GetCliCommand(string jobPath) => "officeimo workflow job run " + Quote(jobPath, !OperatingSystem.IsWindows());
    /// <summary>Returns a PowerShell command with a literal quoted job path.</summary>
    public static string GetPowerShellCommand(string jobPath) => "Invoke-OfficeConversionJob -Path " + Quote(jobPath);
    private static string Quote(string path, bool posix = false) {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        if (path.Any(char.IsControl)) throw new ArgumentException("A copied job path cannot contain control characters.", nameof(path));
        if (posix) return "'" + path.Replace("'", "'\\''") + "'";
        var quoted = new System.Text.StringBuilder("'");
        foreach (char character in path) {
            quoted.Append(character);
            if (character is '\'' or '\u2018' or '\u2019' or '\u201A' or '\u201B') quoted.Append(character);
        }
        return quoted.Append('\'').ToString();
    }
}
