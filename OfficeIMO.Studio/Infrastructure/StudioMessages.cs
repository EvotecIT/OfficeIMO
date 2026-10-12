using System.Globalization;
using System.Text;

namespace OfficeIMO.Studio.Infrastructure;

/// <summary>Turns engine identifiers and exception text into wording that suits the interface.</summary>
internal static class StudioMessages {
    /// <summary>
    /// Describes a failure for the interface. Common file-system failures get plain, translated wording with a next
    /// step; engine messages, which are already written for people, keep their text without .NET parameter suffixes.
    /// </summary>
    internal static string Describe(Exception error) {
        ArgumentNullException.ThrowIfNull(error);
        if (DescribeSystemFailure(error) is { } system) return system;
        string message = error.Message;
        if (error is ArgumentException { ParamName: { Length: > 0 } parameter })
            message = message.Replace(" (Parameter '" + parameter + "')", string.Empty, StringComparison.Ordinal);
        return message.Trim();
    }

    private static string? DescribeSystemFailure(Exception error) {
        // Providers and engines raise these types with specific, readable text; only the runtime's generic
        // wording (which also exposes full paths) is replaced.
        string? key = error switch {
            UnauthorizedAccessException when OfficeIMO.Core.Internal.OfficeFileCommit.IsReadOnlyDestination(error) => "Errors.ReadOnly",
            UnauthorizedAccessException when IsRuntimeMessage(error) => "Errors.AccessDenied",
            FileNotFoundException or DirectoryNotFoundException or DriveNotFoundException when IsRuntimeMessage(error) => "Errors.NotFound",
            PathTooLongException when IsRuntimeMessage(error) => "Errors.PathTooLong",
            IOException io when IsFileInUse(io) => "Errors.FileInUse",
            IOException io when IsDiskFull(io) => "Errors.DiskFull",
            _ => null
        };
        return key is null ? null : Localization.StudioLocalization.Current.Get(key);
    }

    private static readonly string[] RuntimeMessagePrefixes = [
        "Access to the path", "Access denied", "Could not find", "The specified path", "The path", "Could not load file", "The system cannot find"
    ];

    private static bool IsRuntimeMessage(Exception error) =>
        RuntimeMessagePrefixes.Any(prefix => error.Message.StartsWith(prefix, StringComparison.Ordinal));

    // Windows reports sharing and lock violations as ERROR_SHARING_VIOLATION (32) and ERROR_LOCK_VIOLATION (33).
    private static bool IsFileInUse(IOException error) =>
        OperatingSystem.IsWindows() && (error.HResult & 0xFFFF) is 32 or 33;

    // ERROR_HANDLE_DISK_FULL (39) and ERROR_DISK_FULL (112) on Windows; ENOSPC (28) elsewhere.
    private static bool IsDiskFull(IOException error) =>
        OperatingSystem.IsWindows() ? (error.HResult & 0xFFFF) is 39 or 112 : error.HResult == 28;

    /// <summary>"render.resource.font-substitution" or "RouteContract" becomes "Font substitution" or "Route contract".</summary>
    internal static string Humanize(string? code) {
        if (string.IsNullOrWhiteSpace(code)) return string.Empty;
        string last = code.Split('.', StringSplitOptions.RemoveEmptyEntries).LastOrDefault() ?? code;
        var words = new List<string>();
        var current = new StringBuilder();
        void Flush() { if (current.Length > 0) { words.Add(current.ToString()); current.Clear(); } }
        for (int index = 0; index < last.Length; index++) {
            char character = last[index];
            if (character is '-' or '_' or ' ') { Flush(); continue; }
            bool startsWord = index > 0 && char.IsUpper(character) &&
                (char.IsLower(last[index - 1]) || index + 1 < last.Length && char.IsLower(last[index + 1]) && char.IsUpper(last[index - 1]));
            if (startsWord) Flush();
            current.Append(character);
        }
        Flush();
        // Acronyms such as PDF stay upper case; other words read as a sentence.
        string text = string.Join(" ", words.Select(word => word.Length > 1 && word.All(char.IsUpper) ? word : word.ToLower(CultureInfo.InvariantCulture)));
        return text.Length == 0 ? code : char.ToUpper(text[0], CultureInfo.CurrentCulture) + text[1..];
    }
}
