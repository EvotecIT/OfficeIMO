using System;
using System.IO;

namespace OfficeIMO.Core.Internal;

/// <summary>Creates a bounded portable file-name stem from an untrusted display name.</summary>
internal static class OfficePortableFileName {
    private const string PortableInvalidCharacters = "<>:\"/\\|?*";
    private static readonly char[] PlatformInvalidCharacters = Path.GetInvalidFileNameChars();

    internal static string SanitizeBaseName(string name, int maximumLength = 120) {
        if (name == null) throw new ArgumentNullException(nameof(name));
        if (maximumLength < 1) throw new ArgumentOutOfRangeException(nameof(maximumLength));
        char[] chars = name.ToCharArray();
        for (int index = 0; index < chars.Length; index++) {
            if (chars[index] < 32 || PortableInvalidCharacters.IndexOf(chars[index]) >= 0 ||
                Array.IndexOf(PlatformInvalidCharacters, chars[index]) >= 0) chars[index] = '_';
        }

        string sanitized = new string(chars).Trim().TrimEnd('.', ' ');
        if (sanitized.Length > maximumLength) {
            int length = maximumLength;
            if (char.IsHighSurrogate(sanitized[length - 1])) length--;
            sanitized = sanitized.Substring(0, length).TrimEnd('.', ' ');
        }
        if (IsReservedWindowsFileName(sanitized)) sanitized = "_" + sanitized;
        return sanitized;
    }

    private static bool IsReservedWindowsFileName(string name) {
        if (string.IsNullOrWhiteSpace(name)) return false;
        string candidate = name;
        int dot = candidate.IndexOf('.');
        if (dot >= 0) candidate = candidate.Substring(0, dot);
        if (candidate.Equals("CON", StringComparison.OrdinalIgnoreCase) ||
            candidate.Equals("PRN", StringComparison.OrdinalIgnoreCase) ||
            candidate.Equals("AUX", StringComparison.OrdinalIgnoreCase) ||
            candidate.Equals("NUL", StringComparison.OrdinalIgnoreCase)) return true;
        if (candidate.Length == 4 &&
            (candidate.StartsWith("COM", StringComparison.OrdinalIgnoreCase) ||
             candidate.StartsWith("LPT", StringComparison.OrdinalIgnoreCase))) {
            char digit = candidate[3];
            return digit >= '1' && digit <= '9' || digit is '\u00B9' or '\u00B2' or '\u00B3';
        }
        return false;
    }
}
