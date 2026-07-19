using System;

namespace OfficeIMO.Reader;

/// <summary>
/// Extracts logical names from archive entries, attachments, and other identifiers that are not
/// required to be valid paths on the current operating system.
/// </summary>
internal static class ReaderLogicalPath {
    internal static string GetFileName(string value) {
        if (value == null) throw new ArgumentNullException(nameof(value));

        int separator = Math.Max(value.LastIndexOf('/'), value.LastIndexOf('\\'));
        return separator < 0 ? value : value.Substring(separator + 1);
    }

    internal static string GetExtension(string value) {
        string name = GetFileName(value);
        int dot = name.LastIndexOf('.');
        return dot <= 0 ? string.Empty : name.Substring(dot);
    }
}
