namespace OfficeIMO.Markdown;

using System;

/// <summary>
/// Explicit transcript transport-marker helpers for hosts ingesting or exporting IntelligenceX-style markdown.
/// </summary>
public static class MarkdownTranscriptTransportMarkers {
    /// <summary>
    /// Removes the cached-evidence transport marker emitted in IntelligenceX transcript flows while preserving
    /// surrounding content and line-ending style.
    /// </summary>
    /// <param name="markdown">Transcript markdown source.</param>
    /// <returns>Transcript markdown without cached-evidence transport marker lines.</returns>
    public static string StripIntelligenceXCachedEvidenceTransportMarkers(string? markdown) {
        var value = markdown ?? string.Empty;
        if (value.Length == 0) {
            return string.Empty;
        }

        var newline = value.Contains("\r\n") ? "\r\n" : "\n";
        var normalized = value.Replace("\r\n", "\n").Replace('\r', '\n');
        var lines = normalized.Split('\n');
        for (var i = 0; i < lines.Length; i++) {
            var line = lines[i] ?? string.Empty;
            if (line.Trim().Equals("ix:cached-tool-evidence:v1", StringComparison.OrdinalIgnoreCase)) {
                lines[i] = string.Empty;
            }
        }

        return string.Join(newline, lines);
    }
}
