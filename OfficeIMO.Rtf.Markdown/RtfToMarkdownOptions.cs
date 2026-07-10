using System;
using System.Collections.Generic;
using OfficeIMO.Markdown;
using OfficeIMO.Rtf;

namespace OfficeIMO.Rtf.Markdown;

/// <summary>
/// Controls semantic conversion from an OfficeIMO RTF document into Markdown.
/// </summary>
public sealed class RtfToMarkdownOptions {
    private readonly List<RtfMarkdownConversionDiagnostic> _diagnostics = new List<RtfMarkdownConversionDiagnostic>();

    /// <summary>
    /// Markdown writer options used when rendering the resulting Markdown document.
    /// </summary>
    public MarkdownWriteOptions? MarkdownWriteOptions { get; set; }

    /// <summary>
    /// Markdown reader options used when building inline Markdown sequences.
    /// </summary>
    public MarkdownReaderOptions? InlineReaderOptions { get; set; }

    /// <summary>
    /// Includes RTF hidden text in generated Markdown when set to true.
    /// </summary>
    public bool IncludeHiddenText { get; set; }

    /// <summary>
    /// Emits HTML comments for unsupported RTF block features that Markdown cannot represent directly.
    /// </summary>
    public bool EmitUnsupportedHtmlComments { get; set; } = true;

    /// <summary>
    /// Creates an image path for each RTF image encountered during conversion.
    /// </summary>
    public Func<RtfImage, int, string>? ImagePathFactory { get; set; }

    /// <summary>
    /// Exports each encountered RTF image payload to the logical path selected by
    /// <see cref="ImagePathFactory"/>. The callback receives the unescaped path; exceptions propagate.
    /// </summary>
    public Action<RtfImage, int, string>? ImageExporter { get; set; }

    /// <summary>
    /// Receives diagnostics as they are emitted.
    /// </summary>
    public Action<RtfMarkdownConversionDiagnostic>? DiagnosticHandler { get; set; }

    /// <summary>
    /// Diagnostics produced by this conversion options instance.
    /// </summary>
    public IReadOnlyList<RtfMarkdownConversionDiagnostic> Diagnostics => _diagnostics;

    /// <summary>Shared cross-adapter fidelity report for this conversion.</summary>
    public RtfConversionReport ConversionReport { get; } = new RtfConversionReport();

    internal RtfMarkdownNoteRegistry? NoteRegistry { get; set; }

    internal void Report(string code, RtfMarkdownDiagnosticSeverity severity, string message, string? source = null) {
        var diagnostic = new RtfMarkdownConversionDiagnostic(code, severity, message, source);
        _diagnostics.Add(diagnostic);
        RtfMarkdownConversionReportMapper.Add(ConversionReport, diagnostic);
        DiagnosticHandler?.Invoke(diagnostic);
    }
}

internal sealed class RtfMarkdownNoteRegistry {
    private readonly Dictionary<RtfNote, string> _labels = new Dictionary<RtfNote, string>();
    private readonly List<KeyValuePair<string, RtfNote>> _ordered = new List<KeyValuePair<string, RtfNote>>();
    private int _footnoteIndex;
    private int _endnoteIndex;

    internal IReadOnlyList<KeyValuePair<string, RtfNote>> Ordered => _ordered;

    internal string? Register(RtfNote note, RtfToMarkdownOptions options) {
        if (note.Kind == RtfNoteKind.Annotation) {
            options.Report("RTFMD012", RtfMarkdownDiagnosticSeverity.Warning, "RTF annotation omitted from Markdown output.", note.Id ?? note.Kind.ToString());
            return null;
        }

        if (_labels.TryGetValue(note, out string? existing)) return existing;
        string label = note.Kind == RtfNoteKind.Endnote
            ? "en" + (++_endnoteIndex).ToString(System.Globalization.CultureInfo.InvariantCulture)
            : "fn" + (++_footnoteIndex).ToString(System.Globalization.CultureInfo.InvariantCulture);
        _labels.Add(note, label);
        _ordered.Add(new KeyValuePair<string, RtfNote>(label, note));
        return label;
    }
}
