using System;
using System.Collections.Generic;
using OfficeIMO.Markdown;

namespace OfficeIMO.Rtf.Markdown;

/// <summary>
/// Controls semantic conversion from Markdown into an OfficeIMO RTF document.
/// </summary>
public sealed class MarkdownToRtfOptions {
    private readonly List<RtfMarkdownConversionDiagnostic> _diagnostics = new List<RtfMarkdownConversionDiagnostic>();

    /// <summary>
    /// Markdown reader options used when parsing Markdown text.
    /// </summary>
    public MarkdownReaderOptions? ReaderOptions { get; set; }

    /// <summary>
    /// Preserves raw HTML as visible text. When false, raw HTML is omitted with a diagnostic.
    /// </summary>
    public bool PreserveRawHtmlAsText { get; set; }

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

    internal void Report(string code, RtfMarkdownDiagnosticSeverity severity, string message, string? source = null) {
        var diagnostic = new RtfMarkdownConversionDiagnostic(code, severity, message, source);
        _diagnostics.Add(diagnostic);
        RtfMarkdownConversionReportMapper.Add(ConversionReport, diagnostic);
        DiagnosticHandler?.Invoke(diagnostic);
    }
}
