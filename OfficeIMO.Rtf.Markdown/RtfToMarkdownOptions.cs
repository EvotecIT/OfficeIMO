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
