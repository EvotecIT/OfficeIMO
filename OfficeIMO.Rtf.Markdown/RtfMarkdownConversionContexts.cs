using OfficeIMO.Markdown;
using OfficeIMO.Rtf;

namespace OfficeIMO.Rtf.Markdown;

internal enum RtfMarkdownDiagnosticSeverity {
    Info,
    Warning,
    Error
}

internal abstract class RtfMarkdownConversionContext {
    internal RtfConversionReport ConversionReport { get; } = new RtfConversionReport();

    internal void AddDiagnostic(
        string code,
        RtfMarkdownDiagnosticSeverity severity,
        string message,
        string? source = null,
        RtfConversionAction action = RtfConversionAction.Preserved) {
        ConversionReport.Add(
            severity == RtfMarkdownDiagnosticSeverity.Error
                ? RtfConversionSeverity.Error
                : severity == RtfMarkdownDiagnosticSeverity.Warning
                    ? RtfConversionSeverity.Warning
                    : RtfConversionSeverity.Information,
            code,
            message,
            action,
            sourcePath: source,
            feature: source);
    }
}

internal sealed class RtfToMarkdownConversionContext : RtfMarkdownConversionContext {
    private readonly RtfToMarkdownOptions _options;

    internal RtfToMarkdownConversionContext(RtfToMarkdownOptions options) {
        _options = options ?? throw new ArgumentNullException(nameof(options));
    }

    internal MarkdownReaderOptions? InlineReaderOptions => _options.InlineReaderOptions;
    internal bool IncludeHiddenText => _options.IncludeHiddenText;
    internal bool EmitUnsupportedHtmlComments => _options.EmitUnsupportedHtmlComments;
    internal Func<RtfImage, int, string>? ImagePathFactory => _options.ImagePathFactory;
    internal Action<RtfImage, int, string>? ImageExporter => _options.ImageExporter;
    internal RtfMarkdownNoteRegistry? NoteRegistry { get; set; }

    internal void Report(
        string code,
        RtfMarkdownDiagnosticSeverity severity,
        string message,
        string? source = null,
        RtfConversionAction action = RtfConversionAction.Preserved) =>
        AddDiagnostic(code, severity, message, source, action);
}

internal sealed class MarkdownToRtfConversionContext : RtfMarkdownConversionContext {
    private readonly MarkdownToRtfOptions _options;

    internal MarkdownToRtfConversionContext(MarkdownToRtfOptions options) {
        _options = options ?? throw new ArgumentNullException(nameof(options));
        if (_options.MaxListNestingDepth <= 0) {
            throw new ArgumentOutOfRangeException(nameof(options.MaxListNestingDepth), _options.MaxListNestingDepth, "Maximum list nesting depth must be positive.");
        }

        if (_options.MaxTableCells <= 0) {
            throw new ArgumentOutOfRangeException(nameof(options.MaxTableCells), _options.MaxTableCells, "Maximum table cells must be positive.");
        }
    }

    internal MarkdownReaderOptions? ReaderOptions => _options.ReaderOptions;
    internal bool PreserveRawHtmlAsText => _options.PreserveRawHtmlAsText;
    internal int MaxListNestingDepth => _options.MaxListNestingDepth;
    internal int MaxTableCells => _options.MaxTableCells;

    internal void Report(
        string code,
        RtfMarkdownDiagnosticSeverity severity,
        string message,
        string? source = null,
        RtfConversionAction action = RtfConversionAction.Preserved) =>
        AddDiagnostic(code, severity, message, source, action);
}

internal sealed class RtfMarkdownNoteRegistry {
    private readonly Dictionary<RtfNote, string> _labels = new Dictionary<RtfNote, string>();
    private readonly List<KeyValuePair<string, RtfNote>> _ordered = new List<KeyValuePair<string, RtfNote>>();
    private int _footnoteIndex;
    private int _endnoteIndex;

    internal IReadOnlyList<KeyValuePair<string, RtfNote>> Ordered => _ordered;

    internal string? Register(RtfNote note, RtfToMarkdownConversionContext context) {
        if (note.Kind == RtfNoteKind.Annotation) {
            context.Report(
                "RTFMD012",
                RtfMarkdownDiagnosticSeverity.Warning,
                "RTF annotation omitted from Markdown output.",
                note.Id ?? note.Kind.ToString(),
                RtfConversionAction.Omitted);
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
