namespace OfficeIMO.Bibliography;

internal static class BibliographyReader {
    internal static BibliographyReadResult Parse(string source, BibliographyFormat format, BibliographyReadOptions? options, byte[]? originalBytes, CancellationToken cancellationToken) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        options ??= new BibliographyReadOptions();
        options.Validate();
        if (source.Length > options.MaximumInputCharacters) throw new InvalidDataException($"Bibliography input exceeds the configured {options.MaximumInputCharacters} character limit.");
        cancellationToken.ThrowIfCancellationRequested();

        var diagnostics = new List<BibliographyDiagnostic>();
        var nativeEntries = new List<BibliographyNativeEntry>();
        IList<BibliographyItem> items;
        try {
            switch (format) {
                case BibliographyFormat.BibTex:
                case BibliographyFormat.BibLatex:
                    items = BibCodec.Parse(source, format, options, diagnostics, nativeEntries, cancellationToken);
                    break;
                case BibliographyFormat.CslJson:
                    items = CslJsonCodec.Parse(source, options, diagnostics, cancellationToken);
                    break;
                case BibliographyFormat.Ris:
                    items = TaggedCodec.ParseRis(source, options, diagnostics, cancellationToken);
                    break;
                case BibliographyFormat.Nbib:
                    items = TaggedCodec.ParseNbib(source, options, diagnostics, cancellationToken);
                    break;
                case BibliographyFormat.EndNoteXml:
                    items = EndNoteXmlCodec.Parse(source, options, diagnostics, nativeEntries, cancellationToken);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(format), format, "Unknown bibliography format.");
            }
        } catch (BibliographyLimitException exception) {
            diagnostics.Add(new BibliographyDiagnostic("BIBLIM001", BibliographyDiagnosticSeverity.Error, exception.Message, exception.Offset));
            items = exception.PartialItems;
        } catch (BibliographyDiagnosticLimitException exception) {
            diagnostics.Add(new BibliographyDiagnostic("BIBLIM002", BibliographyDiagnosticSeverity.Error, "Maximum bibliography diagnostic count was exceeded.", exception.Offset, exception.Line, exception.Column));
            items = exception.PartialItems;
        }

        var document = new BibliographyDocument(format, items, nativeEntries, source, originalBytes, diagnostics.AsReadOnly());
        return new BibliographyReadResult(document, document.Diagnostics);
    }
}

internal sealed class BibliographyLimitGuard {
    private readonly BibliographyReadOptions _options;
    private int _items;
    private int _values;

    internal BibliographyLimitGuard(BibliographyReadOptions options) => _options = options;

    internal void AddItem(IList<BibliographyItem> partial, int offset) {
        _items++;
        if (_items > _options.MaximumItemCount) throw new BibliographyLimitException("Maximum bibliography item count was exceeded.", partial, offset);
    }

    internal void AddValue(IList<BibliographyItem> partial, string? value, int offset) {
        _values++;
        if (_values > _options.MaximumValueCount) throw new BibliographyLimitException("Maximum bibliography value count was exceeded.", partial, offset);
        if (value != null && value.Length > _options.MaximumValueLength) throw new BibliographyLimitException("Maximum bibliography value length was exceeded.", partial, offset);
    }

    internal void CheckDepth(IList<BibliographyItem> partial, int depth, int offset) {
        if (depth > _options.MaximumNestingDepth) throw new BibliographyLimitException("Maximum bibliography nesting depth was exceeded.", partial, offset);
    }

    internal void CheckValueLength(IList<BibliographyItem> partial, string value, int offset) {
        CheckValueLength(partial, value.Length, offset);
    }

    internal void CheckValueLength(IList<BibliographyItem> partial, int length, int offset) {
        if (length > _options.MaximumValueLength) throw new BibliographyLimitException("Maximum bibliography value length was exceeded.", partial, offset);
    }

    internal void CheckAdditionalValueLength(IList<BibliographyItem> partial, int currentLength, int additionalLength, int offset) {
        if (currentLength < 0 || additionalLength < 0 || currentLength > _options.MaximumValueLength || additionalLength > _options.MaximumValueLength - currentLength)
            throw new BibliographyLimitException("Maximum bibliography value length was exceeded.", partial, offset);
    }
}

internal sealed class BibliographyDiagnosticGuard {
    private readonly List<BibliographyDiagnostic> _diagnostics;
    private readonly int _maximumCount;
    private readonly IList<BibliographyItem> _partialItems;

    internal BibliographyDiagnosticGuard(BibliographyReadOptions options, List<BibliographyDiagnostic> diagnostics, IList<BibliographyItem> partialItems) {
        _maximumCount = options.MaximumDiagnosticCount;
        _diagnostics = diagnostics;
        _partialItems = partialItems;
    }

    internal void Add(BibliographyDiagnostic diagnostic) {
        if (_diagnostics.Count >= _maximumCount) throw new BibliographyDiagnosticLimitException(_partialItems, diagnostic.Offset, diagnostic.Line, diagnostic.Column);
        _diagnostics.Add(diagnostic);
    }
}

internal sealed class BibliographyDiagnosticLimitException : Exception {
    internal BibliographyDiagnosticLimitException(IList<BibliographyItem> partialItems, int offset, int line, int column)
        : base("Maximum bibliography diagnostic count was exceeded.") {
        PartialItems = partialItems;
        Offset = offset;
        Line = line;
        Column = column;
    }

    internal IList<BibliographyItem> PartialItems { get; }
    internal int Offset { get; }
    internal int Line { get; }
    internal int Column { get; }
}

internal sealed class BibliographyLimitException : Exception {
    internal BibliographyLimitException(string message, IList<BibliographyItem> partialItems, int offset) : base(message) {
        PartialItems = partialItems;
        Offset = offset;
    }

    internal IList<BibliographyItem> PartialItems { get; }
    internal int Offset { get; }
}

internal static class BibliographyFormatDetector {
    internal static bool TryDetectPath(string path, out BibliographyFormat format) {
        string extension = Path.GetExtension(path).ToLowerInvariant();
        switch (extension) {
            case ".bib": format = BibliographyFormat.BibLatex; return true;
            case ".json": format = BibliographyFormat.CslJson; return true;
            case ".ris": format = BibliographyFormat.Ris; return true;
            case ".nbib": case ".medline": format = BibliographyFormat.Nbib; return true;
            case ".xml": format = BibliographyFormat.EndNoteXml; return true;
            default: format = default; return false;
        }
    }

    internal static BibliographyFormat Detect(string source) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        int start = SkipWhitespace(source, 0);
        if (start < source.Length && source[start] == '%') {
            string bib = TrimLeadingComments(source, start, true);
            if (bib.StartsWith("@", StringComparison.Ordinal)) return BibliographyFormat.BibLatex;
            throw new FormatException("Bibliography format could not be detected. Pass an explicit BibliographyFormat.");
        }
        if (start + 1 < source.Length && source[start] == '/' && (source[start + 1] == '/' || source[start + 1] == '*')) {
            string json = TrimLeadingComments(source, start, false);
            if (json.StartsWith("[", StringComparison.Ordinal) || json.StartsWith("{", StringComparison.Ordinal) && json.IndexOf("\"type\"", StringComparison.OrdinalIgnoreCase) >= 0) return BibliographyFormat.CslJson;
            throw new FormatException("Bibliography format could not be detected. Pass an explicit BibliographyFormat.");
        }
        string value = source.Substring(start);
        if (value.StartsWith("[", StringComparison.Ordinal) || value.StartsWith("{", StringComparison.Ordinal) && value.IndexOf("\"type\"", StringComparison.OrdinalIgnoreCase) >= 0) return BibliographyFormat.CslJson;
        if (value.StartsWith("<?xml", StringComparison.OrdinalIgnoreCase) || value.StartsWith("<xml", StringComparison.OrdinalIgnoreCase) || value.StartsWith("<records", StringComparison.OrdinalIgnoreCase)) return BibliographyFormat.EndNoteXml;
        if (value.StartsWith("@", StringComparison.Ordinal)) return BibliographyFormat.BibLatex;
        if (value.StartsWith("TY  -", StringComparison.Ordinal)) return BibliographyFormat.Ris;
        if (value.StartsWith("PMID-", StringComparison.Ordinal) || value.StartsWith("PMID -", StringComparison.Ordinal) || value.StartsWith("OWN -", StringComparison.Ordinal)) return BibliographyFormat.Nbib;
        throw new FormatException("Bibliography format could not be detected. Pass an explicit BibliographyFormat.");
    }

    private static string TrimLeadingComments(string source, int position, bool bibComments) {
        while (position < source.Length) {
            position = SkipWhitespace(source, position);
            if (bibComments && position < source.Length && source[position] == '%') {
                int end = source.IndexOf('\n', position + 1);
                position = end < 0 ? source.Length : end + 1;
                continue;
            }
            if (!bibComments && position + 1 < source.Length && source[position] == '/' && source[position + 1] == '/') {
                int end = source.IndexOf('\n', position + 2);
                position = end < 0 ? source.Length : end + 1;
                continue;
            }
            if (!bibComments && position + 1 < source.Length && source[position] == '/' && source[position + 1] == '*') {
                int end = source.IndexOf("*/", position + 2, StringComparison.Ordinal);
                if (end < 0) return source.Substring(position);
                position = end + 2;
                continue;
            }
            break;
        }
        return source.Substring(position);
    }

    private static int SkipWhitespace(string source, int position) {
        while (position < source.Length && char.IsWhiteSpace(source[position])) position++;
        return position;
    }
}
