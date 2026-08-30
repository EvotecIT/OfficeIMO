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
        bool cslJsonSingleObjectRoot = false;
        bool endNoteRecordsRoot = false;
        IList<BibliographyItem> items;
        try {
            switch (format) {
                case BibliographyFormat.BibTex:
                case BibliographyFormat.BibLatex:
                    items = BibCodec.Parse(source, format, options, diagnostics, nativeEntries, cancellationToken);
                    break;
                case BibliographyFormat.CslJson:
                    items = CslJsonCodec.Parse(source, options, diagnostics, out cslJsonSingleObjectRoot, cancellationToken);
                    break;
                case BibliographyFormat.Ris:
                    items = TaggedCodec.ParseRis(source, options, diagnostics, cancellationToken);
                    break;
                case BibliographyFormat.Nbib:
                    items = TaggedCodec.ParseNbib(source, options, diagnostics, cancellationToken);
                    break;
                case BibliographyFormat.EndNoteXml:
                    items = EndNoteXmlCodec.Parse(source, options, diagnostics, nativeEntries, out endNoteRecordsRoot, cancellationToken);
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

        var document = new BibliographyDocument(format, items, nativeEntries, source, originalBytes, diagnostics.AsReadOnly(), cslJsonSingleObjectRoot, endNoteRecordsRoot);
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
        AddValue(partial, value?.Length ?? 0, offset);
    }

    internal void AddValue(IList<BibliographyItem> partial, int valueLength, int offset) {
        _values++;
        if (_values > _options.MaximumValueCount) throw new BibliographyLimitException("Maximum bibliography value count was exceeded.", partial, offset);
        if (valueLength > _options.MaximumValueLength) throw new BibliographyLimitException("Maximum bibliography value length was exceeded.", partial, offset);
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

    internal static BibliographyFormat Detect(string source, BibliographyReadOptions? options = null, CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        options ??= new BibliographyReadOptions();
        options.Validate();
        if (source.Length > options.MaximumInputCharacters) throw new InvalidDataException($"Bibliography input exceeds the configured {options.MaximumInputCharacters} character limit.");
        cancellationToken.ThrowIfCancellationRequested();
        int start = SkipWhitespace(source, 0, cancellationToken);
        if (start < source.Length && source[start] == '%') {
            int bib = SkipLeadingComments(source, start, true, cancellationToken);
            if (StartsWith(source, bib, "@", StringComparison.Ordinal)) return BibliographyFormat.BibLatex;
            throw new FormatException("Bibliography format could not be detected. Pass an explicit BibliographyFormat.");
        }
        if (start + 1 < source.Length && source[start] == '/' && (source[start + 1] == '/' || source[start + 1] == '*')) {
            int json = SkipLeadingComments(source, start, false, cancellationToken);
            if (LooksLikeCsl(source, json)) return BibliographyFormat.CslJson;
            throw new FormatException("Bibliography format could not be detected. Pass an explicit BibliographyFormat.");
        }
        if (LooksLikeCsl(source, start)) return BibliographyFormat.CslJson;
        if (start < source.Length && source[start] == '<') {
            int xml = SkipLeadingXmlTrivia(source, start, cancellationToken);
            if (LooksLikeEndNoteRoot(source, xml, cancellationToken)) return BibliographyFormat.EndNoteXml;
        }
        if (StartsWith(source, start, "@", StringComparison.Ordinal)) return BibliographyFormat.BibLatex;
        if (StartsWith(source, start, "TY  -", StringComparison.OrdinalIgnoreCase)) return BibliographyFormat.Ris;
        if (StartsWith(source, start, "PMID-", StringComparison.OrdinalIgnoreCase) || StartsWith(source, start, "PMID -", StringComparison.OrdinalIgnoreCase) || StartsWith(source, start, "OWN -", StringComparison.OrdinalIgnoreCase)) return BibliographyFormat.Nbib;
        throw new FormatException("Bibliography format could not be detected. Pass an explicit BibliographyFormat.");
    }

    private static int SkipLeadingComments(string source, int position, bool bibComments, CancellationToken cancellationToken) {
        while (position < source.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            position = SkipWhitespace(source, position, cancellationToken);
            if (bibComments && position < source.Length && source[position] == '%') {
                int end = FindCharacter(source, position + 1, '\n', cancellationToken);
                position = end < 0 ? source.Length : end + 1;
                continue;
            }
            if (!bibComments && position + 1 < source.Length && source[position] == '/' && source[position + 1] == '/') {
                int end = FindCharacter(source, position + 2, '\n', cancellationToken);
                position = end < 0 ? source.Length : end + 1;
                continue;
            }
            if (!bibComments && position + 1 < source.Length && source[position] == '/' && source[position + 1] == '*') {
                int end = FindSequence(source, position + 2, "*/", cancellationToken);
                if (end < 0) return position;
                position = end + 2;
                continue;
            }
            break;
        }
        return position;
    }

    private static int SkipLeadingXmlTrivia(string source, int position, CancellationToken cancellationToken) {
        while (position < source.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            position = SkipWhitespace(source, position, cancellationToken);
            if (StartsWith(source, position, "<!--", StringComparison.Ordinal)) {
                int end = FindSequence(source, position + 4, "-->", cancellationToken);
                if (end < 0) return position;
                position = end + 3;
                continue;
            }
            if (StartsWith(source, position, "<?", StringComparison.Ordinal)) {
                int end = FindSequence(source, position + 2, "?>", cancellationToken);
                if (end < 0) return position;
                position = end + 2;
                continue;
            }
            break;
        }
        return position;
    }

    private static bool LooksLikeCsl(string source, int position) =>
        StartsWith(source, position, "[", StringComparison.Ordinal) || StartsWith(source, position, "{", StringComparison.Ordinal);

    private static bool LooksLikeEndNoteRoot(string source, int position, CancellationToken cancellationToken) {
        if (position < 0 || position >= source.Length || source[position] != '<' || position + 1 >= source.Length) return false;
        int nameStart = position + 1;
        if (source[nameStart] == '/' || source[nameStart] == '!' || source[nameStart] == '?') return false;
        int nameEnd = nameStart;
        while (nameEnd < source.Length && !char.IsWhiteSpace(source[nameEnd]) && source[nameEnd] != '/' && source[nameEnd] != '>') {
            if (((nameEnd - nameStart) & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            nameEnd++;
        }
        if (nameEnd == nameStart) return false;
        int localStart = nameStart;
        for (int index = nameStart; index < nameEnd; index++) if (source[index] == ':') localStart = index + 1;
        int localLength = nameEnd - localStart;
        return localLength == 3 && string.Compare(source, localStart, "xml", 0, 3, StringComparison.OrdinalIgnoreCase) == 0 ||
            localLength == 7 && string.Compare(source, localStart, "records", 0, 7, StringComparison.OrdinalIgnoreCase) == 0;
    }

    private static bool StartsWith(string source, int position, string value, StringComparison comparison) =>
        position >= 0 && position <= source.Length - value.Length && string.Compare(source, position, value, 0, value.Length, comparison) == 0;

    private static int SkipWhitespace(string source, int position, CancellationToken cancellationToken) {
        int start = position;
        while (position < source.Length && (char.IsWhiteSpace(source[position]) || source[position] == '\uFEFF')) {
            if (((position - start) & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            position++;
        }
        return position;
    }

    private static int FindCharacter(string source, int position, char value, CancellationToken cancellationToken) {
        int start = position;
        while (position < source.Length) {
            if (((position - start) & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (source[position] == value) return position;
            position++;
        }
        return -1;
    }

    private static int FindSequence(string source, int position, string value, CancellationToken cancellationToken) {
        int start = position;
        while (position <= source.Length - value.Length) {
            if (((position - start) & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (StartsWith(source, position, value, StringComparison.Ordinal)) return position;
            position++;
        }
        return -1;
    }
}
