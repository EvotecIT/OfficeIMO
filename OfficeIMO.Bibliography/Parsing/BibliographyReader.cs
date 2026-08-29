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
        if (value.Length > _options.MaximumValueLength) throw new BibliographyLimitException("Maximum bibliography value length was exceeded.", partial, offset);
    }
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
        string value = source.TrimStart();
        if (value.StartsWith("[", StringComparison.Ordinal) || value.StartsWith("{", StringComparison.Ordinal) && value.IndexOf("\"type\"", StringComparison.OrdinalIgnoreCase) >= 0) return BibliographyFormat.CslJson;
        if (value.StartsWith("<?xml", StringComparison.OrdinalIgnoreCase) || value.StartsWith("<xml", StringComparison.OrdinalIgnoreCase) || value.StartsWith("<records", StringComparison.OrdinalIgnoreCase)) return BibliographyFormat.EndNoteXml;
        if (value.StartsWith("@", StringComparison.Ordinal)) return BibliographyFormat.BibLatex;
        if (value.StartsWith("TY  -", StringComparison.Ordinal)) return BibliographyFormat.Ris;
        if (value.StartsWith("PMID-", StringComparison.Ordinal) || value.StartsWith("PMID -", StringComparison.Ordinal) || value.StartsWith("OWN -", StringComparison.Ordinal)) return BibliographyFormat.Nbib;
        throw new FormatException("Bibliography format could not be detected. Pass an explicit BibliographyFormat.");
    }
}
