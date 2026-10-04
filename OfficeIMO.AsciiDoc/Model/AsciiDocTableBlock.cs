namespace OfficeIMO.AsciiDoc;

/// <summary>Supported AsciiDoc table data formats.</summary>
public enum AsciiDocTableFormat {
    /// <summary>Prefix-separated values with per-cell specifiers.</summary>
    Psv = 0,
    /// <summary>Comma-separated values with quoting.</summary>
    Csv,
    /// <summary>Tab-separated values with quoting.</summary>
    Tsv,
    /// <summary>Backslash-escaped delimiter-separated values.</summary>
    Dsv
}

/// <summary>Typed source-backed AsciiDoc table.</summary>
public sealed class AsciiDocTableBlock : AsciiDocDelimitedBlock {
    private AsciiDocTable _table;
    private AsciiDocParseOptions _tableOptions = new AsciiDocParseOptions();
    internal AsciiDocTableBlock(
        AsciiDocSyntaxNode syntax,
        string delimiter,
        string openingText,
        string content,
        string closingText,
        bool isTerminated,
        string trailingLineEnding,
        AsciiDocTable table)
        : base(syntax, AsciiDocDelimitedBlockKind.Table, delimiter, openingText, content, closingText, isTerminated, trailingLineEnding) {
        _table = table;
    }

    /// <summary>Rows, cells, spans, format, and separator semantics.</summary>
    public AsciiDocTable Table => _table;

    /// <summary>Current table source. Assignment reparses the typed cells before changing the block.</summary>
    public override string Content {
        get => Table.IsModified ? Table.Write(new AsciiDocWriterContext(AsciiDocWriterMode.Preserve, "\n")) : base.Content;
        set {
            string replacement = value ?? string.Empty;
            if (string.Equals(Content, replacement, StringComparison.Ordinal)) return;
            AsciiDocParser.ValidateOptions(replacement, _tableOptions);
            var factory = new AsciiDocSyntaxFactory(new AsciiDocSourceText(replacement), default, _tableOptions);
            var configuration = AsciiDocTableConfiguration.Create(Delimiter, AttributeLists, _tableOptions.MaximumTableColumnCount);
            AsciiDocTable parsed = AsciiDocTableParser.Parse(factory, 0, replacement.Length, configuration).Table;
            foreach (AsciiDocTableCell cell in parsed.Cells) cell.SetBodyOptions(_tableOptions);
            base.Content = replacement;
            SetValue(ref _table, parsed);
        }
    }

    internal override void SetBodyOptions(AsciiDocParseOptions options) {
        _tableOptions = options.Copy();
        base.SetBodyOptions(options);
        foreach (AsciiDocTableCell cell in Table.Cells) cell.SetBodyOptions(options);
    }

    /// <inheritdoc />
    public override bool IsModified => base.IsModified || Table.IsModified;

    internal override string WriteCore(AsciiDocWriterContext context) {
        if (IsContentAssigned) return base.WriteCore(context);
        string content = Table.Write(context);
        if (context.Mode == AsciiDocWriterMode.Preserve) {
            if (IsTerminated && content.Length > 0 && !AsciiDocText.EndsWithLineEnding(content)) content += context.LineEnding;
            return OpeningText + content + ClosingText;
        }
        var output = new StringBuilder();
        output.Append(Delimiter).Append(context.LineEnding).Append(AsciiDocText.NormalizeLineEndings(content, context.LineEnding));
        if (IsTerminated) {
            if (content.Length > 0 && !AsciiDocText.EndsWithLineEnding(content)) output.Append(context.LineEnding);
            output.Append(Delimiter).Append(EffectiveTrailingLineEnding(context));
        }
        return output.ToString();
    }
}

/// <summary>Semantic table model that retains exact cell source.</summary>
public sealed class AsciiDocTable {
    private readonly IReadOnlyList<AsciiDocTableCell> _cells;
    private readonly IReadOnlyList<AsciiDocTableRow> _rows;

    internal AsciiDocTable(
        AsciiDocSyntaxNode syntax,
        AsciiDocTableFormat format,
        string separator,
        int columnCount,
        string prefix,
        string suffix,
        IReadOnlyList<AsciiDocTableCell> cells,
        IReadOnlyList<AsciiDocTableRow> rows) {
        Syntax = syntax;
        Format = format;
        Separator = separator;
        ColumnCount = columnCount;
        Prefix = prefix;
        Suffix = suffix;
        _cells = cells;
        _rows = rows;
    }

    /// <summary>Lossless table-content syntax.</summary>
    public AsciiDocSyntaxNode Syntax { get; }

    /// <summary>Table data format.</summary>
    public AsciiDocTableFormat Format { get; }

    /// <summary>Cell separator.</summary>
    public string Separator { get; }

    /// <summary>Effective number of columns.</summary>
    public int ColumnCount { get; }

    /// <summary>Cells in logical source order.</summary>
    public IReadOnlyList<AsciiDocTableCell> Cells => _cells;

    /// <summary>Rows grouped using explicit or inferred column count.</summary>
    public IReadOnlyList<AsciiDocTableRow> Rows => _rows;

    /// <summary>True when a cell changed.</summary>
    public bool IsModified => Cells.Any(static cell => cell.IsModified);

    internal string Prefix { get; }
    internal string Suffix { get; }

    internal string Write(AsciiDocWriterContext context) {
        if (context.Mode == AsciiDocWriterMode.Preserve && !IsModified) return Syntax.OriginalText;
        var output = new StringBuilder(Syntax.OriginalText.Length);
        output.Append(context.Mode == AsciiDocWriterMode.Canonical ? AsciiDocText.NormalizeLineEndings(Prefix, context.LineEnding) : Prefix);
        for (int index = 0; index < Cells.Count; index++) output.Append(Cells[index].Write(context));
        output.Append(context.Mode == AsciiDocWriterMode.Canonical ? AsciiDocText.NormalizeLineEndings(Suffix, context.LineEnding) : Suffix);
        return output.ToString();
    }
}

/// <summary>Logical table row.</summary>
public sealed class AsciiDocTableRow {
    internal AsciiDocTableRow(int index, bool isHeader, IReadOnlyList<AsciiDocTableCell> cells) {
        Index = index;
        IsHeader = isHeader;
        Cells = cells;
    }

    /// <summary>Zero-based row index.</summary>
    public int Index { get; }

    /// <summary>True when header semantics apply.</summary>
    public bool IsHeader { get; }

    /// <summary>Cells that begin in this row.</summary>
    public IReadOnlyList<AsciiDocTableCell> Cells { get; }
}

/// <summary>Source-backed table cell.</summary>
public sealed class AsciiDocTableCell {
    private string _content;
    private bool _isModified;
    private AsciiDocDocument? _body;
    private AsciiDocInlineSequence? _inlines;
    private AsciiDocParseOptions _bodyOptions = new AsciiDocParseOptions();

    internal AsciiDocTableCell(
        AsciiDocSyntaxNode syntax,
        AsciiDocTableFormat format,
        string separator,
        string leadingText,
        string specifier,
        string content,
        int rowIndex,
        int columnIndex,
        char columnStyle = 'd') {
        Syntax = syntax;
        Format = format;
        Separator = separator;
        LeadingText = leadingText;
        Specifier = specifier;
        _content = content;
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
        ParseSpan(specifier, out int columnSpan, out int rowSpan);
        ColumnSpan = columnSpan;
        RowSpan = rowSpan;
        Style = ParseStyle(specifier, columnStyle);
    }

    /// <summary>Lossless cell syntax, including its leading separator or row boundary.</summary>
    public AsciiDocSyntaxNode Syntax { get; }

    /// <summary>Zero-based logical row.</summary>
    public int RowIndex { get; internal set; }

    /// <summary>Zero-based logical column.</summary>
    public int ColumnIndex { get; internal set; }

    /// <summary>Column span from a PSV specifier.</summary>
    public int ColumnSpan { get; }

    /// <summary>Row span from a PSV specifier.</summary>
    public int RowSpan { get; }

    /// <summary>Raw PSV cell specifier, excluding the separator.</summary>
    public string Specifier { get; }

    /// <summary>Cell style operator, or <c>d</c> for default.</summary>
    public char Style { get; }

    /// <summary>Exact raw content after the leading separator and specifier.</summary>
    public string Content {
        get => _body?.IsModified == true ? _body.ToAsciiDoc() : _inlines?.IsModified == true ? ReplaceValue(_content, _inlines.ToAsciiDoc()) : _content;
        set {
            string normalized = value ?? string.Empty;
            if (string.Equals(Content, normalized, StringComparison.Ordinal)) return;
            _content = normalized;
            _body = null;
            _inlines = null;
            _isModified = true;
        }
    }

    /// <summary>Decoded and trimmed value for data-table formats; trimmed content for PSV.</summary>
    public string Value {
        get => Decode(Content, Format);
        set {
            Content = ReplaceValue(Content, value ?? string.Empty);
        }
    }

    /// <summary>True when content changed.</summary>
    public bool IsModified => _isModified || _body?.IsModified == true || _inlines?.IsModified == true;

    /// <summary>Typed editable content for default, emphasis, header, or strong cells; null for block and verbatim styles.</summary>
    public AsciiDocInlineSequence? Inlines => GetInlines();
    /// <summary>Parses inline cell content with the owning document's limits and cancellation.</summary>
    public AsciiDocInlineSequence? GetInlines(System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (Style == 'a' || Style == 'l' || Style == 'm') return null;
        return _inlines ??= AsciiDocInlineSequence.Parse(Value, _bodyOptions, cancellationToken);
    }

    /// <summary>Typed editable child blocks for an AsciiDoc-style PSV cell; null for other cell styles.</summary>
    public AsciiDocDocument? Body => GetBody();
    /// <summary>Parses the cell body with the owning document's limits and cooperative cancellation.</summary>
    public AsciiDocDocument? GetBody(System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (Style != 'a' || Format != AsciiDocTableFormat.Psv) return null;
        return _body ??= AsciiDocDocument.Parse(_content, _bodyOptions, cancellationToken);
    }
    internal void SetBodyOptions(AsciiDocParseOptions options) => _bodyOptions = options.Copy();

    internal AsciiDocTableFormat Format { get; }
    internal string Separator { get; }
    internal string LeadingText { get; }

    private string ReplaceValue(string content, string value) {
        int leading = 0;
        while (leading < content.Length && char.IsWhiteSpace(content[leading])) leading++;
        int trailing = content.Length;
        while (trailing > leading && char.IsWhiteSpace(content[trailing - 1])) trailing--;
        return content.Substring(0, leading) + Encode(value, Format, Separator) + content.Substring(trailing);
    }

    internal string Write(AsciiDocWriterContext context) {
        if (context.Mode == AsciiDocWriterMode.Preserve && !IsModified) return Syntax.OriginalText;
        string leading = context.Mode == AsciiDocWriterMode.Canonical
            ? AsciiDocText.NormalizeLineEndings(LeadingText, context.LineEnding)
            : LeadingText;
        string content = context.Mode == AsciiDocWriterMode.Canonical
            ? AsciiDocText.NormalizeLineEndings(Content, context.LineEnding)
            : Content;
        return leading + content;
    }

    private static void ParseSpan(string specifier, out int columnSpan, out int rowSpan) {
        columnSpan = 1;
        rowSpan = 1;
        int plus = specifier.IndexOf('+');
        if (plus <= 0) return;
        string span = specifier.Substring(0, plus);
        int dot = span.IndexOf('.');
        if (dot == 0) {
            if (int.TryParse(span.Substring(1), out int rows) && rows > 0) rowSpan = rows;
        } else if (dot > 0) {
            if (int.TryParse(span.Substring(0, dot), out int columns) && columns > 0) columnSpan = columns;
            if (int.TryParse(span.Substring(dot + 1), out int rows) && rows > 0) rowSpan = rows;
        } else if (int.TryParse(span, out int columns) && columns > 0) {
            columnSpan = columns;
        }
    }

    private static char ParseStyle(string specifier, char fallback) {
        for (int index = specifier.Length - 1; index >= 0; index--) {
            char value = specifier[index];
            if (value == 'a' || value == 'd' || value == 'e' || value == 'h' || value == 'l' || value == 'm' || value == 's') return value;
        }
        return fallback;
    }

    private static string Decode(string value, AsciiDocTableFormat format) {
        string trimmed = value.Trim();
        if ((format == AsciiDocTableFormat.Csv || format == AsciiDocTableFormat.Tsv) &&
            trimmed.Length >= 2 && trimmed[0] == '"' && trimmed[trimmed.Length - 1] == '"') {
            return trimmed.Substring(1, trimmed.Length - 2).Replace("\"\"", "\"");
        }
        if (format == AsciiDocTableFormat.Dsv) return trimmed.Replace("\\:", ":").Replace("\\\\", "\\");
        return trimmed;
    }

    private static string Encode(string value, AsciiDocTableFormat format, string separator) {
        if (format == AsciiDocTableFormat.Csv || format == AsciiDocTableFormat.Tsv) {
            if (value.IndexOf(separator, StringComparison.Ordinal) >= 0 || value.IndexOf('"') >= 0 || value.IndexOf('\r') >= 0 || value.IndexOf('\n') >= 0) {
                return "\"" + value.Replace("\"", "\"\"") + "\"";
            }
        } else if (format == AsciiDocTableFormat.Dsv) {
            return value.Replace("\\", "\\\\").Replace(separator, "\\" + separator);
        }
        return value;
    }
}
