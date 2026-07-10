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
        Table = table;
    }

    /// <summary>Rows, cells, spans, format, and separator semantics.</summary>
    public AsciiDocTable Table { get; }

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

    internal AsciiDocTableCell(
        AsciiDocSyntaxNode syntax,
        AsciiDocTableFormat format,
        string separator,
        string leadingText,
        string specifier,
        string content,
        int rowIndex,
        int columnIndex) {
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
        Style = ParseStyle(specifier);
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
        get => _content;
        set {
            string normalized = value ?? string.Empty;
            if (string.Equals(_content, normalized, StringComparison.Ordinal)) return;
            _content = normalized;
            _isModified = true;
        }
    }

    /// <summary>Decoded and trimmed value for data-table formats; trimmed content for PSV.</summary>
    public string Value {
        get => Decode(Content, Format);
        set {
            string encoded = Encode(value ?? string.Empty, Format, Separator);
            int leading = 0;
            while (leading < Content.Length && char.IsWhiteSpace(Content[leading])) leading++;
            int trailing = Content.Length;
            while (trailing > leading && char.IsWhiteSpace(Content[trailing - 1])) trailing--;
            Content = Content.Substring(0, leading) + encoded + Content.Substring(trailing);
        }
    }

    /// <summary>True when content changed.</summary>
    public bool IsModified => _isModified;

    internal AsciiDocTableFormat Format { get; }
    internal string Separator { get; }
    internal string LeadingText { get; }

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

    private static char ParseStyle(string specifier) {
        for (int index = specifier.Length - 1; index >= 0; index--) {
            char value = specifier[index];
            if (value == 'a' || value == 'd' || value == 'e' || value == 'h' || value == 'l' || value == 'm' || value == 's') return value;
        }
        return 'd';
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
