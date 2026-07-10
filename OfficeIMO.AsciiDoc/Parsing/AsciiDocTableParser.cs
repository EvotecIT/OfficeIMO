namespace OfficeIMO.AsciiDoc;

internal sealed class AsciiDocTableConfiguration {
    internal AsciiDocTableConfiguration(AsciiDocTableFormat format, string separator, int? columnCount, bool header) {
        Format = format;
        Separator = separator;
        ColumnCount = columnCount;
        Header = header;
    }

    internal AsciiDocTableFormat Format { get; }
    internal string Separator { get; }
    internal int? ColumnCount { get; }
    internal bool Header { get; }

    internal static AsciiDocTableConfiguration Create(string delimiter, IReadOnlyList<AsciiDocBlockAttributeList> metadata) {
        string? formatValue = GetNamedValue(metadata, "format");
        AsciiDocTableFormat format = delimiter == ",==="
            ? AsciiDocTableFormat.Csv
            : delimiter == ":==="
                ? AsciiDocTableFormat.Dsv
                : ParseFormat(formatValue);
        string? separatorValue = GetNamedValue(metadata, "separator");
        string separator = DecodeSeparator(separatorValue) ?? DefaultSeparator(format);
        int? columns = ParseColumnCount(GetNamedValue(metadata, "cols"));
        bool header = HasOption(metadata, "header");
        return new AsciiDocTableConfiguration(format, separator, columns, header);
    }

    private static string? GetNamedValue(IReadOnlyList<AsciiDocBlockAttributeList> metadata, string name) {
        for (int index = metadata.Count - 1; index >= 0; index--) {
            string? value = metadata[index].Attributes.GetNamedValue(name);
            if (value != null) return value;
        }
        return null;
    }

    private static bool HasOption(IReadOnlyList<AsciiDocBlockAttributeList> metadata, string option) {
        for (int index = 0; index < metadata.Count; index++) {
            if (metadata[index].Attributes.Options.Any(value => string.Equals(value, option, StringComparison.OrdinalIgnoreCase))) return true;
        }
        return false;
    }

    private static AsciiDocTableFormat ParseFormat(string? value) {
        if (string.Equals(value, "csv", StringComparison.OrdinalIgnoreCase)) return AsciiDocTableFormat.Csv;
        if (string.Equals(value, "tsv", StringComparison.OrdinalIgnoreCase)) return AsciiDocTableFormat.Tsv;
        if (string.Equals(value, "dsv", StringComparison.OrdinalIgnoreCase)) return AsciiDocTableFormat.Dsv;
        return AsciiDocTableFormat.Psv;
    }

    private static string DefaultSeparator(AsciiDocTableFormat format) {
        switch (format) {
            case AsciiDocTableFormat.Csv: return ",";
            case AsciiDocTableFormat.Tsv: return "\t";
            case AsciiDocTableFormat.Dsv: return ":";
            default: return "|";
        }
    }

    private static string? DecodeSeparator(string? value) {
        if (value == null) return null;
        if (string.Equals(value, "\\t", StringComparison.Ordinal)) return "\t";
        return value.Length == 0 ? null : value;
    }

    private static int? ParseColumnCount(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return null;
        string trimmed = value!.Trim();
        if (trimmed.EndsWith("*", StringComparison.Ordinal) &&
            int.TryParse(trimmed.Substring(0, trimmed.Length - 1), out int multiplier) && multiplier > 0) return multiplier;
        string[] parts = trimmed.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
        return parts.Length == 0 ? null : parts.Length;
    }
}

internal sealed class AsciiDocTableParseResult {
    internal AsciiDocTableParseResult(AsciiDocSyntaxNode syntax, AsciiDocTable table) {
        Syntax = syntax;
        Table = table;
    }

    internal AsciiDocSyntaxNode Syntax { get; }
    internal AsciiDocTable Table { get; }
}

internal static class AsciiDocTableParser {
    internal static AsciiDocTableParseResult Parse(
        AsciiDocSyntaxFactory factory,
        int start,
        int end,
        AsciiDocTableConfiguration configuration) {
        List<CellDescriptor> descriptors = configuration.Format == AsciiDocTableFormat.Psv
            ? ParsePsv(factory.Source.Text, start, end, configuration.Separator)
            : ParseData(factory.Source.Text, start, end, configuration.Format, configuration.Separator);

        int columns = configuration.ColumnCount ?? InferColumnCount(factory.Source.Text, descriptors, configuration.Format);
        if (columns < 1) columns = Math.Max(1, descriptors.Count);
        var cells = new List<AsciiDocTableCell>(descriptors.Count);
        var syntaxNodes = new List<AsciiDocSyntaxNode>(descriptors.Count);
        int logicalRow = 0;
        int logicalColumn = 0;
        for (int index = 0; index < descriptors.Count; index++) {
            CellDescriptor descriptor = descriptors[index];
            if (configuration.Format != AsciiDocTableFormat.Psv) {
                logicalRow = descriptor.Row;
                logicalColumn = descriptor.Column;
            }
            AsciiDocSyntaxNode cellSyntax = CreateCellSyntax(factory, descriptor, configuration);
            string leading = factory.Source.Text.Substring(descriptor.Start, descriptor.ContentStart - descriptor.Start);
            string content = factory.Source.Text.Substring(descriptor.ContentStart, descriptor.End - descriptor.ContentStart);
            var cell = new AsciiDocTableCell(
                cellSyntax,
                configuration.Format,
                configuration.Separator,
                leading,
                descriptor.Specifier,
                content,
                logicalRow,
                logicalColumn);
            cells.Add(cell);
            syntaxNodes.Add(cellSyntax);
            if (configuration.Format == AsciiDocTableFormat.Psv) {
                logicalColumn += cell.ColumnSpan;
                if (logicalColumn >= columns) { logicalRow++; logicalColumn = 0; }
            }
        }

        bool autoHeader = configuration.Header || HasImplicitHeader(factory.Source.Text, cells, columns);
        IReadOnlyList<AsciiDocTableRow> rows = CreateRows(cells, autoHeader);
        AsciiDocSyntaxNode contentSyntax = factory.Node(AsciiDocSyntaxKind.BlockContent, start, end, syntaxNodes);
        string prefix = descriptors.Count == 0
            ? factory.Source.Text.Substring(start, end - start)
            : factory.Source.Text.Substring(start, descriptors[0].Start - start);
        string suffix = descriptors.Count == 0
            ? string.Empty
            : factory.Source.Text.Substring(descriptors[descriptors.Count - 1].End, end - descriptors[descriptors.Count - 1].End);
        var table = new AsciiDocTable(contentSyntax, configuration.Format, configuration.Separator, columns, prefix, suffix, cells, rows);
        return new AsciiDocTableParseResult(contentSyntax, table);
    }

    private static List<CellDescriptor> ParsePsv(string source, int start, int end, string separator) {
        var starts = new List<CellStart>();
        for (int index = start; index + separator.Length <= end; index++) {
            if (!Matches(source, separator, index, end) || IsEscaped(source, index, start)) continue;
            int specifierStart = FindSpecifierStart(source, index, start);
            bool atLineStart = specifierStart == FindLineStart(source, index, start);
            bool separated = index == start || char.IsWhiteSpace(source[index - 1]);
            if (!atLineStart && !separated) continue;
            if (specifierStart < index && !IsCellSpecifier(source, specifierStart, index)) {
                if (!separated) continue;
                specifierStart = index;
            }
            starts.Add(new CellStart(specifierStart, index, index + separator.Length));
            index += separator.Length - 1;
        }

        var descriptors = new List<CellDescriptor>(starts.Count);
        for (int index = 0; index < starts.Count; index++) {
            CellStart cell = starts[index];
            int cellEnd = index + 1 < starts.Count ? starts[index + 1].Start : end;
            string specifier = source.Substring(cell.Start, cell.SeparatorStart - cell.Start);
            descriptors.Add(new CellDescriptor(cell.Start, cell.ContentStart, cellEnd, specifier, -1, -1, cell.SeparatorStart));
        }
        return descriptors;
    }

    private static List<CellDescriptor> ParseData(
        string source,
        int start,
        int end,
        AsciiDocTableFormat format,
        string separator) {
        var descriptors = new List<CellDescriptor>();
        int leadingStart = start;
        int contentStart = start;
        int leadingSeparatorStart = -1;
        int row = 0;
        int column = 0;
        bool quoted = false;
        int index = start;
        while (index < end) {
            char current = source[index];
            if ((format == AsciiDocTableFormat.Csv || format == AsciiDocTableFormat.Tsv) && current == '"') {
                if (quoted && index + 1 < end && source[index + 1] == '"') { index += 2; continue; }
                quoted = !quoted;
                index++;
                continue;
            }
            if (format == AsciiDocTableFormat.Dsv && current == '\\') { index += Math.Min(2, end - index); continue; }
            if (!quoted && Matches(source, separator, index, end)) {
                descriptors.Add(new CellDescriptor(leadingStart, contentStart, index, string.Empty, row, column++, leadingSeparatorStart));
                leadingStart = index;
                leadingSeparatorStart = index;
                index += separator.Length;
                contentStart = index;
                continue;
            }
            if (!quoted && (current == '\r' || current == '\n')) {
                int lineEnd = index;
                int endingLength = current == '\r' && index + 1 < end && source[index + 1] == '\n' ? 2 : 1;
                bool hasRowContent = column > 0 || !string.IsNullOrWhiteSpace(source.Substring(contentStart, lineEnd - contentStart));
                if (hasRowContent) {
                    descriptors.Add(new CellDescriptor(leadingStart, contentStart, lineEnd, string.Empty, row, column, leadingSeparatorStart));
                    row++;
                }
                column = 0;
                leadingStart = index;
                leadingSeparatorStart = -1;
                index += endingLength;
                contentStart = index;
                continue;
            }
            index++;
        }
        if (column > 0 || contentStart < end) {
            descriptors.Add(new CellDescriptor(leadingStart, contentStart, end, string.Empty, row, column, leadingSeparatorStart));
        }
        return descriptors;
    }

    private static AsciiDocSyntaxNode CreateCellSyntax(
        AsciiDocSyntaxFactory factory,
        CellDescriptor descriptor,
        AsciiDocTableConfiguration configuration) {
        var children = new List<AsciiDocSyntaxNode>();
        if (descriptor.Specifier.Length > 0) {
            children.Add(factory.Node(AsciiDocSyntaxKind.TableCellSpecifier, descriptor.Start, descriptor.SeparatorStart));
        }
        if (descriptor.SeparatorStart >= descriptor.Start) {
            children.Add(factory.Node(
                AsciiDocSyntaxKind.TableCellSeparator,
                descriptor.SeparatorStart,
                descriptor.SeparatorStart + configuration.Separator.Length));
        }
        if (descriptor.ContentStart < descriptor.End) {
            children.Add(factory.Node(AsciiDocSyntaxKind.TableCellContent, descriptor.ContentStart, descriptor.End));
        }
        return factory.Node(AsciiDocSyntaxKind.TableCell, descriptor.Start, descriptor.End, children);
    }

    private static IReadOnlyList<AsciiDocTableRow> CreateRows(IReadOnlyList<AsciiDocTableCell> cells, bool header) {
        var groups = cells.GroupBy(static cell => cell.RowIndex).OrderBy(static group => group.Key);
        var rows = new List<AsciiDocTableRow>();
        foreach (IGrouping<int, AsciiDocTableCell> group in groups) {
            rows.Add(new AsciiDocTableRow(group.Key, header && group.Key == 0, group.ToArray()));
        }
        return rows;
    }

    private static int InferColumnCount(string source, IReadOnlyList<CellDescriptor> cells, AsciiDocTableFormat format) {
        if (cells.Count == 0) return 0;
        if (format != AsciiDocTableFormat.Psv) {
            return cells.GroupBy(static cell => cell.Row).Max(static group => group.Count());
        }
        int firstLineEnd = cells[0].SeparatorStart;
        while (firstLineEnd < source.Length && source[firstLineEnd] != '\r' && source[firstLineEnd] != '\n') firstLineEnd++;
        int count = 0;
        for (int index = 0; index < cells.Count; index++) {
            if (cells[index].SeparatorStart >= firstLineEnd) break;
            count++;
        }
        return count;
    }

    private static bool HasImplicitHeader(string source, IReadOnlyList<AsciiDocTableCell> cells, int columns) {
        if (cells.Count <= columns || columns < 1) return false;
        AsciiDocTableCell last = cells[Math.Min(columns, cells.Count) - 1];
        string original = last.Syntax.OriginalText;
        return original.EndsWith("\n\n", StringComparison.Ordinal) ||
               original.EndsWith("\r\r", StringComparison.Ordinal) ||
               original.EndsWith("\r\n\r\n", StringComparison.Ordinal);
    }

    private static int FindSpecifierStart(string source, int separator, int rangeStart) {
        int lineStart = FindLineStart(source, separator, rangeStart);
        int start = separator;
        while (start > lineStart && IsSpecifierCharacter(source[start - 1])) start--;
        return start;
    }

    private static int FindLineStart(string source, int index, int rangeStart) {
        int current = index;
        while (current > rangeStart && source[current - 1] != '\r' && source[current - 1] != '\n') current--;
        while (current < index && char.IsWhiteSpace(source[current]) && source[current] != '\r' && source[current] != '\n') current++;
        return current;
    }

    private static bool IsCellSpecifier(string source, int start, int end) {
        if (end <= start) return false;
        for (int index = start; index < end; index++) {
            if (!IsSpecifierCharacter(source[index])) return false;
        }
        return true;
    }

    private static bool IsSpecifierCharacter(char value) =>
        char.IsDigit(value) || value == '.' || value == '+' || value == '*' ||
        value == '<' || value == '>' || value == '^' ||
        value == 'a' || value == 'd' || value == 'e' || value == 'h' || value == 'l' || value == 'm' || value == 's' || value == 'v';

    private static bool Matches(string source, string token, int index, int end) {
        if (token.Length == 0 || index + token.Length > end) return false;
        for (int offset = 0; offset < token.Length; offset++) {
            if (source[index + offset] != token[offset]) return false;
        }
        return true;
    }

    private static bool IsEscaped(string source, int index, int rangeStart) {
        int slashes = 0;
        for (int current = index - 1; current >= rangeStart && source[current] == '\\'; current--) slashes++;
        return slashes % 2 != 0;
    }

    private readonly struct CellStart {
        internal CellStart(int start, int separatorStart, int contentStart) {
            Start = start;
            SeparatorStart = separatorStart;
            ContentStart = contentStart;
        }
        internal int Start { get; }
        internal int SeparatorStart { get; }
        internal int ContentStart { get; }
    }

    private readonly struct CellDescriptor {
        internal CellDescriptor(int start, int contentStart, int end, string specifier, int row, int column, int separatorStart) {
            Start = start;
            ContentStart = contentStart;
            End = end;
            Specifier = specifier;
            Row = row;
            Column = column;
            SeparatorStart = separatorStart;
        }
        internal int Start { get; }
        internal int ContentStart { get; }
        internal int End { get; }
        internal string Specifier { get; }
        internal int Row { get; }
        internal int Column { get; }
        internal int SeparatorStart { get; }
    }
}
