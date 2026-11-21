namespace OfficeIMO.Yaml;

/// <summary>
/// One lint of one text. Line rules read the text; structure rules read the loaded <see cref="YamlDocument"/>; the
/// comment rule and the scalar ranges that keep line fixes out of scalar content come from <see cref="YamlReader"/>
/// events, so no rule scans YAML syntax on its own.
/// </summary>
internal sealed class YamlLintRun {
    private static readonly string[] TruthyWords = [
        "YES", "Yes", "yes", "NO", "No", "no", "TRUE", "True", "true", "FALSE", "False", "false", "ON", "On", "on", "OFF", "Off", "off"
    ];

    private readonly string _text;
    private readonly YamlLintOptions _options;
    private readonly List<YamlLintFinding> _findings = new();
    private readonly int[] _lineStarts;
    /// <summary>Source ranges of every scalar the reader reported, in source order, with their styles.</summary>
    private readonly List<ScalarRange> _scalars = new();
    /// <summary>Offsets of each comment's <c>#</c> and the end of its text.</summary>
    private readonly List<(int Hash, int End)> _comments = new();
    private YamlDocument? _document;
    private bool _fixable;
    private int? _spaces;
    private bool? _indentedSequences;

    private readonly struct ScalarRange {
        public ScalarRange(int start, int end, YamlStyle style) {
            Start = start;
            End = end;
            Style = style;
        }

        public int Start { get; }
        public int End { get; }
        public YamlStyle Style { get; }
    }

    public YamlLintRun(string text, YamlLintOptions options) {
        _text = text;
        _options = options;
        _lineStarts = BuildLineStarts(text);
    }

    public IReadOnlyList<YamlLintFinding> Run() {
        ReadEvents();
        try {
            _document = YamlDocument.Parse(_text, _options.LoadOptions);
            _fixable = true;
        } catch (YamlException exception) {
            if (exception.Code == YamlErrorCodes.DuplicateKey && _options.KeyDuplicates.Enabled) {
                Add(_options.KeyDuplicates, Offset(exception.Position), FirstLine(exception.Message), fix: null);
            } else {
                _findings.Add(new YamlLintFinding(YamlLintRules.Syntax, OfficeDiagnosticSeverity.Error, Clamp(exception.Position), "Syntax error: " + FirstLine(exception.Message), null));
            }
        }
        try {
            if (_document is not null) CheckStructure(_document);
            CheckLines();
            CheckComments();
        } finally {
            _document?.Dispose();
        }
        // A stable sort: findings at one position keep rule order.
        return _findings.OrderBy(static f => f.Position.Offset).ToList();
    }

    // ------------------------------------------------------------------ events

    private void ReadEvents() {
        YamlLoadOptions? load = _options.LoadOptions;
        var readerOptions = new YamlReaderOptions { IncludeComments = true, Dialect = load?.Dialect ?? YamlDialect.Standard, Strict = load?.Strict ?? false };
        try {
            using var reader = new YamlReader(_text, readerOptions);
            while (reader.Read()) {
                if (reader.TokenType == YamlTokenType.Scalar) {
                    _scalars.Add(new ScalarRange((int)reader.Start.Offset, (int)reader.End.Offset, reader.Style));
                } else if (reader.TokenType == YamlTokenType.Comment) {
                    _comments.Add(((int)reader.Start.Offset, (int)reader.End.Offset));
                }
            }
        } catch (YamlException) {
            // The load reports the error; the ranges read so far still guard the lines before it.
        }
    }

    /// <summary>Whether <paramref name="offset"/> lies strictly inside a scalar, such as a line start in a multi-line one.</summary>
    private bool InsideScalar(int offset) {
        int low = 0, high = _scalars.Count - 1;
        while (low <= high) {
            int middle = (low + high) >> 1;
            ScalarRange range = _scalars[middle];
            if (offset <= range.Start) high = middle - 1;
            else if (offset >= range.End) low = middle + 1;
            else return true;
        }
        return false;
    }

    /// <summary>Whether a literal or folded scalar ends exactly at <paramref name="offset"/>, as one that runs to the end of the text does.</summary>
    private bool EndsBlockScalar(int offset) {
        int low = 0, high = _scalars.Count - 1, last = -1;
        while (low <= high) {
            int middle = (low + high) >> 1;
            if (_scalars[middle].Start < offset) {
                last = middle;
                low = middle + 1;
            } else {
                high = middle - 1;
            }
        }
        return last >= 0 && _scalars[last] is { Style: YamlStyle.Literal or YamlStyle.Folded } range && range.End == offset;
    }

    // ------------------------------------------------------------------ structure rules

    private void CheckStructure(YamlDocument document) {
        foreach (OfficeDiagnostic diagnostic in document.Diagnostics) {
            if (diagnostic.Code == OfficeDiagnosticCodes.YamlDuplicateKey) {
                Add(_options.KeyDuplicates, DiagnosticOffset(diagnostic), FirstLine(diagnostic.Message), fix: null);
            } else if (diagnostic.Code == OfficeDiagnosticCodes.YamlLenientIndentation) {
                Add(_options.Indentation, DiagnosticOffset(diagnostic), FirstLine(diagnostic.Message), fix: null);
            }
        }
        IReadOnlyList<YamlDocumentPart> parts = document.Documents;
        if (parts.Count > 0 && _options.DocumentStart.Enabled) CheckDocumentStart(parts[0]);
        foreach (YamlDocumentPart part in parts) {
            if (part.RootNode is { } root) Visit(root, key: false, flow: false);
        }
    }

    private void CheckDocumentStart(YamlDocumentPart part) {
        bool directives = part.Version is not null || part.TagDirectives.Count > 0;
        if (_options.DocumentStart.Present) {
            if (part.HasStartMarker || part.RootNode is null || part.RootNode._start < 0) return;
            string newLine = DominantNewLine();
            Add(_options.DocumentStart, part.RootNode._start, "Missing document start marker \"---\".", new YamlLintFix(0, 0, "---" + newLine));
            return;
        }
        if (!part.HasStartMarker || directives || part._markerEnd < 3) return;
        int marker = part._markerEnd - 3;
        int lineStart = YamlTrivia.LineStart(_text, marker);
        int contentEnd = YamlTrivia.LineContentEnd(_text, marker);
        bool bare = marker == lineStart && _text.Substring(part._markerEnd, contentEnd - part._markerEnd).Trim(' ', '\t').Length == 0;
        // Only a marker alone on its line is removed, so no comment, tag or node on that line is lost.
        YamlLintFix? fix = bare && _fixable ? new YamlLintFix(lineStart, YamlTrivia.AfterBreak(_text, contentEnd) - lineStart, string.Empty) : null;
        Add(_options.DocumentStart, marker, "Found a forbidden document start marker \"---\".", fix);
    }

    private void Visit(YamlNode node, bool key, bool flow) {
        switch (node) {
            case YamlScalar scalar:
                CheckScalar(scalar, key, flow);
                break;
            case YamlMapping mapping: {
                bool inFlow = flow || mapping._style == YamlStyle.Flow;
                string? previous = null;
                for (int i = 0; i < mapping._count; i++) {
                    YamlNode k = mapping.Entries[i].Key, v = mapping.Entries[i].Value;
                    Visit(k, key: true, inFlow);
                    if (!inFlow) CheckNestedIndentation(k, v);
                    if (_options.EmptyValues.Enabled && (inFlow ? _options.EmptyValues.ForbidInFlowMappings : _options.EmptyValues.ForbidInBlockMappings) && IsEmpty(v)) {
                        Add(_options.EmptyValues, v._start, "Empty value in " + (inFlow ? "flow" : "block") + " mapping.", fix: null);
                    }
                    if (_options.KeyOrdering.Enabled && k is YamlScalar { HasSourceSpan: true } ks && ks.Value != "<<") {
                        string text = ks.Value;
                        if (previous is not null && string.CompareOrdinal(previous, text) > 0) {
                            Add(_options.KeyOrdering, k._start, "Wrong ordering of key \"" + Shorten(text) + "\" in mapping.", fix: null);
                        }
                        previous = text;
                    }
                    Visit(v, key: false, inFlow);
                }
                break;
            }
            case YamlSequence sequence: {
                bool inFlow = flow || sequence._style == YamlStyle.Flow;
                for (int i = 0; i < sequence._count; i++) {
                    YamlNode item = sequence.Items[i];
                    if (!inFlow && _options.EmptyValues.Enabled && _options.EmptyValues.ForbidInBlockSequences && IsEmpty(item)) {
                        Add(_options.EmptyValues, YamlTrivia.DashPosition(_text, item._start), "Empty value in block sequence.", fix: null);
                    }
                    Visit(item, key: false, inFlow);
                }
                break;
            }
        }
    }

    private static bool IsEmpty(YamlNode node) =>
        node is YamlScalar { HasSourceSpan: true } scalar && scalar._style == YamlStyle.Plain && scalar._valueEnd == scalar._valueStart &&
        scalar.Tag is null && scalar.Anchor is null;

    /// <summary>Checks how far a block collection under a key is indented from that key.</summary>
    private void CheckNestedIndentation(YamlNode key, YamlNode value) {
        if (!_options.Indentation.Enabled || !key.HasSourceSpan || !value.HasSourceSpan || value._style != YamlStyle.Block) return;
        int keyColumn = Column(key._start), keyLine = LineOf(key._start);
        if (value is YamlMapping { _count: > 0 } mapping) {
            int child = mapping.Entries[0].Key._start;
            if (child < 0 || LineOf(child) == keyLine) return;
            int indent = Column(child) - keyColumn;
            int expected = _spaces ??= _options.Indentation.Spaces ?? (indent > 0 ? indent : 2);
            if (indent != expected) Add(_options.Indentation, child, "Wrong indentation: expected " + expected + " but found " + indent + ".", fix: null);
        } else if (value is YamlSequence { _count: > 0 } sequence && sequence.Anchor is null && sequence.Tag is null) {
            // A sequence with an anchor or tag on the key line is left alone: writers indent it so its first dash
            // cannot be misread, whatever style the file uses elsewhere.
            int dash = YamlTrivia.DashPosition(_text, sequence.Items[0]._start);
            if (LineOf(dash) == keyLine) return;
            int offset = Column(dash) - keyColumn;
            bool indented = offset > 0;
            YamlLintSequenceIndentation style = _options.Indentation.Sequences;
            bool? wanted = style switch {
                YamlLintSequenceIndentation.Indented => true,
                YamlLintSequenceIndentation.NotIndented => false,
                YamlLintSequenceIndentation.Consistent => _indentedSequences ??= indented,
                _ => null
            };
            if (wanted == false && indented) {
                Add(_options.Indentation, dash, "Wrong indentation: expected 0 but found " + offset + ".", fix: null);
            } else if (wanted == true && !indented) {
                Add(_options.Indentation, dash, "Wrong indentation: expected " + (_spaces ?? _options.Indentation.Spaces ?? 2) + " but found 0.", fix: null);
            } else if (indented) {
                int expected = _spaces ??= _options.Indentation.Spaces ?? offset;
                if (offset != expected) Add(_options.Indentation, dash, "Wrong indentation: expected " + expected + " but found " + offset + ".", fix: null);
            }
        }
    }

    private void CheckScalar(YamlScalar scalar, bool key, bool flow) {
        if (!scalar.HasSourceSpan || scalar.Tag is not null) return;
        if (scalar._style == YamlStyle.Plain) {
            CheckTruthy(scalar, key);
            CheckOctal(scalar);
        }
        CheckQuoting(scalar, key, flow);
    }

    private void CheckTruthy(YamlScalar scalar, bool key) {
        YamlLintTruthyOptions rule = _options.Truthy;
        if (!rule.Enabled || (key && !rule.CheckKeys)) return;
        string text = scalar.RawText;
        if (Array.IndexOf(TruthyWords, text) < 0 || rule.AllowedValues.Contains(text)) return;
        YamlScalarKind kind = scalar.Kind;
        string message;
        string? replacement;
        if (kind == YamlScalarKind.Boolean) {
            bool value = scalar.GetValue<bool>();
            replacement = value ? "true" : "false";
            message = "Truthy value \"" + text + "\" should be \"" + replacement + "\"" +
                (IsCoreBoolean(text) ? "." : ": YAML 1.2 readers read it as a string.");
        } else {
            replacement = YamlModelWriter.Render(new YamlScalar(text, PreferredQuote()));
            message = "Truthy value \"" + text + "\" is a string here, but YAML 1.1 readers such as PyYAML, go-yaml v2 and Ruby Psych read it as a boolean; quote it.";
        }
        Add(rule, scalar._start, message, ScalarFix(scalar, replacement));
    }

    private static bool IsCoreBoolean(string text) => text is "true" or "True" or "TRUE" or "false" or "False" or "FALSE";

    private void CheckOctal(YamlScalar scalar) {
        YamlLintOctalValuesOptions rule = _options.OctalValues;
        if (!rule.Enabled) return;
        string text = scalar.RawText;
        int digits = text.Length > 0 && (text[0] == '-' || text[0] == '+') ? 1 : 0;
        if (rule.ForbidImplicitOctal && text.Length - digits >= 2 && text[digits] == '0' && AllOctal(text, digits + 1)) {
            Add(rule, scalar._start, "Implicit octal value \"" + text + "\": YAML 1.1 readers read it as octal, YAML 1.2 readers as decimal; quote it or write it as 0o" + text.Substring(digits + 1) + ".", fix: null);
        } else if (rule.ForbidExplicitOctal && text.Length - digits >= 3 && text[digits] == '0' && text[digits + 1] == 'o' && AllOctal(text, digits + 2)) {
            Add(rule, scalar._start, "Explicit octal value \"" + text + "\": YAML 1.1 readers read it as a string.", fix: null);
        }
    }

    private static bool AllOctal(string text, int from) {
        for (int i = from; i < text.Length; i++) {
            if (text[i] < '0' || text[i] > '7') return false;
        }
        return true;
    }

    private void CheckQuoting(YamlScalar scalar, bool key, bool flow) {
        YamlLintQuotedStringsOptions rule = _options.QuotedStrings;
        if (!rule.Enabled || (key && !rule.CheckKeys)) return;
        YamlStyle style = scalar._style;
        if (style is YamlStyle.Literal or YamlStyle.Folded || scalar.Kind != YamlScalarKind.String) return;
        string value = scalar.Value;
        bool singleLine = LineOf(scalar._start) == LineOf(Math.Max(scalar._start, scalar._end - 1)) && value.IndexOf('\n') < 0;
        if (style == YamlStyle.Plain) {
            if (rule.Required != YamlLintQuoteRequirement.Always || (key && scalar._valueEnd == scalar._valueStart)) return;
            string? fix = singleLine ? YamlModelWriter.Render(new YamlScalar(value, PreferredQuote(rule.QuoteType))) : null;
            Add(rule, scalar._start, "String value is not quoted.", fix is null ? null : ScalarFix(scalar, fix));
            return;
        }
        if (rule.Required == YamlLintQuoteRequirement.OnlyWhenNeeded && singleLine && value.Length > 0 &&
            YamlScalarFormat.CanBePlain(value.AsSpan(), flow, key) && !YamlScalarFormat.IsSchemaAmbiguous(value.AsSpan())) {
            Add(rule, scalar._start, "String value is redundantly quoted.", ScalarFix(scalar, value));
            return;
        }
        YamlStyle wanted = rule.QuoteType switch {
            YamlLintQuoteType.Single => YamlStyle.SingleQuoted,
            YamlLintQuoteType.Double => YamlStyle.DoubleQuoted,
            _ => style
        };
        if (style == wanted) return;
        // Double quotes stay where the value needs escapes that single quotes cannot write.
        if (wanted == YamlStyle.SingleQuoted && !YamlScalarFormat.CanBeSingleQuoted(value.AsSpan())) return;
        string? requoted = singleLine ? YamlModelWriter.Render(new YamlScalar(value, wanted)) : null;
        Add(rule, scalar._start, "String value is not quoted with " + (wanted == YamlStyle.SingleQuoted ? "single" : "double") + " quotes.",
            requoted is null ? null : ScalarFix(scalar, requoted));
    }

    private YamlStyle PreferredQuote(YamlLintQuoteType type = YamlLintQuoteType.Any) => type switch {
        YamlLintQuoteType.Single => YamlStyle.SingleQuoted,
        YamlLintQuoteType.Double => YamlStyle.DoubleQuoted,
        _ => _document?.Layout.Quote ?? YamlStyle.SingleQuoted
    };

    /// <summary>Replaces a scalar's whole text, quotes included and anchor or tag excluded.</summary>
    private YamlLintFix? ScalarFix(YamlScalar scalar, string replacement) {
        if (!_fixable) return null;
        int start = scalar._valueStart, end = scalar._valueEnd;
        if (scalar._style is YamlStyle.SingleQuoted or YamlStyle.DoubleQuoted) {
            start--;
            end++;
        }
        if (start < 0 || end > _text.Length || end < start) return null;
        return new YamlLintFix(start, end - start, replacement);
    }

    // ------------------------------------------------------------------ line rules

    private void CheckLines() {
        int lines = LineCount();
        int blankRun = 0, blankRunStart = -1, lastContentLine = -1;
        bool seenContent = false;
        for (int line = 0; line < lines; line++) {
            int start = _lineStarts[line];
            int end = ContentEnd(line);
            bool blank = IsBlank(start, end);
            bool inScalar = InsideScalar(start);
            CheckTrailingSpaces(start, end, line);
            CheckLineLength(start, end);
            CheckTabs(start, end, blank, inScalar);
            if (blank && !inScalar) {
                if (blankRun == 0) blankRunStart = line;
                blankRun++;
                continue;
            }
            if (blankRun > 0) EndBlankRun(blankRunStart, blankRun, seenContent ? _options.EmptyLines.Max : _options.EmptyLines.MaxStart, atStart: !seenContent);
            blankRun = 0;
            seenContent = true;
            lastContentLine = line;
        }
        if (blankRun > 0) EndBlankRun(blankRunStart, blankRun, seenContent ? _options.EmptyLines.MaxEnd : _options.EmptyLines.MaxStart, atStart: !seenContent, atEnd: seenContent);
        CheckFinalNewLine();
    }

    private void CheckTrailingSpaces(int start, int end, int line) {
        if (!_options.TrailingSpaces.Enabled || end == start) return;
        int trimmed = end;
        while (trimmed > start && (_text[trimmed - 1] == ' ' || _text[trimmed - 1] == '\t')) trimmed--;
        if (trimmed == end) return;
        // Spaces before a line break inside a scalar can be content (block scalars); those are reported, not removed.
        bool content = InsideScalar(end) || EndsBlockScalar(end);
        Add(_options.TrailingSpaces, trimmed, "Trailing spaces.", content || !_fixable ? null : new YamlLintFix(trimmed, end - trimmed, string.Empty));
    }

    private void CheckLineLength(int start, int end) {
        YamlLintLineLengthOptions rule = _options.LineLength;
        if (!rule.Enabled || end - start <= rule.Max) return;
        if (rule.AllowNonBreakableWords || rule.AllowNonBreakableInlineMappings) {
            int i = start;
            while (i < end && _text[i] == ' ') i++;
            if (i < end && _text[i] == '#') {
                while (i < end && _text[i] == '#') i++;
                i++;
            } else if (i + 1 < end && _text[i] == '-' && _text[i + 1] == ' ') {
                i += 2;
            }
            if (rule.AllowNonBreakableWords && NoSpace(i, end)) return;
            if (rule.AllowNonBreakableInlineMappings) {
                int colon = _text.IndexOf(": ", i, end - i, StringComparison.Ordinal);
                if (colon > 0) {
                    int value = colon + 2;
                    while (value < end && _text[value] == ' ') value++;
                    if (value < end && NoSpace(value, end)) return;
                }
            }
        }
        Add(rule, start + rule.Max, "Line too long (" + (end - start) + " > " + rule.Max + " characters).", fix: null);
    }

    private bool NoSpace(int from, int end) {
        for (int i = from; i < end; i++) {
            if (_text[i] == ' ' || _text[i] == '\t') return false;
        }
        return true;
    }

    private void CheckTabs(int start, int end, bool blank, bool inScalar) {
        if (!_options.Tabs.Enabled || blank || inScalar) return;
        for (int i = start; i < end && (_text[i] == ' ' || _text[i] == '\t'); i++) {
            if (_text[i] == '\t') {
                Add(_options.Tabs, i, "Tab character in indentation.", fix: null);
                return;
            }
        }
    }

    private void EndBlankRun(int firstLine, int count, int max, bool atStart, bool atEnd = false) {
        if (!_options.EmptyLines.Enabled || count <= max) return;
        int firstExtra = firstLine + max;
        int from = _lineStarts[firstExtra];
        int lastLine = firstLine + count - 1;
        int to = lastLine + 1 < _lineStarts.Length ? _lineStarts[lastLine + 1] : _text.Length;
        string message = atStart && !atEnd ? "Too many blank lines at the start of the file (" + count + " > " + max + ")."
            : atEnd ? "Too many blank lines at the end of the file (" + count + " > " + max + ")."
            : "Too many blank lines (" + count + " > " + max + ").";
        // At the end, the last kept line keeps its line break, so the removal starts at the first extra line.
        Add(_options.EmptyLines, from, message, _fixable ? new YamlLintFix(from, to - from, string.Empty) : null);
    }

    private void CheckFinalNewLine() {
        if (!_options.NewLineAtEndOfFile.Enabled || _text.Length == 0) return;
        char last = _text[_text.Length - 1];
        if (last == '\n' || last == '\r') return;
        // A block scalar that runs to the end would gain the line break in its value.
        bool blockAtEnd = _scalars.Count > 0 && _scalars[_scalars.Count - 1] is { Style: YamlStyle.Literal or YamlStyle.Folded } range && range.End >= _text.Length;
        Add(_options.NewLineAtEndOfFile, _text.Length, "No new line character at the end of the file.",
            _fixable && !blockAtEnd ? new YamlLintFix(_text.Length, 0, DominantNewLine()) : null);
    }

    // ------------------------------------------------------------------ comments

    private void CheckComments() {
        YamlLintCommentsOptions rule = _options.Comments;
        if (!rule.Enabled) return;
        foreach ((int hash, int end) in _comments) {
            if (hash >= _text.Length || _text[hash] != '#') continue;
            int lineStart = YamlTrivia.LineStart(_text, hash);
            int before = hash;
            while (before > lineStart && (_text[before - 1] == ' ' || _text[before - 1] == '\t')) before--;
            bool inline = before > lineStart;
            if (inline && rule.MinSpacesFromContent > 0 && hash - before < rule.MinSpacesFromContent) {
                int missing = rule.MinSpacesFromContent - (hash - before);
                Add(rule, hash, "Too few spaces before comment: expected " + rule.MinSpacesFromContent + ".",
                    _fixable ? new YamlLintFix(hash, 0, new string(' ', missing)) : null);
            }
            if (!rule.RequireStartingSpace) continue;
            if (rule.IgnoreShebangs && hash == 0 && hash + 1 < _text.Length && _text[1] == '!') continue;
            int text = hash;
            while (text < end && _text[text] == '#') text++;
            if (text >= end || _text[text] == ' ' || _text[text] == '\t') continue;
            Add(rule, text, "Missing starting space in comment.", _fixable ? new YamlLintFix(text, 0, " ") : null);
        }
    }

    // ------------------------------------------------------------------ helpers

    private void Add(YamlLintRuleOptions rule, int offset, string message, YamlLintFix? fix) {
        if (!rule.Enabled) return;
        _findings.Add(new YamlLintFinding(rule.Id, rule.Severity, Position(offset), message, fix));
    }

    private string DominantNewLine() {
        if (_document is not null) return _document.Layout.NewLine;
        int lf = _text.IndexOf('\n');
        return lf > 0 && _text[lf - 1] == '\r' ? "\r\n" : "\n";
    }

    private static string FirstLine(string message) {
        int end = message.IndexOfAny(['\r', '\n']);
        return end < 0 ? message : message.Substring(0, end);
    }

    private static string Shorten(string text) => text.Length > 64 ? text.Substring(0, 61) + "..." : text;

    private int DiagnosticOffset(OfficeDiagnostic diagnostic) {
        string? location = diagnostic.Location;
        int colon = location?.IndexOf(':') ?? -1;
        if (colon > 0 && int.TryParse(location!.Substring(0, colon), NumberStyles.None, CultureInfo.InvariantCulture, out int line) &&
            int.TryParse(location.Substring(colon + 1), NumberStyles.None, CultureInfo.InvariantCulture, out int column) &&
            line >= 1 && line <= _lineStarts.Length) {
            return Math.Min(_lineStarts[line - 1] + column - 1, _text.Length);
        }
        return 0;
    }

    private int Offset(YamlPosition position) => Math.Max(0, Math.Min((int)position.Offset, _text.Length));

    private YamlPosition Clamp(YamlPosition position) => Position(Offset(position));

    private YamlPosition Position(int offset) {
        int line = LineOf(offset);
        return new YamlPosition(offset, line + 1, offset - _lineStarts[line] + 1);
    }

    private int LineOf(int offset) {
        int low = 0, high = _lineStarts.Length - 1;
        while (low < high) {
            int middle = (low + high + 1) >> 1;
            if (_lineStarts[middle] <= offset) low = middle;
            else high = middle - 1;
        }
        return low;
    }

    private int Column(int offset) => offset - _lineStarts[LineOf(offset)];

    /// <summary>Lines that hold text: the empty line after a final line break is not one.</summary>
    private int LineCount() => _lineStarts.Length > 1 && _lineStarts[_lineStarts.Length - 1] == _text.Length ? _lineStarts.Length - 1 : _lineStarts.Length;

    private int ContentEnd(int line) => YamlTrivia.LineContentEnd(_text, _lineStarts[line]);

    private bool IsBlank(int start, int end) {
        for (int i = start; i < end; i++) {
            if (_text[i] != ' ' && _text[i] != '\t') return false;
        }
        return true;
    }

    private static int[] BuildLineStarts(string text) {
        var starts = new List<int> { 0 };
        for (int i = 0; i < text.Length; i++) {
            char ch = text[i];
            if (ch == '\r') {
                if (i + 1 < text.Length && text[i + 1] == '\n') i++;
                starts.Add(i + 1);
            } else if (ch == '\n') {
                starts.Add(i + 1);
            }
        }
        return starts.ToArray();
    }
}
