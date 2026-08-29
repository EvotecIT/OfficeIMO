namespace OfficeIMO.Bibliography;

internal static class BibCodec {
    private static readonly HashSet<string> ReservedTypedFieldNames = new HashSet<string>(new[] {
        "title", "author", "editor", "translator", "journal", "journaltitle", "booktitle", "publisher", "institution", "organization", "address", "location", "edition", "volume", "number", "issue", "pages", "eid", "abstract", "language", "langid", "url", "date", "year", "month", "urldate", "keywords", "note"
    }, StringComparer.OrdinalIgnoreCase);

    internal static IList<BibliographyItem> Parse(string source, BibliographyFormat format, BibliographyReadOptions options, List<BibliographyDiagnostic> diagnostics, IList<BibliographyNativeEntry> nativeEntries, CancellationToken cancellationToken) {
        var parser = new Parser(source, format, options, diagnostics, nativeEntries, cancellationToken);
        return parser.Parse();
    }

    internal static string Write(BibliographyDocument document, BibliographyFormat format, BibliographyWriteOptions options, BibliographyConversionReport report, CancellationToken cancellationToken) {
        var builder = new StringBuilder();
        foreach (BibliographyNativeEntry entry in document.NativeEntries.Where(entry => IsBibFamily(entry.Format) && IsBibFamily(format))) {
            cancellationToken.ThrowIfCancellationRequested();
            if (entry.Kind == "string") builder.Append("@string{").Append(entry.Name).Append(" = {").Append(Escape(entry.Value)).Append("}}").Append(options.LineEnding).Append(options.LineEnding);
            else if (entry.Kind == "preamble") builder.Append("@preamble{{").Append(Escape(entry.Value)).Append("}}").Append(options.LineEnding).Append(options.LineEnding);
            else if (entry.Kind == "comment") builder.Append("@comment{").Append(entry.Value).Append('}').Append(options.LineEnding).Append(options.LineEnding);
            else if (entry.Kind == "line-comment") builder.Append('%').Append(entry.Value).Append(options.LineEnding);
            if (entry.Kind == "string" || entry.Kind == "preamble" || entry.Kind == "comment" || entry.Kind == "line-comment") report.Add("BIBCONV010", BibliographyDiagnosticSeverity.Information, $"Preserved native BibTeX @{entry.Kind} entry.", BibliographyConversionAction.PreservedExtension, field: entry.Name ?? entry.Kind);
            else report.Add("BIBCONV118", BibliographyDiagnosticSeverity.Warning, $"Native BibTeX document entry '{entry.Kind}' is not safe to write.", BibliographyConversionAction.Omitted, field: entry.Name ?? entry.Kind);
        }
        foreach (BibliographyNativeEntry entry in document.NativeEntries.Where(entry => !IsBibFamily(entry.Format))) {
            report.Add("BIBCONV110", BibliographyDiagnosticSeverity.Warning, $"Document-level {entry.Format} entry '{entry.Kind}' cannot be represented in {format}.", BibliographyConversionAction.Omitted, field: entry.Name ?? entry.Kind);
        }

        for (int itemIndex = 0; itemIndex < document.Items.Count; itemIndex++) {
            BibliographyItem item = document.Items[itemIndex];
            cancellationToken.ThrowIfCancellationRequested();
            string type = item.Type == BibliographyItemType.Unknown && IsSafeTypeName(item.NativeType) ? item.NativeType! : CodecMappings.ToBibType(item.Type);
            builder.Append('@').Append(type.ToLowerInvariant()).Append('{').Append(SafeKey(CodecMappings.OutputKey(item, itemIndex))).Append(',').Append(options.LineEnding);
            var fields = new List<KeyValuePair<string, string>>();
            Add(fields, "title", item.Title);
            AddNames(fields, "author", item, BibliographyContributorRole.Author);
            AddNames(fields, "editor", item, BibliographyContributorRole.Editor);
            AddNames(fields, "translator", item, BibliographyContributorRole.Translator);
            Add(fields, format == BibliographyFormat.BibLatex ? "journaltitle" : "journal", item.ContainerTitle);
            Add(fields, "booktitle", item.CollectionTitle);
            Add(fields, "publisher", item.Publisher);
            Add(fields, "address", item.PublisherPlace);
            Add(fields, "edition", item.Edition);
            Add(fields, "volume", item.Volume);
            Add(fields, "number", item.Issue);
            Add(fields, "pages", item.Pages);
            Add(fields, "abstract", item.Abstract);
            Add(fields, "language", item.Language);
            Add(fields, "url", item.Url);
            BibliographyDate? issued = item.GetDate(BibliographyDateRole.Issued);
            if (issued != null) {
                if (format == BibliographyFormat.BibLatex) Add(fields, "date", CodecMappings.FormatDate(issued));
                else {
                    Add(fields, "year", issued.Year?.ToString(CultureInfo.InvariantCulture) ?? issued.Literal);
                    Add(fields, "month", issued.Month?.ToString(CultureInfo.InvariantCulture));
                }
            }
            BibliographyDate? accessed = item.GetDate(BibliographyDateRole.Accessed);
            if (accessed != null) Add(fields, "urldate", CodecMappings.FormatDate(accessed));
            foreach (BibliographyIdentifier identifier in item.Identifiers) {
                string fieldName = identifier.Scheme.ToLowerInvariant();
                if (IsSafeFieldName(fieldName) && !ReservedTypedFieldNames.Contains(fieldName)) Add(fields, fieldName, identifier.Value);
                else report.Add("BIBCONV129", BibliographyDiagnosticSeverity.Warning, $"Identifier scheme '{identifier.Scheme}' cannot be represented as a safe, non-conflicting BibTeX field.", BibliographyConversionAction.Omitted, item, "identifiers." + identifier.Scheme);
            }
            if (item.Keywords.Count > 0) Add(fields, "keywords", string.Join(", ", item.Keywords));
            if (item.Notes.Count > 0) Add(fields, "note", string.Join("; ", item.Notes));

            var emitted = new HashSet<string>(fields.Select(static pair => pair.Key), StringComparer.OrdinalIgnoreCase);
            foreach (BibliographyNativeField field in item.NativeFields) {
                if (IsBibFamily(field.Format) && IsSafeFieldName(field.Name) && !emitted.Contains(field.Name)) {
                    fields.Add(new KeyValuePair<string, string>(field.Name.ToLowerInvariant(), field.Value));
                    emitted.Add(field.Name);
                    report.Add("BIBCONV011", BibliographyDiagnosticSeverity.Information, $"Preserved native field '{field.Name}'.", BibliographyConversionAction.PreservedExtension, item, field.Name);
                } else if (!IsBibFamily(field.Format)) {
                    report.Add("BIBCONV111", BibliographyDiagnosticSeverity.Warning, $"Native {field.Format} field '{field.Name}' cannot be represented safely in {format}.", BibliographyConversionAction.Omitted, item, field.Name);
                } else {
                    report.Add("BIBCONV119", BibliographyDiagnosticSeverity.Warning, $"Native {format} field '{field.Name}' conflicts with a typed field or has an unsafe name.", BibliographyConversionAction.Omitted, item, field.Name);
                }
            }

            for (int index = 0; index < fields.Count; index++) {
                KeyValuePair<string, string> field = fields[index];
                builder.Append("  ").Append(field.Key).Append(" = {").Append(Escape(field.Value)).Append('}');
                if (index + 1 < fields.Count) builder.Append(',');
                builder.Append(options.LineEnding);
            }
            builder.Append('}').Append(options.LineEnding);
            if (itemIndex + 1 < document.Items.Count) builder.Append(options.LineEnding);
        }
        return builder.ToString();
    }

    private static void Add(ICollection<KeyValuePair<string, string>> fields, string name, string? value) {
        if (!string.IsNullOrWhiteSpace(value)) fields.Add(new KeyValuePair<string, string>(name, value!));
    }

    private static void AddNames(ICollection<KeyValuePair<string, string>> fields, string name, BibliographyItem item, BibliographyContributorRole role) {
        string[] names = item.Contributors.Where(contributor => contributor.Role == role).Select(contributor => FormatBibName(contributor.Name)).Where(static value => value.Length > 0).ToArray();
        if (names.Length > 0) Add(fields, name, string.Join(" and ", names));
    }

    private static string FormatBibName(BibliographyName name) =>
        string.IsNullOrWhiteSpace(name.Literal) ? CodecMappings.FormatName(name) : "{" + name.Literal + "}";

    private static string Escape(string value) {
        int depth = 0;
        for (int index = 0; index < value.Length; index++) {
            if (value[index] == '\\' && index + 1 < value.Length) { index++; continue; }
            if (value[index] == '{') depth++;
            else if (value[index] == '}') { if (depth == 0) return EscapeAllBraces(value); depth--; }
        }
        return depth == 0 ? value : EscapeAllBraces(value);
    }
    private static string EscapeAllBraces(string value) => value.Replace("{", "\\{").Replace("}", "\\}");
    private static string SafeKey(string key) => string.IsNullOrWhiteSpace(key) ? "item" : new string(key.Select(character => char.IsWhiteSpace(character) || character == ',' || character == '}' ? '_' : character).ToArray());
    private static bool IsSafeFieldName(string name) => name.Length > 0 && name.All(character => char.IsLetterOrDigit(character) || character == '-' || character == '_' || character == ':');
    private static bool IsSafeTypeName(string? name) => !string.IsNullOrWhiteSpace(name) && name!.All(character => char.IsLetterOrDigit(character) || character == '-' || character == '_' || character == ':' || character == '.');
    private static bool IsBibFamily(BibliographyFormat format) => format == BibliographyFormat.BibTex || format == BibliographyFormat.BibLatex;

    private sealed class Parser {
        private readonly string _source;
        private readonly BibliographyFormat _format;
        private readonly List<BibliographyDiagnostic> _diagnostics;
        private readonly IList<BibliographyNativeEntry> _nativeEntries;
        private readonly CancellationToken _cancellationToken;
        private readonly BibliographyLimitGuard _limits;
        private readonly int _maximumDiagnosticCount;
        private readonly Dictionary<string, string> _strings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private readonly List<BibliographyItem> _items = new List<BibliographyItem>();
        private int _position;
        private int _locationOffset;
        private int _locationLine = 1;
        private int _locationColumn = 1;
        private bool _diagnosticLimitReported;

        internal Parser(string source, BibliographyFormat format, BibliographyReadOptions options, List<BibliographyDiagnostic> diagnostics, IList<BibliographyNativeEntry> nativeEntries, CancellationToken cancellationToken) {
            _source = source; _format = format; _diagnostics = diagnostics; _nativeEntries = nativeEntries; _cancellationToken = cancellationToken; _limits = new BibliographyLimitGuard(options); _maximumDiagnosticCount = options.MaximumDiagnosticCount;
        }

        internal IList<BibliographyItem> Parse() {
            while (_position < _source.Length) {
                _cancellationToken.ThrowIfCancellationRequested();
                SkipTrivia();
                if (_position >= _source.Length) break;
                if (_source[_position] != '@') {
                    int invalidStart = _position;
                    while (_position < _source.Length && _source[_position] != '@') _position++;
                    AddDiagnostic("BIBBIB001", "Ignored text outside a BibTeX entry.", invalidStart);
                    continue;
                }
                int entryStart = _position++;
                string type = ReadIdentifier();
                SkipWhitespace();
                if (_position >= _source.Length || (_source[_position] != '{' && _source[_position] != '(')) { AddDiagnostic("BIBBIB002", "Expected '{' or '(' after the BibTeX entry type.", entryStart, severity: BibliographyDiagnosticSeverity.Error); RecoverToNextEntry(); continue; }
                char open = _source[_position++];
                char close = open == '{' ? '}' : ')';
                if (string.Equals(type, "comment", StringComparison.OrdinalIgnoreCase)) { string value = ReadBalancedRaw(close); _limits.AddValue(_items, value, entryStart); _nativeEntries.Add(new BibliographyNativeEntry(_format, "comment", value)); continue; }
                if (string.Equals(type, "preamble", StringComparison.OrdinalIgnoreCase)) { string value = ReadValue(close); _limits.AddValue(_items, value, entryStart); ConsumeClose(close); _nativeEntries.Add(new BibliographyNativeEntry(_format, "preamble", value)); continue; }
                if (string.Equals(type, "string", StringComparison.OrdinalIgnoreCase)) { ParseString(close, entryStart); continue; }
                ParseItem(type, close, entryStart);
            }
            return _items;
        }

        private void ParseString(char close, int entryStart) {
            SkipWhitespace();
            string name = ReadIdentifier();
            SkipWhitespace();
            if (!Consume('=')) AddDiagnostic("BIBBIB003", "Expected '=' in a BibTeX string directive.", _position, severity: BibliographyDiagnosticSeverity.Error);
            string value = ReadValue(close);
            _limits.AddValue(_items, value, entryStart);
            _strings[name] = value;
            _nativeEntries.Add(new BibliographyNativeEntry(_format, "string", value, name));
            ConsumeClose(close);
        }

        private void ParseItem(string nativeType, char close, int entryStart) {
            _limits.AddItem(_items, entryStart);
            SkipWhitespace();
            int keyStart = _position;
            while (_position < _source.Length && _source[_position] != ',' && _source[_position] != close) _position++;
            string key = _source.Substring(keyStart, _position - keyStart).Trim();
            var item = new BibliographyItem { Key = key, NativeType = nativeType, Type = CodecMappings.ParseType(nativeType) };
            _items.Add(item);
            if (string.IsNullOrWhiteSpace(key)) AddDiagnostic("BIBBIB010", "BibTeX entry has no citation key.", keyStart, severity: BibliographyDiagnosticSeverity.Warning);
            if (_position < _source.Length && _source[_position] == close) { _position++; return; }
            Consume(',');
            while (_position < _source.Length) {
                _cancellationToken.ThrowIfCancellationRequested();
                SkipWhitespaceAndCommas();
                if (_position >= _source.Length) { AddDiagnostic("BIBBIB004", "BibTeX entry ended before its closing delimiter.", entryStart, item.Key, severity: BibliographyDiagnosticSeverity.Error); return; }
                if (_source[_position] == close) { _position++; return; }
                int fieldStart = _position;
                string name = ReadIdentifier();
                if (name.Length == 0) { AddDiagnostic("BIBBIB005", "Expected a BibTeX field name.", fieldStart, item.Key, severity: BibliographyDiagnosticSeverity.Error); RecoverToDelimiter(close); continue; }
                SkipWhitespace();
                if (!Consume('=')) { AddDiagnostic("BIBBIB006", "Expected '=' after a BibTeX field name.", _position, item.Key, name, BibliographyDiagnosticSeverity.Error); RecoverToDelimiter(close); continue; }
                string value = ReadValue(close);
                _limits.AddValue(_items, value, fieldStart);
                Bind(item, name, value);
            }
        }

        private string ReadValue(char entryClose) {
            var builder = new StringBuilder();
            while (true) {
                SkipWhitespace();
                if (_position >= _source.Length || _source[_position] == entryClose || _source[_position] == ',') break;
                if (_source[_position] == '{') builder.Append(ReadDelimited('{', '}'));
                else if (_source[_position] == '"') builder.Append(ReadDelimited('"', '"'));
                else {
                    string atom = ReadValueAtom(entryClose);
                    if (_strings.TryGetValue(atom, out string? expanded)) builder.Append(expanded); else builder.Append(atom);
                }
                SkipWhitespace();
                if (_position < _source.Length && _source[_position] == '#') { _position++; continue; }
                break;
            }
            return builder.ToString();
        }

        private string ReadDelimited(char open, char close) {
            int start = _position++;
            var builder = new StringBuilder();
            int depth = 1;
            while (_position < _source.Length) {
                _cancellationToken.ThrowIfCancellationRequested();
                char current = _source[_position++];
                if (current == '\\' && _position < _source.Length) { builder.Append(current).Append(_source[_position++]); continue; }
                if (open != '"' && current == open) { depth++; _limits.CheckDepth(_items, depth, _position - 1); if (depth > 1) builder.Append(current); continue; }
                if (current == close) { depth--; if (depth == 0) return builder.ToString(); builder.Append(current); continue; }
                builder.Append(current);
            }
            AddDiagnostic("BIBBIB007", "Delimited BibTeX value was not closed.", start, severity: BibliographyDiagnosticSeverity.Error);
            return builder.ToString();
        }

        private string ReadBalancedRaw(char close) {
            int start = _position;
            int depth = 1;
            while (_position < _source.Length) {
                char current = _source[_position++];
                if (current == '\\' && _position < _source.Length) { _position++; continue; }
                if (current == close) { depth--; if (depth == 0) return _source.Substring(start, _position - start - 1); }
                else if (current == (close == '}' ? '{' : '(')) { depth++; _limits.CheckDepth(_items, depth, _position - 1); }
            }
            AddDiagnostic("BIBBIB008", "BibTeX directive was not closed.", start, severity: BibliographyDiagnosticSeverity.Error);
            return _source.Substring(start);
        }

        private string ReadValueAtom(char close) {
            int start = _position;
            while (_position < _source.Length && _source[_position] != '#' && _source[_position] != ',' && _source[_position] != close && !char.IsWhiteSpace(_source[_position])) _position++;
            return _source.Substring(start, _position - start);
        }

        private void Bind(BibliographyItem item, string name, string value) {
            string field = name.ToLowerInvariant();
            switch (field) {
                case "title": item.Title = value; break;
                case "journal": case "journaltitle": item.ContainerTitle = value; break;
                case "booktitle": item.CollectionTitle = value; break;
                case "publisher": case "institution": case "organization": item.Publisher = value; break;
                case "address": case "location": item.PublisherPlace = value; break;
                case "edition": item.Edition = value; break;
                case "volume": item.Volume = value; break;
                case "number": case "issue": item.Issue = value; break;
                case "pages": case "eid": item.Pages = value; break;
                case "abstract": item.Abstract = value; break;
                case "language": case "langid": item.Language = value; break;
                case "url": item.Url = value; break;
                case "author": AddNames(item, BibliographyContributorRole.Author, value); break;
                case "editor": AddNames(item, BibliographyContributorRole.Editor, value); break;
                case "translator": AddNames(item, BibliographyContributorRole.Translator, value); break;
                case "date": item.Dates.Add(CodecMappings.ParseDate(BibliographyDateRole.Issued, value)); break;
                case "year": SetYear(item, value); break;
                case "month": SetMonth(item, value); break;
                case "urldate": item.Dates.Add(CodecMappings.ParseDate(BibliographyDateRole.Accessed, value)); break;
                case "doi": case "isbn": case "issn": case "pmid": case "pmcid": item.Identifiers.Add(new BibliographyIdentifier(field, value)); break;
                case "keywords": foreach (string keyword in value.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)) item.Keywords.Add(keyword.Trim()); break;
                case "note": item.Notes.Add(value); break;
                default: item.NativeFields.Add(new BibliographyNativeField(_format, name, value)); break;
            }
        }

        private void AddNames(BibliographyItem item, BibliographyContributorRole role, string value) {
            foreach (string part in SplitNames(value)) { _limits.AddValue(_items, part, _position); item.Contributors.Add(new BibliographyContributor(role, CodecMappings.ParseCommaName(part))); }
        }

        private static IEnumerable<string> SplitNames(string value) {
            int start = 0;
            int depth = 0;
            for (int index = 0; index <= value.Length - 5; index++) {
                if (value[index] == '{') depth++; else if (value[index] == '}') depth--;
                if (depth == 0 && string.Equals(value.Substring(index, 5), " and ", StringComparison.OrdinalIgnoreCase)) { yield return value.Substring(start, index - start).Trim(); start = index + 5; index += 4; }
            }
            if (start <= value.Length) yield return value.Substring(start).Trim();
        }

        private static void SetYear(BibliographyItem item, string value) {
            BibliographyDate date = item.GetDate(BibliographyDateRole.Issued) ?? new BibliographyDate { Role = BibliographyDateRole.Issued };
            if (!item.Dates.Contains(date)) item.Dates.Add(date);
            if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int year)) date.Year = year; else date.Literal = value;
        }

        private static void SetMonth(BibliographyItem item, string value) {
            BibliographyDate date = item.GetDate(BibliographyDateRole.Issued) ?? new BibliographyDate { Role = BibliographyDateRole.Issued };
            if (!item.Dates.Contains(date)) item.Dates.Add(date);
            int? month = CodecMappings.ParseMonth(value);
            if (month.HasValue) date.Month = month; else date.Literal = string.IsNullOrEmpty(date.Literal) ? value : date.Literal + " " + value;
        }

        private string ReadIdentifier() {
            SkipWhitespace();
            int start = _position;
            while (_position < _source.Length && (char.IsLetterOrDigit(_source[_position]) || _source[_position] == '-' || _source[_position] == '_' || _source[_position] == ':' || _source[_position] == '.')) _position++;
            return _source.Substring(start, _position - start);
        }

        private bool Consume(char value) { SkipWhitespace(); if (_position < _source.Length && _source[_position] == value) { _position++; return true; } return false; }
        private void ConsumeClose(char close) { SkipWhitespaceAndCommas(); if (!Consume(close)) AddDiagnostic("BIBBIB009", $"Expected closing '{close}'.", _position, severity: BibliographyDiagnosticSeverity.Error); }
        private void SkipWhitespace() { while (_position < _source.Length && char.IsWhiteSpace(_source[_position])) _position++; }
        private void SkipWhitespaceAndCommas() { while (_position < _source.Length && (char.IsWhiteSpace(_source[_position]) || _source[_position] == ',')) _position++; }
        private void SkipTrivia() { while (_position < _source.Length) { if (char.IsWhiteSpace(_source[_position])) { _position++; continue; } if (_source[_position] == '%') { int start = ++_position; while (_position < _source.Length && _source[_position] != '\n' && _source[_position] != '\r') _position++; string value = _source.Substring(start, _position - start); _limits.AddValue(_items, value, start); _nativeEntries.Add(new BibliographyNativeEntry(_format, "line-comment", value)); continue; } break; } }
        private void RecoverToNextEntry() { int next = _source.IndexOf('@', _position); _position = next < 0 ? _source.Length : next; }
        private void RecoverToDelimiter(char close) { while (_position < _source.Length && _source[_position] != ',' && _source[_position] != close) _position++; }
        private void AddDiagnostic(string code, string message, int offset, string? key = null, string? field = null, BibliographyDiagnosticSeverity severity = BibliographyDiagnosticSeverity.Warning) {
            GetLocation(offset, out int line, out int column);
            if (_diagnostics.Count >= _maximumDiagnosticCount) {
                if (!_diagnosticLimitReported) {
                    _diagnostics.Add(new BibliographyDiagnostic("BIBLIM002", BibliographyDiagnosticSeverity.Error, "Maximum bibliography diagnostic count was exceeded.", offset, line, column));
                    _diagnosticLimitReported = true;
                }
                _position = _source.Length;
                return;
            }
            _diagnostics.Add(new BibliographyDiagnostic(code, severity, message, offset, line, column, key, field));
        }

        private void GetLocation(int offset, out int line, out int column) {
            if (offset < _locationOffset) { _locationOffset = 0; _locationLine = 1; _locationColumn = 1; }
            int target = Math.Min(offset, _source.Length);
            while (_locationOffset < target) {
                if (_source[_locationOffset++] == '\n') { _locationLine++; _locationColumn = 1; } else _locationColumn++;
            }
            line = _locationLine; column = _locationColumn;
        }
    }
}
