namespace OfficeIMO.Bibliography;

internal static class BibCodec {
    private static readonly HashSet<string> ClassicBibTypes = new HashSet<string>(new[] {
        "article", "book", "booklet", "conference", "inbook", "incollection", "inproceedings", "manual", "mastersthesis", "misc", "phdthesis", "proceedings", "techreport", "unpublished"
    }, StringComparer.OrdinalIgnoreCase);
    private static readonly HashSet<string> BibLatexOnlyTypedFields = new HashSet<string>(new[] { "journaltitle", "location", "issue", "eid", "langid" }, StringComparer.OrdinalIgnoreCase);
    private static readonly HashSet<string> ReservedTypedFieldNames = new HashSet<string>(new[] {
        "title", "author", "editor", "translator", "journal", "journaltitle", "booktitle", "series", "publisher", "institution", "organization", "address", "location", "edition", "volume", "number", "issue", "pages", "eid", "abstract", "language", "langid", "url", "date", "year", "month", "urldate", "keywords", "note"
    }, StringComparer.OrdinalIgnoreCase);
    private static readonly HashSet<string> TypedIdentifierFieldNames = new HashSet<string>(new[] { "doi", "isbn", "issn", "pmid", "pmcid" }, StringComparer.OrdinalIgnoreCase);

    internal static IList<BibliographyItem> Parse(string source, BibliographyFormat format, BibliographyReadOptions options, List<BibliographyDiagnostic> diagnostics, IList<BibliographyNativeEntry> nativeEntries, CancellationToken cancellationToken) {
        var parser = new Parser(source, format, options, diagnostics, nativeEntries, cancellationToken);
        return parser.Parse();
    }

    internal static string Write(BibliographyDocument document, BibliographyFormat format, BibliographyWriteOptions options, BibliographyConversionReport report, CancellationToken cancellationToken) {
        var builder = new StringBuilder();
        string[] outputKeys = CodecMappings.OutputKeys(document.Items, format, cancellationToken);
        foreach (BibliographyNativeEntry entry in document.NativeEntries.Where(entry => IsBibFamily(entry.Format) && IsBibFamily(format))) {
            cancellationToken.ThrowIfCancellationRequested();
            if (TryWriteNativeEntry(builder, entry, options.LineEnding, cancellationToken)) report.Add("BIBCONV010", BibliographyDiagnosticSeverity.Information, $"Preserved native BibTeX @{entry.Kind} entry.", BibliographyConversionAction.PreservedExtension, field: entry.Name ?? entry.Kind);
            else report.Add("BIBCONV118", BibliographyDiagnosticSeverity.Warning, $"Native BibTeX document entry '{entry.Kind}' is not safe to write.", BibliographyConversionAction.Omitted, field: entry.Name ?? entry.Kind);
        }
        foreach (BibliographyNativeEntry entry in document.NativeEntries.Where(entry => !IsBibFamily(entry.Format))) {
            cancellationToken.ThrowIfCancellationRequested();
            report.Add("BIBCONV110", BibliographyDiagnosticSeverity.Warning, $"Document-level {entry.Format} entry '{entry.Kind}' cannot be represented in {format}.", BibliographyConversionAction.Omitted, field: entry.Name ?? entry.Kind);
        }

        for (int itemIndex = 0; itemIndex < document.Items.Count; itemIndex++) {
            BibliographyItem item = document.Items[itemIndex];
            cancellationToken.ThrowIfCancellationRequested();
            string type = CanPreserveNativeType(document.SourceFormat, format, item) ? item.NativeType! : OutputType(item.Type, format);
            builder.Append('@').Append(type).Append('{').Append(outputKeys[itemIndex]).Append(',').Append(options.LineEnding);
            var fields = new List<KeyValuePair<string, string>>();
            Add(fields, "title", item.Title);
            AddNames(fields, "author", item, BibliographyContributorRole.Author, cancellationToken);
            AddNames(fields, "editor", item, BibliographyContributorRole.Editor, cancellationToken);
            AddNames(fields, "translator", item, BibliographyContributorRole.Translator, cancellationToken);
            Add(fields, GetBibFieldName(item, "container-title", DefaultContainerField(item, format), format), item.ContainerTitle);
            Add(fields, "series", item.CollectionTitle);
            Add(fields, GetBibFieldName(item, "publisher", "publisher", format), item.Publisher);
            Add(fields, GetBibFieldName(item, "publisher-place", format == BibliographyFormat.BibLatex ? "location" : "address", format), item.PublisherPlace);
            Add(fields, "edition", item.Edition);
            Add(fields, "volume", item.Volume);
            Add(fields, GetBibFieldName(item, "issue", "number", format), item.Issue);
            Add(fields, GetBibFieldName(item, "pages", "pages", format), item.Pages);
            Add(fields, "abstract", item.Abstract);
            Add(fields, GetBibFieldName(item, "language", "language", format), item.Language);
            Add(fields, "url", item.Url);
            if (format == BibliographyFormat.BibLatex) {
                var emittedDateRoles = new HashSet<BibliographyDateRole>();
                foreach (BibliographyDate date in item.Dates) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!emittedDateRoles.Add(date.Role)) continue;
                    if (date.Role == BibliographyDateRole.Issued) Add(fields, "date", CodecMappings.FormatDate(date));
                    else if (date.Role == BibliographyDateRole.Accessed) Add(fields, "urldate", CodecMappings.FormatDate(date));
                }
            } else {
                BibliographyDate? issued = item.GetDate(BibliographyDateRole.Issued);
                if (issued != null) {
                    Add(fields, "year", issued.Year?.ToString(CultureInfo.InvariantCulture) ?? issued.Literal);
                    Add(fields, "month", FormatClassicMonth(item, issued.Month));
                }
            }
            foreach (BibliographyIdentifier identifier in item.Identifiers) {
                cancellationToken.ThrowIfCancellationRequested();
                string fieldName = identifier.Scheme;
                if (CodecMappings.IsBibIdentifierScheme(identifier.Scheme) && IsSafeFieldName(fieldName) && !ReservedTypedFieldNames.Contains(fieldName)) Add(fields, fieldName, identifier.Value);
                else report.Add("BIBCONV129", BibliographyDiagnosticSeverity.Warning, $"Identifier scheme '{identifier.Scheme}' cannot be represented as a safe, non-conflicting BibTeX field.", BibliographyConversionAction.Omitted, item, "identifiers." + identifier.Scheme);
            }
            if (item.Keywords.Count > 0) Add(fields, "keywords", JoinValues(item.Keywords, ", ", cancellationToken, FormatBibListItem));
            if (item.Notes.Count > 0) Add(fields, "note", JoinValues(item.Notes, "; ", cancellationToken));

            var emitted = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (KeyValuePair<string, string> field in fields) { cancellationToken.ThrowIfCancellationRequested(); emitted.Add(field.Key); }
            foreach (BibliographyNativeField field in item.NativeFields) {
                cancellationToken.ThrowIfCancellationRequested();
                bool typedDuplicate = IsBibFamily(field.Format) && (ReservedTypedFieldNames.Contains(field.Name) || TypedIdentifierFieldNames.Contains(field.Name));
                bool canRemainNative = !typedDuplicate || CanRemainNativeBibField(field, emitted);
                if (IsBibFamily(field.Format) && IsSafeFieldName(field.Name) && IsFieldAllowedInTarget(field.Name, format) && canRemainNative) {
                    fields.Add(new KeyValuePair<string, string>(field.Name, field.Value));
                    emitted.Add(field.Name);
                    report.Add("BIBCONV011", BibliographyDiagnosticSeverity.Information, $"Preserved native field '{field.Name}'.", BibliographyConversionAction.PreservedExtension, item, field.Name);
                } else if (!IsBibFamily(field.Format)) {
                    report.Add("BIBCONV111", BibliographyDiagnosticSeverity.Warning, $"Native {field.Format} field '{field.Name}' cannot be represented safely in {format}.", BibliographyConversionAction.Omitted, item, field.Name);
                } else {
                    report.Add("BIBCONV119", BibliographyDiagnosticSeverity.Warning, $"Native {format} field '{field.Name}' conflicts with a typed field or has an unsafe name.", BibliographyConversionAction.Omitted, item, field.Name);
                }
            }

            var writableFields = new List<KeyValuePair<string, string>>(fields.Count);
            for (int index = 0; index < fields.Count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                KeyValuePair<string, string> field = fields[index];
                bool normalizesTerminalBackslash = HasOddTrailingBackslash(field.Value, cancellationToken);
                string escaped = Escape(field.Value, cancellationToken);
                if (normalizesTerminalBackslash)
                    report.Add("BIBCONV133", BibliographyDiagnosticSeverity.Warning, $"Bib field '{field.Key}' ends in an odd backslash run that must be normalized before the closing delimiter.", BibliographyConversionAction.Approximated, item, field.Key);
                if (!IsSafeDelimitedValue(escaped, cancellationToken)) {
                    report.Add("BIBCONV134", BibliographyDiagnosticSeverity.Warning, $"Bib field '{field.Key}' cannot be enclosed safely after escaping and was omitted.", BibliographyConversionAction.Omitted, item, field.Key);
                    continue;
                }
                writableFields.Add(new KeyValuePair<string, string>(field.Key, escaped));
            }
            for (int index = 0; index < writableFields.Count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                KeyValuePair<string, string> field = writableFields[index];
                builder.Append("  ").Append(field.Key).Append(" = {").Append(field.Value).Append('}');
                if (index + 1 < writableFields.Count) builder.Append(',');
                builder.Append(options.LineEnding);
            }
            builder.Append('}').Append(options.LineEnding);
            if (itemIndex + 1 < document.Items.Count) builder.Append(options.LineEnding);
        }
        return builder.ToString();
    }

    private static void Add(ICollection<KeyValuePair<string, string>> fields, string name, string? value) {
        if (value != null) fields.Add(new KeyValuePair<string, string>(name, value));
    }

    private static string JoinValues(IEnumerable<string> values, string separator, CancellationToken cancellationToken, Func<string, string>? transform = null) {
        var builder = new StringBuilder();
        foreach (string value in values) {
            cancellationToken.ThrowIfCancellationRequested();
            if (builder.Length > 0) builder.Append(separator);
            builder.Append(transform == null ? value : transform(value));
        }
        return builder.ToString();
    }

    private static bool CanRemainNativeBibField(BibliographyNativeField field, ISet<string> emitted) {
        string name = field.Name.ToLowerInvariant();
        switch (name) {
            case "title": case "series": case "edition": case "volume": case "abstract": case "url": return emitted.Contains(name);
            case "journal": case "journaltitle": case "booktitle": return ContainsAny(emitted, "journal", "journaltitle", "booktitle");
            case "publisher": case "institution": case "organization": return ContainsAny(emitted, "publisher", "institution", "organization");
            case "address": case "location": return ContainsAny(emitted, "address", "location");
            case "number": case "issue": return ContainsAny(emitted, "number", "issue");
            case "pages": case "eid": return ContainsAny(emitted, "pages", "eid");
            case "language": case "langid": return ContainsAny(emitted, "language", "langid");
            case "date": case "year": return ContainsAny(emitted, "date", "year");
            case "month": return ContainsAny(emitted, "date", "month") || !CodecMappings.ParseMonth(field.Value).HasValue;
            case "urldate": return false;
            case "author": case "editor": case "translator": return false;
            case "doi": case "isbn": case "issn": case "pmid": case "pmcid": return string.IsNullOrWhiteSpace(field.Value);
            case "keywords": case "note": return false;
            default: return false;
        }
    }

    private static bool ContainsAny(ISet<string> values, params string[] candidates) => candidates.Any(values.Contains);

    private static string? FormatClassicMonth(BibliographyItem item, int? month) {
        if (!month.HasValue) return null;
        if (item.BibMonthWasNumeric || month.Value < 1 || month.Value > 12) return month.Value.ToString(CultureInfo.InvariantCulture);
        return CultureInfo.InvariantCulture.DateTimeFormat.GetMonthName(month.Value);
    }

    private static string GetBibFieldName(BibliographyItem item, string property, string fallback, BibliographyFormat format) =>
        item.BibFieldNames.TryGetValue(property, out string? fieldName) && IsFieldAllowedInTarget(fieldName, format) ? fieldName : fallback;

    private static string DefaultContainerField(BibliographyItem item, BibliographyFormat format) =>
        item.Type == BibliographyItemType.Chapter || item.Type == BibliographyItemType.PaperConference ? "booktitle" : format == BibliographyFormat.BibLatex ? "journaltitle" : "journal";

    private static void AddNames(ICollection<KeyValuePair<string, string>> fields, string name, BibliographyItem item, BibliographyContributorRole role, CancellationToken cancellationToken) {
        var names = new List<string>();
        foreach (BibliographyContributor contributor in item.Contributors) {
            cancellationToken.ThrowIfCancellationRequested();
            if (contributor.Role != role) continue;
            string formatted = FormatBibName(contributor.Name);
            if (formatted.Length > 0) names.Add(formatted);
        }
        if (names.Count > 0) Add(fields, name, string.Join(" and ", names));
    }

    private static string FormatBibName(BibliographyName name) =>
        name.Literal == null ? FormatStructuredBibName(name) : "{" + name.Literal + "}";

    private static string FormatStructuredBibName(BibliographyName name) {
        string family = string.Join(" ", new[] { name.NonDroppingParticle, name.Family }.Where(static part => !string.IsNullOrWhiteSpace(part)));
        string given = string.Join(" ", new[] { name.Given, name.DroppingParticle }.Where(static part => !string.IsNullOrWhiteSpace(part)));
        if (!string.IsNullOrWhiteSpace(name.Suffix)) return family + ", " + name.Suffix + ", " + given;
        return family + ", " + given;
    }

    private static string FormatBibListItem(string value) =>
        string.IsNullOrWhiteSpace(value) || !string.Equals(value, value.Trim(), StringComparison.Ordinal) || value.IndexOf(',') >= 0 || value.IndexOf(';') >= 0 || value.Length >= 2 && value[0] == '{' && value[value.Length - 1] == '}' ? "{" + value + "}" : value;

    private static string Escape(string value, CancellationToken cancellationToken) {
        int depth = 0;
        for (int index = 0; index < value.Length; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (value[index] == '\\' && index + 1 < value.Length) { index++; continue; }
            if (value[index] == '{') depth++;
            else if (value[index] == '}') { if (depth == 0) return EscapeAllBraces(value, cancellationToken); depth--; }
        }
        cancellationToken.ThrowIfCancellationRequested();
        string escaped = depth == 0 ? value : EscapeAllBraces(value, cancellationToken);
        return HasOddTrailingBackslash(escaped, cancellationToken) ? escaped + "\\" : escaped;
    }
    private static string EscapeAllBraces(string value, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        var builder = new StringBuilder(value.Length);
        for (int index = 0; index < value.Length; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (value[index] == '{' || value[index] == '}') builder.Append('\\');
            builder.Append(value[index]);
        }
        cancellationToken.ThrowIfCancellationRequested();
        return builder.ToString();
    }
    internal static bool HasOddTrailingBackslash(string value, CancellationToken cancellationToken = default) {
        int count = 0;
        for (int index = value.Length - 1; index >= 0 && value[index] == '\\'; index--) {
            if ((count & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            count++;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return (count & 1) != 0;
    }
    private static bool IsSafeDelimitedValue(string value, CancellationToken cancellationToken) {
        int depth = 0;
        for (int index = 0; index < value.Length; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (value[index] == '\\') {
                if (++index >= value.Length) return false;
                continue;
            }
            if (value[index] == '{') depth++;
            else if (value[index] == '}' && --depth < 0) return false;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return depth == 0;
    }
    private static bool TryWriteNativeEntry(StringBuilder builder, BibliographyNativeEntry entry, string lineEnding, CancellationToken cancellationToken) {
        if (entry.Kind == "string" && IsSafeFieldName(entry.Name ?? string.Empty) && string.Equals(Escape(entry.Value, cancellationToken), entry.Value, StringComparison.Ordinal)) {
            builder.Append("@string{").Append(entry.Name).Append(" = {").Append(entry.Value).Append("}}").Append(lineEnding).Append(lineEnding);
            return true;
        }
        if (entry.Kind == "preamble" && string.Equals(Escape(entry.Value, cancellationToken), entry.Value, StringComparison.Ordinal)) {
            builder.Append("@preamble{{").Append(entry.Value).Append("}}").Append(lineEnding).Append(lineEnding);
            return true;
        }
        if (entry.Kind == "comment" && string.Equals(Escape(entry.Value, cancellationToken), entry.Value, StringComparison.Ordinal)) {
            builder.Append("@comment{").Append(entry.Value).Append('}').Append(lineEnding).Append(lineEnding);
            return true;
        }
        if (entry.Kind == "line-comment" && !ContainsLineBreak(entry.Value, cancellationToken)) {
            builder.Append('%').Append(entry.Value).Append(lineEnding);
            return true;
        }
        return false;
    }
    private static bool ContainsLineBreak(string value, CancellationToken cancellationToken) {
        for (int index = 0; index < value.Length; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (value[index] == '\r' || value[index] == '\n') return true;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return false;
    }
    internal static string SafeKey(string key, CancellationToken cancellationToken) {
        if (IsNullOrWhiteSpace(key, cancellationToken)) return "item";
        var characters = new char[key.Length];
        for (int index = 0; index < key.Length; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            characters[index] = IsSafeKeyCharacter(key[index]) ? key[index] : '_';
        }
        cancellationToken.ThrowIfCancellationRequested();
        return new string(characters);
    }
    internal static bool IsNullOrWhiteSpace(string value, CancellationToken cancellationToken) {
        if (value == null) return true;
        for (int index = 0; index < value.Length; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (!char.IsWhiteSpace(value[index])) return false;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return true;
    }
    internal static bool HasUnsafeKeyCharacter(string key, CancellationToken cancellationToken) {
        for (int index = 0; index < key.Length; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (!IsSafeKeyCharacter(key[index])) return true;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return false;
    }
    internal static bool IsSafeKeyCharacter(char character) => !char.IsWhiteSpace(character) && !char.IsControl(character) && "\\\"#%'(),={}".IndexOf(character) < 0;
    private static bool IsSafeFieldName(string name) => name.Length > 0 && name.All(character => char.IsLetterOrDigit(character) || character == '-' || character == '_' || character == ':');
    private static bool IsSafeTypeName(string? name) => !string.IsNullOrWhiteSpace(name) && name!.All(character => char.IsLetterOrDigit(character) || character == '-' || character == '_' || character == ':' || character == '.');
    private static bool IsBibFamily(BibliographyFormat format) => format == BibliographyFormat.BibTex || format == BibliographyFormat.BibLatex;
    private static bool IsFieldAllowedInTarget(string name, BibliographyFormat format) => format == BibliographyFormat.BibLatex || !BibLatexOnlyTypedFields.Contains(name);

    internal static bool CanPreserveNativeType(BibliographyFormat sourceFormat, BibliographyFormat targetFormat, BibliographyItem item) {
        if (!IsBibFamily(sourceFormat) || !IsBibFamily(targetFormat) || !IsSafeTypeName(item.NativeType) || CodecMappings.ParseType(item.NativeType) != item.Type) return false;
        return sourceFormat == targetFormat || targetFormat == BibliographyFormat.BibLatex || ClassicBibTypes.Contains(item.NativeType!);
    }

    internal static bool CanRoundTripType(BibliographyItemType type, BibliographyFormat targetFormat) =>
        type == BibliographyItemType.ArticleJournal || type == BibliographyItemType.Book || type == BibliographyItemType.Chapter ||
        type == BibliographyItemType.PaperConference || type == BibliographyItemType.Proceedings || type == BibliographyItemType.Report ||
        type == BibliographyItemType.Manuscript || targetFormat == BibliographyFormat.BibLatex && type == BibliographyItemType.Thesis;

    private static string OutputType(BibliographyItemType type, BibliographyFormat format) =>
        format == BibliographyFormat.BibLatex && type == BibliographyItemType.Thesis ? "thesis" : CodecMappings.ToBibType(type);

    internal static bool CanRoundTripStructuredName(BibliographyName name, CancellationToken cancellationToken) {
        if (name.Literal != null) return name.Given == null && name.Family == null && name.Suffix == null && name.NonDroppingParticle == null && name.DroppingParticle == null;
        foreach (string? value in new[] { name.Given, name.Family, name.Suffix, name.NonDroppingParticle, name.DroppingParticle })
            if (ContainsBibNameSyntaxSeparator(value, cancellationToken) || HasNormalizedBibNameWhitespace(value, cancellationToken)) return false;
        if (!IsLowercaseParticle(name.NonDroppingParticle) || !IsLowercaseParticle(name.DroppingParticle)) return false;
        string family = string.Join(" ", new[] { name.NonDroppingParticle, name.Family }.Where(static part => !string.IsNullOrWhiteSpace(part)));
        string given = string.Join(" ", new[] { name.Given, name.DroppingParticle }.Where(static part => !string.IsNullOrWhiteSpace(part)));
        return CountLeadingBibParticleWords(family) == CountWords(name.NonDroppingParticle) && CountTrailingLowercaseWords(given) == CountWords(name.DroppingParticle);
    }

    private static bool ContainsBibNameSyntaxSeparator(string? value, CancellationToken cancellationToken) {
        if (value == null) return false;
        for (int index = 0; index < value.Length; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (value[index] == ',') return true;
            if (char.IsWhiteSpace(value[index]) && TryGetNameSeparatorEnd(value, index, cancellationToken, out _)) return true;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return false;
    }
    private static bool HasNormalizedBibNameWhitespace(string? value, CancellationToken cancellationToken) {
        if (value == null || value.Length == 0) return false;
        if (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[value.Length - 1])) return true;
        for (int index = 1; index < value.Length; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (value[index - 1] == ' ' && value[index] == ' ') return true;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return false;
    }

    private static bool TryGetNameSeparatorEnd(string value, int index, CancellationToken cancellationToken, out int separatorEnd) {
        separatorEnd = index;
        if (index >= value.Length || !char.IsWhiteSpace(value[index])) return false;
        int wordStart = index;
        while (wordStart < value.Length && char.IsWhiteSpace(value[wordStart])) {
            if (((wordStart - index) & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            wordStart++;
        }
        if (wordStart + 3 >= value.Length ||
            (value[wordStart] != 'a' && value[wordStart] != 'A') ||
            (value[wordStart + 1] != 'n' && value[wordStart + 1] != 'N') ||
            (value[wordStart + 2] != 'd' && value[wordStart + 2] != 'D') ||
            !char.IsWhiteSpace(value[wordStart + 3])) return false;
        separatorEnd = wordStart + 4;
        while (separatorEnd < value.Length && char.IsWhiteSpace(value[separatorEnd])) {
            if (((separatorEnd - wordStart) & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            separatorEnd++;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return true;
    }

    private static bool IsLowercaseParticle(string? value) => string.IsNullOrWhiteSpace(value) || value!.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).All(static word => StartsWithLowercaseLetter(word));
    private static int CountWords(string? value) => string.IsNullOrWhiteSpace(value) ? 0 : value!.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Length;
    private static int CountLeadingBibParticleWords(string value) { string[] words = value.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries); int count = 0; while (count < words.Length - 1 && StartsWithLowercaseLetter(words[count])) count++; return count; }
    private static int CountTrailingLowercaseWords(string value) { string[] words = value.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries); int start = words.Length; while (start > 0 && StartsWithLowercaseLetter(words[start - 1])) start--; return words.Length - start; }
    private static bool StartsWithLowercaseLetter(string value) { char first = value.FirstOrDefault(char.IsLetter); return first != default(char) && char.IsLower(first); }

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
                    while (_position < _source.Length && _source[_position] != '@') { CheckScanCancellation(); _position++; }
                    AddDiagnostic("BIBBIB001", "Ignored text outside a BibTeX entry.", invalidStart);
                    continue;
                }
                int entryStart = _position++;
                string type = ReadIdentifier();
                _limits.CheckValueLength(_items, type, entryStart);
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
            _limits.CheckValueLength(_items, name, entryStart);
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
            var keyBuilder = new StringBuilder();
            while (_position < _source.Length && _source[_position] != ',' && _source[_position] != close) { CheckScanCancellation(); AppendValue(keyBuilder, _source[_position++], keyStart); }
            string key = keyBuilder.ToString().Trim();
            _limits.AddValue(_items, key, keyStart);
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
                _limits.CheckValueLength(_items, name, fieldStart);
                if (name.Length == 0) { AddDiagnostic("BIBBIB005", "Expected a BibTeX field name.", fieldStart, item.Key, severity: BibliographyDiagnosticSeverity.Error); RecoverField(close); continue; }
                SkipWhitespace();
                if (!Consume('=')) { AddDiagnostic("BIBBIB006", "Expected '=' after a BibTeX field name.", _position, item.Key, name, BibliographyDiagnosticSeverity.Error); RecoverField(close); continue; }
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
                if (_source[_position] == '{') AppendValue(builder, ReadDelimited('{', '}'), _position);
                else if (_source[_position] == '"') AppendValue(builder, ReadDelimited('"', '"'), _position);
                else {
                    string atom = ReadValueAtom(entryClose);
                    if (_strings.TryGetValue(atom, out string? expanded)) {
                        _limits.AddExpandedCharacters(_items, expanded.Length, _position - atom.Length);
                        AppendValue(builder, expanded, _position - atom.Length);
                    } else AppendValue(builder, atom, _position - atom.Length);
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
            int quotedBraceDepth = 0;
            while (_position < _source.Length) {
                _cancellationToken.ThrowIfCancellationRequested();
                char current = _source[_position++];
                if (current == '\\' && _position < _source.Length) { AppendValue(builder, current, _position - 1); AppendValue(builder, _source[_position++], _position - 1); continue; }
                if (open == '"') {
                    if (current == '{') { quotedBraceDepth++; _limits.CheckDepth(_items, quotedBraceDepth + 1, _position - 1); AppendValue(builder, current, _position - 1); continue; }
                    if (current == '}' && quotedBraceDepth > 0) { quotedBraceDepth--; AppendValue(builder, current, _position - 1); continue; }
                    if (current == close && quotedBraceDepth == 0) return builder.ToString();
                    AppendValue(builder, current, _position - 1);
                    continue;
                }
                if (open != '"' && current == open) { depth++; _limits.CheckDepth(_items, depth, _position - 1); if (depth > 1) AppendValue(builder, current, _position - 1); continue; }
                if (current == close) { depth--; if (depth == 0) return builder.ToString(); AppendValue(builder, current, _position - 1); continue; }
                AppendValue(builder, current, _position - 1);
            }
            AddDiagnostic("BIBBIB007", "Delimited BibTeX value was not closed.", start, severity: BibliographyDiagnosticSeverity.Error);
            return builder.ToString();
        }

        private string ReadBalancedRaw(char close) {
            int start = _position;
            var builder = new StringBuilder();
            int depth = 1;
            while (_position < _source.Length) {
                _cancellationToken.ThrowIfCancellationRequested();
                char current = _source[_position++];
                if (current == '\\' && _position < _source.Length) { AppendValue(builder, current, _position - 1); AppendValue(builder, _source[_position++], _position - 1); continue; }
                if (current == close) { depth--; if (depth == 0) return builder.ToString(); }
                else if (current == (close == '}' ? '{' : '(')) { depth++; _limits.CheckDepth(_items, depth, _position - 1); }
                AppendValue(builder, current, _position - 1);
            }
            AddDiagnostic("BIBBIB008", "BibTeX directive was not closed.", start, severity: BibliographyDiagnosticSeverity.Error);
            return builder.ToString();
        }

        private string ReadValueAtom(char close) {
            int start = _position;
            var builder = new StringBuilder();
            while (_position < _source.Length && _source[_position] != '#' && _source[_position] != ',' && _source[_position] != close && !char.IsWhiteSpace(_source[_position])) { CheckScanCancellation(); AppendValue(builder, _source[_position++], start); }
            return builder.ToString();
        }

        private void Bind(BibliographyItem item, string name, string value) {
            string field = name.ToLowerInvariant();
            switch (field) {
                case "title": SetScalar(item, field, value, () => item.Title, assigned => item.Title = assigned); break;
                case "journal": case "journaltitle": if (item.ContainerTitle == null) { item.ContainerTitle = value; item.BibFieldNames["container-title"] = field; } else PreserveAdditionalField(item, field, value); break;
                case "booktitle": if (item.ContainerTitle == null) { item.ContainerTitle = value; item.BibFieldNames["container-title"] = field; } else PreserveAdditionalField(item, field, value); break;
                case "series": SetScalar(item, field, value, () => item.CollectionTitle, assigned => item.CollectionTitle = assigned); break;
                case "publisher": case "institution": case "organization": if (item.Publisher == null) { item.Publisher = value; item.BibFieldNames["publisher"] = field; } else PreserveAdditionalField(item, field, value); break;
                case "address": case "location": if (item.PublisherPlace == null) { item.PublisherPlace = value; item.BibFieldNames["publisher-place"] = field; } else PreserveAdditionalField(item, field, value); break;
                case "edition": SetScalar(item, field, value, () => item.Edition, assigned => item.Edition = assigned); break;
                case "volume": SetScalar(item, field, value, () => item.Volume, assigned => item.Volume = assigned); break;
                case "number": case "issue": if (item.Issue == null) { item.Issue = value; item.BibFieldNames["issue"] = field; } else PreserveAdditionalField(item, field, value); break;
                case "pages": case "eid": if (item.Pages == null) { item.Pages = value; item.BibFieldNames["pages"] = field; } else PreserveAdditionalField(item, field, value); break;
                case "abstract": SetScalar(item, field, value, () => item.Abstract, assigned => item.Abstract = assigned); break;
                case "language": case "langid": if (item.Language == null) { item.Language = value; item.BibFieldNames["language"] = field; } else PreserveAdditionalField(item, field, value); break;
                case "url": SetScalar(item, field, value, () => item.Url, assigned => item.Url = assigned); break;
                case "author": AddNames(item, BibliographyContributorRole.Author, name, value); break;
                case "editor": AddNames(item, BibliographyContributorRole.Editor, name, value); break;
                case "translator": AddNames(item, BibliographyContributorRole.Translator, name, value); break;
                case "date": SetIssuedDate(item, value); break;
                case "year": SetYear(item, value); break;
                case "month": SetMonth(item, value); break;
                case "urldate": item.Dates.Add(CodecMappings.ParseDate(BibliographyDateRole.Accessed, value)); break;
                case "doi": case "isbn": case "issn": case "pmid": case "pmcid": AddIdentifier(item, name, value); break;
                case "keywords": AddKeywords(item, value); break;
                case "note": item.Notes.Add(value); break;
                default: item.NativeFields.Add(new BibliographyNativeField(_format, name, value)); break;
            }
        }

        private void AddKeywords(BibliographyItem item, string value) {
            foreach (string keyword in SplitBibList(value)) {
                string parsedKeyword = UnwrapBibListItem(keyword);
                _limits.AddValue(_items, parsedKeyword, _position);
                item.Keywords.Add(parsedKeyword);
            }
        }

        private void AddNames(BibliographyItem item, BibliographyContributorRole role, string fieldName, string value) {
            bool hasSurplusSegments = false;
            foreach (string part in SplitNames(value)) {
                _limits.AddValue(_items, part, _position);
                item.Contributors.Add(new BibliographyContributor(role, ParseBibName(part, out bool contributorHasSurplusSegments)));
                hasSurplusSegments |= contributorHasSurplusSegments;
            }
            if (!hasSurplusSegments) return;
            PreserveAdditionalField(item, fieldName, value);
            AddDiagnostic("BIBBIB011", "A BibTeX contributor contains surplus top-level comma segments; the complete field was retained as native source data.", _position, item.Key, fieldName, BibliographyDiagnosticSeverity.Error);
        }

        private void PreserveAdditionalField(BibliographyItem item, string fieldName, string value) => item.NativeFields.Add(new BibliographyNativeField(_format, fieldName, value));

        private void AddIdentifier(BibliographyItem item, string fieldName, string value) {
            if (string.IsNullOrWhiteSpace(value)) PreserveAdditionalField(item, fieldName, value);
            else item.Identifiers.Add(new BibliographyIdentifier(fieldName, value));
        }

        private void SetScalar(BibliographyItem item, string fieldName, string value, Func<string?> read, Action<string> write) {
            if (read() == null) write(value);
            else PreserveAdditionalField(item, fieldName, value);
        }

        private void AppendValue(StringBuilder builder, string value, int offset) { _limits.CheckAdditionalValueLength(_items, builder.Length, value.Length, offset); builder.Append(value); }
        private void AppendValue(StringBuilder builder, char value, int offset) { _limits.CheckAdditionalValueLength(_items, builder.Length, 1, offset); builder.Append(value); }

        private BibliographyName ParseBibName(string value, out bool hasSurplusSegments) {
            string trimmed = value.Trim();
            hasSurplusSegments = false;
            if (trimmed.Length >= 2 && trimmed[0] == '{' && trimmed[trimmed.Length - 1] == '}') return new BibliographyName { Literal = trimmed.Substring(1, trimmed.Length - 2) };
            string[] parts = SplitTopLevel(trimmed, ',').Take(4).ToArray();
            hasSurplusSegments = parts.Length > 3;
            if (parts.Length == 1) return ParseBibFirstVonLast(trimmed);
            SplitBibFamily(parts[0], out string? particle, out string? family);
            SplitBibGiven(parts.Length == 3 ? parts[2] : parts[1], out string? given, out string? droppingParticle);
            return new BibliographyName { Family = family, NonDroppingParticle = particle, Suffix = parts.Length == 3 ? NullIfEmpty(parts[1]) : null, Given = given, DroppingParticle = droppingParticle };
        }

        private static void SplitBibFamily(string value, out string? particle, out string? family) {
            string[] words = value.Trim().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            int particleCount = 0;
            while (particleCount < words.Length - 1 && StartsWithLowercaseLetter(words[particleCount])) particleCount++;
            particle = NullIfEmpty(string.Join(" ", words.Take(particleCount)));
            family = NullIfEmpty(string.Join(" ", words.Skip(particleCount)));
        }

        private static BibliographyName ParseBibFirstVonLast(string value) {
            string[] words = value.Trim().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (words.Length == 0) return new BibliographyName();
            if (words.Length == 1) return new BibliographyName { Family = words[0] };
            int particleStart = Array.FindIndex(words, StartsWithLowercaseLetter);
            if (particleStart < 0) return new BibliographyName { Given = string.Join(" ", words.Take(words.Length - 1)), Family = words[words.Length - 1] };
            int familyStart = particleStart + 1;
            while (familyStart < words.Length - 1 && StartsWithLowercaseLetter(words[familyStart])) familyStart++;
            return new BibliographyName {
                Given = NullIfEmpty(string.Join(" ", words.Take(particleStart))),
                NonDroppingParticle = NullIfEmpty(string.Join(" ", words.Skip(particleStart).Take(familyStart - particleStart))),
                Family = NullIfEmpty(string.Join(" ", words.Skip(familyStart)))
            };
        }

        private static void SplitBibGiven(string value, out string? given, out string? droppingParticle) {
            string[] words = value.Trim().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            int particleStart = words.Length;
            while (particleStart > 0 && StartsWithLowercaseLetter(words[particleStart - 1])) particleStart--;
            given = NullIfEmpty(string.Join(" ", words.Take(particleStart)));
            droppingParticle = NullIfEmpty(string.Join(" ", words.Skip(particleStart)));
        }

        private static bool StartsWithLowercaseLetter(string value) => BibCodec.StartsWithLowercaseLetter(value);
        private static string? NullIfEmpty(string value) => string.IsNullOrWhiteSpace(value) ? null : value.Trim();

        private IEnumerable<string> SplitBibList(string value) => SplitTopLevel(value, ',', ';').Select(static part => part.Trim());

        private static string UnwrapBibListItem(string value) {
            string trimmed = value.Trim();
            return trimmed.Length >= 2 && trimmed[0] == '{' && trimmed[trimmed.Length - 1] == '}' ? trimmed.Substring(1, trimmed.Length - 2) : trimmed;
        }

        private IEnumerable<string> SplitTopLevel(string value, params char[] separators) {
            int start = 0;
            int depth = 0;
            for (int index = 0; index < value.Length; index++) {
                if ((index & 4095) == 0) _cancellationToken.ThrowIfCancellationRequested();
                if (value[index] == '\\' && index + 1 < value.Length) { index++; continue; }
                if (value[index] == '{') depth++;
                else if (value[index] == '}' && depth > 0) depth--;
                else if (depth == 0 && separators.Contains(value[index])) { yield return value.Substring(start, index - start); start = index + 1; }
            }
            yield return value.Substring(start);
        }

        private IEnumerable<string> SplitNames(string value) {
            int start = 0;
            int depth = 0;
            for (int index = 0; index <= value.Length - 5; index++) {
                if ((index & 4095) == 0) _cancellationToken.ThrowIfCancellationRequested();
                if (value[index] == '\\' && index + 1 < value.Length) { index++; continue; }
                if (value[index] == '{') depth++; else if (value[index] == '}' && depth > 0) depth--;
                if (depth == 0 && TryGetNameSeparatorEnd(value, index, _cancellationToken, out int separatorEnd)) { yield return value.Substring(start, index - start).Trim(); start = separatorEnd; index = separatorEnd - 1; }
            }
            if (start <= value.Length) yield return value.Substring(start).Trim();
        }

        private void SetYear(BibliographyItem item, string value) {
            BibliographyDate date = item.GetDate(BibliographyDateRole.Issued) ?? new BibliographyDate { Role = BibliographyDateRole.Issued };
            if (!item.Dates.Contains(date)) item.Dates.Add(date);
            if (item.BibFieldNames.ContainsKey("issued-year")) { PreserveAdditionalField(item, "year", value); return; }
            item.BibFieldNames["issued-year"] = "year";
            if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int year)) date.Year = year; else date.Literal = value;
        }

        private void SetIssuedDate(BibliographyItem item, string value) {
            if (item.GetDate(BibliographyDateRole.Issued) != null) { PreserveAdditionalField(item, "date", value); return; }
            item.Dates.Add(CodecMappings.ParseDate(BibliographyDateRole.Issued, value));
            item.BibFieldNames["issued-date"] = "date";
            item.BibFieldNames["issued-year"] = "date";
            item.BibFieldNames["issued-month"] = "date";
        }

        private void SetMonth(BibliographyItem item, string value) {
            if (item.BibFieldNames.ContainsKey("issued-month")) { PreserveAdditionalField(item, "month", value); return; }
            int? month = CodecMappings.ParseMonth(value);
            if (!month.HasValue) { PreserveAdditionalField(item, "month", value); return; }
            BibliographyDate date = item.GetDate(BibliographyDateRole.Issued) ?? new BibliographyDate { Role = BibliographyDateRole.Issued };
            if (!item.Dates.Contains(date)) item.Dates.Add(date);
            item.BibFieldNames["issued-month"] = "month";
            item.BibMonthWasNumeric = int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out _);
            date.Month = month;
        }

        private string ReadIdentifier() {
            SkipWhitespace();
            int start = _position;
            while (_position < _source.Length && (char.IsLetterOrDigit(_source[_position]) || _source[_position] == '-' || _source[_position] == '_' || _source[_position] == ':' || _source[_position] == '.')) { CheckScanCancellation(); _position++; }
            return _source.Substring(start, _position - start);
        }

        private bool Consume(char value) { SkipWhitespace(); if (_position < _source.Length && _source[_position] == value) { _position++; return true; } return false; }
        private void ConsumeClose(char close) { SkipWhitespaceAndCommas(); if (!Consume(close)) AddDiagnostic("BIBBIB009", $"Expected closing '{close}'.", _position, severity: BibliographyDiagnosticSeverity.Error); }
        private void SkipWhitespace() { while (_position < _source.Length && char.IsWhiteSpace(_source[_position])) { CheckScanCancellation(); _position++; } }
        private void SkipWhitespaceAndCommas() { while (_position < _source.Length && (char.IsWhiteSpace(_source[_position]) || _source[_position] == ',')) { CheckScanCancellation(); _position++; } }
        private void SkipTrivia() { while (_position < _source.Length) { CheckScanCancellation(); if (char.IsWhiteSpace(_source[_position]) || _source[_position] == '\uFEFF') { _position++; continue; } if (_source[_position] == '%') { int start = ++_position; var builder = new StringBuilder(); while (_position < _source.Length && _source[_position] != '\n' && _source[_position] != '\r') { CheckScanCancellation(); AppendValue(builder, _source[_position++], start); } string value = builder.ToString(); _limits.AddValue(_items, value, start); _nativeEntries.Add(new BibliographyNativeEntry(_format, "line-comment", value)); continue; } break; } }
        private void RecoverToNextEntry() { while (_position < _source.Length && _source[_position] != '@') { CheckScanCancellation(); _position++; } }
        private void RecoverField(char close) {
            int braceDepth = 0;
            bool quoted = false;
            bool escaped = false;
            while (_position < _source.Length) {
                CheckScanCancellation();
                char current = _source[_position];
                if (escaped) { escaped = false; _position++; continue; }
                if (current == '\\') { escaped = true; _position++; continue; }
                if (current == '"' && braceDepth == 0) { quoted = !quoted; _position++; continue; }
                if (!quoted) {
                    if (current == '{') { braceDepth++; _position++; continue; }
                    if (current == '}' && braceDepth > 0) { braceDepth--; _position++; continue; }
                    if (braceDepth == 0 && (current == ',' || current == close)) return;
                }
                _position++;
            }
        }
        private void CheckScanCancellation() { if ((_position & 4095) == 0) _cancellationToken.ThrowIfCancellationRequested(); }
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
                if ((_locationOffset & 4095) == 0) _cancellationToken.ThrowIfCancellationRequested();
                char current = _source[_locationOffset++];
                if (current == '\r') {
                    if (_locationOffset < _source.Length && _source[_locationOffset] == '\n') _locationOffset++;
                    _locationLine++;
                    _locationColumn = 1;
                } else if (current == '\n') {
                    _locationLine++;
                    _locationColumn = 1;
                } else _locationColumn++;
            }
            line = _locationLine; column = _locationColumn;
        }
    }
}
