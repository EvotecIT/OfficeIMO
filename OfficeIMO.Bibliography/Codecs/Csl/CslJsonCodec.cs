using System.Text.Json;

namespace OfficeIMO.Bibliography;

internal static class CslJsonCodec {
    private static readonly HashSet<string> KnownProperties = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "id", "type", "title", "container-title", "collection-title", "publisher", "publisher-place", "edition", "volume", "issue", "page", "abstract", "language", "URL",
        "author", "editor", "translator", "recipient", "interviewer", "composer", "collection-editor", "issued", "accessed", "submitted", "original-date", "event-date",
        "DOI", "ISBN", "ISSN", "PMID", "PMCID", "keyword", "note"
    };

    internal static IList<BibliographyItem> Parse(string source, BibliographyReadOptions options, List<BibliographyDiagnostic> diagnostics, CancellationToken cancellationToken) {
        var items = new List<BibliographyItem>();
        var limits = new BibliographyLimitGuard(options);
        var diagnosticGuard = new BibliographyDiagnosticGuard(options, diagnostics, items);
        try {
            using JsonDocument json = JsonDocument.Parse(source, new JsonDocumentOptions { AllowTrailingCommas = true, CommentHandling = JsonCommentHandling.Skip, MaxDepth = options.MaximumNestingDepth });
            if (json.RootElement.ValueKind == JsonValueKind.Array) {
                foreach (JsonElement element in json.RootElement.EnumerateArray()) ParseItem(element, items, limits, diagnosticGuard, cancellationToken);
            } else if (json.RootElement.ValueKind == JsonValueKind.Object) {
                ParseItem(json.RootElement, items, limits, diagnosticGuard, cancellationToken);
            } else {
                diagnosticGuard.Add(new BibliographyDiagnostic("BIBCSL001", BibliographyDiagnosticSeverity.Error, "CSL JSON root must be an object or an array."));
            }
        } catch (JsonException exception) {
            GetJsonLocation(source, exception, out int offset, out int line, out int column);
            diagnosticGuard.Add(new BibliographyDiagnostic("BIBCSL002", BibliographyDiagnosticSeverity.Error, exception.Message, offset, line, column));
        }
        return items;
    }

    internal static string Write(BibliographyDocument document, BibliographyWriteOptions options, BibliographyConversionReport report, CancellationToken cancellationToken) {
        using var stream = new MemoryStream();
        using (var writer = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = true })) {
            writer.WriteStartArray();
            for (int itemIndex = 0; itemIndex < document.Items.Count; itemIndex++) {
                BibliographyItem item = document.Items[itemIndex];
                cancellationToken.ThrowIfCancellationRequested();
                writer.WriteStartObject();
                writer.WriteString("id", CodecMappings.OutputKey(item, itemIndex));
                writer.WriteString("type", item.Type == BibliographyItemType.Unknown && !string.IsNullOrWhiteSpace(item.NativeType) ? item.NativeType : CodecMappings.ToCslType(item.Type));
                WriteString(writer, "title", item.Title); WriteString(writer, "container-title", item.ContainerTitle); WriteString(writer, "collection-title", item.CollectionTitle);
                WriteString(writer, "publisher", item.Publisher); WriteString(writer, "publisher-place", item.PublisherPlace); WriteString(writer, "edition", item.Edition);
                WriteString(writer, "volume", item.Volume); WriteString(writer, "issue", item.Issue); WriteString(writer, "page", item.Pages); WriteString(writer, "abstract", item.Abstract);
                WriteString(writer, "language", item.Language); WriteString(writer, "URL", item.Url);
                WriteNames(writer, item, BibliographyContributorRole.Author, "author", report); WriteNames(writer, item, BibliographyContributorRole.Editor, "editor", report);
                WriteNames(writer, item, BibliographyContributorRole.Translator, "translator", report); WriteNames(writer, item, BibliographyContributorRole.Recipient, "recipient", report);
                WriteNames(writer, item, BibliographyContributorRole.Interviewer, "interviewer", report); WriteNames(writer, item, BibliographyContributorRole.Composer, "composer", report);
                WriteNames(writer, item, BibliographyContributorRole.CollectionEditor, "collection-editor", report);
                WriteDate(writer, item, BibliographyDateRole.Issued, "issued", report); WriteDate(writer, item, BibliographyDateRole.Accessed, "accessed", report);
                WriteDate(writer, item, BibliographyDateRole.Submitted, "submitted", report); WriteDate(writer, item, BibliographyDateRole.Original, "original-date", report); WriteDate(writer, item, BibliographyDateRole.Event, "event-date", report);
                foreach (IGrouping<string, BibliographyIdentifier> group in item.Identifiers.Where(identifier => CodecMappings.IsCslIdentifierScheme(identifier.Scheme)).GroupBy(identifier => identifier.Scheme.ToUpperInvariant(), StringComparer.OrdinalIgnoreCase)) WriteString(writer, group.Key, string.Join("; ", group.Select(static identifier => identifier.Value)));
                if (item.Keywords.Count > 0) writer.WriteString("keyword", string.Join(", ", item.Keywords));
                if (item.Notes.Count > 0) writer.WriteString("note", string.Join("; ", item.Notes));

                var emitted = new HashSet<string>(KnownProperties, StringComparer.OrdinalIgnoreCase);
                foreach (BibliographyNativeField field in item.NativeFields) {
                    if (field.Format == BibliographyFormat.CslJson && !emitted.Contains(field.Name)) {
                        writer.WritePropertyName(field.Name);
                        bool exact = WriteNativeValue(writer, field);
                        emitted.Add(field.Name);
                        report.Add(exact ? "BIBCONV012" : "BIBCONV126", exact ? BibliographyDiagnosticSeverity.Information : BibliographyDiagnosticSeverity.Warning, exact ? $"Preserved native CSL JSON property '{field.Name}'." : $"Native CSL JSON property '{field.Name}' was emitted as a string because its raw JSON value was invalid or too deeply nested.", exact ? BibliographyConversionAction.PreservedExtension : BibliographyConversionAction.Approximated, item, field.Name);
                    } else if (field.Format != BibliographyFormat.CslJson) {
                        report.Add("BIBCONV112", BibliographyDiagnosticSeverity.Warning, $"Native {field.Format} field '{field.Name}' cannot be represented safely in CSL JSON.", BibliographyConversionAction.Omitted, item, field.Name);
                    } else {
                        report.Add("BIBCONV120", BibliographyDiagnosticSeverity.Warning, $"Native CSL JSON property '{field.Name}' conflicts with a typed property.", BibliographyConversionAction.Omitted, item, field.Name);
                    }
                }
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
        }
        string text = Encoding.UTF8.GetString(stream.ToArray());
        foreach (BibliographyNativeEntry entry in document.NativeEntries) report.Add("BIBCONV121", BibliographyDiagnosticSeverity.Warning, $"Document-level {entry.Format} entry '{entry.Kind}' cannot be represented in CSL JSON.", BibliographyConversionAction.Omitted, field: entry.Name ?? entry.Kind);
        return options.LineEnding == "\n" ? text + options.LineEnding : NormalizeLineEndings(text, options.LineEnding) + options.LineEnding;
    }

    private static void ParseItem(JsonElement element, IList<BibliographyItem> items, BibliographyLimitGuard limits, BibliographyDiagnosticGuard diagnostics, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (element.ValueKind != JsonValueKind.Object) { diagnostics.Add(new BibliographyDiagnostic("BIBCSL003", BibliographyDiagnosticSeverity.Warning, "Ignored a non-object CSL JSON array value.")); return; }
        limits.AddItem(items, 0);
        var item = new BibliographyItem();
        items.Add(item);
        CountValues(element, items, limits);
        var seenProperties = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (JsonProperty property in element.EnumerateObject()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!seenProperties.Add(property.Name)) {
                string duplicateRaw = GetBoundedRawValue(property.Value, items, limits);
                item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.CslJson, property.Name, ScalarOrRaw(property.Value, duplicateRaw), duplicateRaw));
                continue;
            }
            switch (property.Name.ToLowerInvariant()) {
                case "id": item.Key = Scalar(property.Value); break;
                case "type": item.NativeType = Scalar(property.Value); item.Type = CodecMappings.ParseType(item.NativeType); break;
                case "title": item.Title = Scalar(property.Value); break;
                case "container-title": item.ContainerTitle = Scalar(property.Value); break;
                case "collection-title": item.CollectionTitle = Scalar(property.Value); break;
                case "publisher": item.Publisher = Scalar(property.Value); break;
                case "publisher-place": item.PublisherPlace = Scalar(property.Value); break;
                case "edition": item.Edition = Scalar(property.Value); break;
                case "volume": item.Volume = Scalar(property.Value); break;
                case "issue": item.Issue = Scalar(property.Value); break;
                case "page": item.Pages = Scalar(property.Value); break;
                case "abstract": item.Abstract = Scalar(property.Value); break;
                case "language": item.Language = Scalar(property.Value); break;
                case "url": item.Url = Scalar(property.Value); break;
                case "author": ParseNames(item, property.Value, BibliographyContributorRole.Author, items, limits); break;
                case "editor": ParseNames(item, property.Value, BibliographyContributorRole.Editor, items, limits); break;
                case "translator": ParseNames(item, property.Value, BibliographyContributorRole.Translator, items, limits); break;
                case "recipient": ParseNames(item, property.Value, BibliographyContributorRole.Recipient, items, limits); break;
                case "interviewer": ParseNames(item, property.Value, BibliographyContributorRole.Interviewer, items, limits); break;
                case "composer": ParseNames(item, property.Value, BibliographyContributorRole.Composer, items, limits); break;
                case "collection-editor": ParseNames(item, property.Value, BibliographyContributorRole.CollectionEditor, items, limits); break;
                case "issued": ParseDate(item, property.Value, BibliographyDateRole.Issued, diagnostics, items, limits); break;
                case "accessed": ParseDate(item, property.Value, BibliographyDateRole.Accessed, diagnostics, items, limits); break;
                case "submitted": ParseDate(item, property.Value, BibliographyDateRole.Submitted, diagnostics, items, limits); break;
                case "original-date": ParseDate(item, property.Value, BibliographyDateRole.Original, diagnostics, items, limits); break;
                case "event-date": ParseDate(item, property.Value, BibliographyDateRole.Event, diagnostics, items, limits); break;
                case "doi": CodecMappings.AddIdentifier(item, "DOI", Scalar(property.Value)); break;
                case "isbn": CodecMappings.AddIdentifier(item, "ISBN", Scalar(property.Value)); break;
                case "issn": CodecMappings.AddIdentifier(item, "ISSN", Scalar(property.Value)); break;
                case "pmid": CodecMappings.AddIdentifier(item, "PMID", Scalar(property.Value)); break;
                case "pmcid": CodecMappings.AddIdentifier(item, "PMCID", Scalar(property.Value)); break;
                case "keyword": string keyword = Scalar(property.Value); if (!string.IsNullOrWhiteSpace(keyword)) item.Keywords.Add(keyword); break;
                case "note": item.Notes.Add(Scalar(property.Value)); break;
                default:
                    string raw = GetBoundedRawValue(property.Value, items, limits);
                    item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.CslJson, property.Name, ScalarOrRaw(property.Value, raw), raw));
                    break;
            }
        }
        if (string.IsNullOrWhiteSpace(item.Key)) diagnostics.Add(new BibliographyDiagnostic("BIBCSL004", BibliographyDiagnosticSeverity.Warning, "CSL JSON item has no id."));
    }

    private static void CountValues(JsonElement element, IList<BibliographyItem> items, BibliographyLimitGuard limits) {
        switch (element.ValueKind) {
            case JsonValueKind.Object:
                foreach (JsonProperty property in element.EnumerateObject()) {
                    limits.AddValue(items, property.Name, 0);
                    if (property.Value.ValueKind == JsonValueKind.Object || property.Value.ValueKind == JsonValueKind.Array) CountValues(property.Value, items, limits);
                    else limits.CheckValueLength(items, property.Value.GetRawText(), 0);
                }
                break;
            case JsonValueKind.Array:
                foreach (JsonElement value in element.EnumerateArray()) {
                    limits.AddValue(items, null, 0);
                    if (value.ValueKind == JsonValueKind.Object || value.ValueKind == JsonValueKind.Array) CountValues(value, items, limits);
                    else limits.CheckValueLength(items, value.GetRawText(), 0);
                }
                break;
        }
    }

    private static string Scalar(JsonElement value) {
        switch (value.ValueKind) {
            case JsonValueKind.String: return value.GetString() ?? string.Empty;
            case JsonValueKind.Number: case JsonValueKind.True: case JsonValueKind.False: return value.GetRawText();
            case JsonValueKind.Null: case JsonValueKind.Undefined: return string.Empty;
            default: return value.GetRawText();
        }
    }

    private static string ScalarOrRaw(JsonElement value, string raw) => value.ValueKind == JsonValueKind.Object || value.ValueKind == JsonValueKind.Array ? raw : Scalar(value);

    private static void ParseNames(BibliographyItem item, JsonElement value, BibliographyContributorRole role, IList<BibliographyItem> items, BibliographyLimitGuard limits) {
        if (value.ValueKind != JsonValueKind.Array) return;
        foreach (JsonElement element in value.EnumerateArray()) {
            if (element.ValueKind == JsonValueKind.String) item.Contributors.Add(new BibliographyContributor(role, new BibliographyName { Literal = element.GetString() }));
            else if (element.ValueKind == JsonValueKind.Object) {
                var name = new BibliographyName();
                var seenProperties = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (JsonProperty property in element.EnumerateObject()) {
                    string scalar = Scalar(property.Value);
                    if (!seenProperties.Add(property.Name)) { string raw = GetBoundedRawValue(property.Value, items, limits); name.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.CslJson, property.Name, ScalarOrRaw(property.Value, raw), raw)); continue; }
                    switch (property.Name.ToLowerInvariant()) {
                        case "given": name.Given = scalar; break; case "family": name.Family = scalar; break; case "literal": name.Literal = scalar; break;
                        case "suffix": name.Suffix = scalar; break; case "dropping-particle": name.DroppingParticle = scalar; break; case "non-dropping-particle": name.NonDroppingParticle = scalar; break;
                        default: string raw = GetBoundedRawValue(property.Value, items, limits); name.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.CslJson, property.Name, ScalarOrRaw(property.Value, raw), raw)); break;
                    }
                }
                item.Contributors.Add(new BibliographyContributor(role, name));
            }
        }
    }

    private static void ParseDate(BibliographyItem item, JsonElement value, BibliographyDateRole role, BibliographyDiagnosticGuard diagnostics, IList<BibliographyItem> items, BibliographyLimitGuard limits) {
        var date = new BibliographyDate { Role = role };
        if (value.ValueKind == JsonValueKind.Object) {
            var seenProperties = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (JsonProperty property in value.EnumerateObject()) {
                if (!seenProperties.Add(property.Name)) {
                    string raw = GetBoundedRawValue(property.Value, items, limits);
                    date.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.CslJson, property.Name, ScalarOrRaw(property.Value, raw), raw));
                    continue;
                }
                if (string.Equals(property.Name, "literal", StringComparison.OrdinalIgnoreCase)) date.Literal = Scalar(property.Value);
                else if (string.Equals(property.Name, "date-parts", StringComparison.OrdinalIgnoreCase)) ParseDateParts(item, date, property.Value, role, diagnostics, items, limits);
                else { string raw = GetBoundedRawValue(property.Value, items, limits); date.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.CslJson, property.Name, ScalarOrRaw(property.Value, raw), raw)); }
            }
        } else date.Literal = Scalar(value);
        item.Dates.Add(date);
    }

    private static void ParseDateParts(BibliographyItem item, BibliographyDate date, JsonElement parts, BibliographyDateRole role, BibliographyDiagnosticGuard diagnostics, IList<BibliographyItem> items, BibliographyLimitGuard limits) {
        JsonElement[] ranges = parts.ValueKind == JsonValueKind.Array ? parts.EnumerateArray().ToArray() : Array.Empty<JsonElement>();
        int? year = null; int? month = null; int? day = null;
        int? endYear = null; int? endMonth = null; int? endDay = null;
        bool valid = ranges.Length >= 1 && ranges.Length <= 2 && TryReadDatePart(ranges[0], out year, out month, out day);
        if (valid && ranges.Length == 2) valid = TryReadDatePart(ranges[1], out endYear, out endMonth, out endDay);
        if (valid) {
            date.Year = year; date.Month = month; date.Day = day;
            date.EndYear = endYear; date.EndMonth = endMonth; date.EndDay = endDay;
            return;
        }
        string raw = GetBoundedRawValue(parts, items, limits);
        date.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.CslJson, "date-parts", raw, raw));
        diagnostics.Add(new BibliographyDiagnostic("BIBCSL005", BibliographyDiagnosticSeverity.Warning, "CSL JSON date-parts could not be represented by the typed date model and were retained as native JSON.", itemKey: item.Key, field: role.ToString()));
    }

    private static void WriteString(Utf8JsonWriter writer, string name, string? value) { if (!string.IsNullOrWhiteSpace(value)) writer.WriteString(name, value); }
    private static void WriteNames(Utf8JsonWriter writer, BibliographyItem item, BibliographyContributorRole role, string property, BibliographyConversionReport report) {
        BibliographyContributor[] contributors = item.Contributors.Where(contributor => contributor.Role == role).ToArray();
        if (contributors.Length == 0) return;
        writer.WritePropertyName(property); writer.WriteStartArray();
        foreach (BibliographyContributor contributor in contributors) {
            writer.WriteStartObject(); WriteString(writer, "literal", contributor.Name.Literal); WriteString(writer, "given", contributor.Name.Given); WriteString(writer, "family", contributor.Name.Family);
            WriteString(writer, "suffix", contributor.Name.Suffix); WriteString(writer, "dropping-particle", contributor.Name.DroppingParticle); WriteString(writer, "non-dropping-particle", contributor.Name.NonDroppingParticle);
            var known = new HashSet<string>(new[] { "literal", "given", "family", "suffix", "dropping-particle", "non-dropping-particle" }, StringComparer.OrdinalIgnoreCase);
            foreach (BibliographyNativeField field in contributor.Name.NativeFields) {
                if (field.Format == BibliographyFormat.CslJson && !known.Contains(field.Name)) { writer.WritePropertyName(field.Name); bool exact = WriteNativeValue(writer, field); known.Add(field.Name); report.Add(exact ? "BIBCONV016" : "BIBCONV127", exact ? BibliographyDiagnosticSeverity.Information : BibliographyDiagnosticSeverity.Warning, exact ? $"Preserved native CSL JSON name property '{field.Name}'." : $"Native CSL JSON name property '{field.Name}' was emitted as a string because its raw JSON value was invalid or too deeply nested.", exact ? BibliographyConversionAction.PreservedExtension : BibliographyConversionAction.Approximated, item, property + "." + field.Name); }
                else report.Add("BIBCONV124", BibliographyDiagnosticSeverity.Warning, $"Native name property '{field.Name}' cannot be represented safely in CSL JSON.", BibliographyConversionAction.Omitted, item, property + "." + field.Name);
            }
            writer.WriteEndObject();
        }
        writer.WriteEndArray();
    }

    private static void WriteDate(Utf8JsonWriter writer, BibliographyItem item, BibliographyDateRole role, string property, BibliographyConversionReport report) {
        BibliographyDate? date = item.GetDate(role); if (date == null) return;
        writer.WritePropertyName(property); writer.WriteStartObject();
        var emitted = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (date.Year.HasValue) {
            writer.WritePropertyName("date-parts"); writer.WriteStartArray();
            WriteDatePart(writer, date.Year.Value, date.Month, date.Day);
            if (date.EndYear.HasValue) WriteDatePart(writer, date.EndYear.Value, date.EndMonth, date.EndDay);
            writer.WriteEndArray(); emitted.Add("date-parts");
        }
        if (!string.IsNullOrWhiteSpace(date.Literal)) { writer.WriteString("literal", date.Literal); emitted.Add("literal"); }
        foreach (BibliographyNativeField field in date.NativeFields) {
            if (field.Format == BibliographyFormat.CslJson && !emitted.Contains(field.Name)) { writer.WritePropertyName(field.Name); bool exact = WriteNativeValue(writer, field); emitted.Add(field.Name); report.Add(exact ? "BIBCONV017" : "BIBCONV128", exact ? BibliographyDiagnosticSeverity.Information : BibliographyDiagnosticSeverity.Warning, exact ? $"Preserved native CSL JSON date property '{field.Name}'." : $"Native CSL JSON date property '{field.Name}' was emitted as a string because its raw JSON value was invalid or too deeply nested.", exact ? BibliographyConversionAction.PreservedExtension : BibliographyConversionAction.Approximated, item, property + "." + field.Name); }
            else report.Add("BIBCONV125", BibliographyDiagnosticSeverity.Warning, $"Native date property '{field.Name}' cannot be represented safely in CSL JSON.", BibliographyConversionAction.Omitted, item, property + "." + field.Name);
        }
        writer.WriteEndObject();
    }

    private static bool TryReadDatePart(JsonElement value, out int? year, out int? month, out int? day) {
        year = null; month = null; day = null;
        if (value.ValueKind != JsonValueKind.Array) return false;
        JsonElement[] parts = value.EnumerateArray().ToArray();
        if (parts.Length < 1 || parts.Length > 3) return false;
        var numbers = new int[parts.Length];
        for (int index = 0; index < parts.Length; index++) if (parts[index].ValueKind != JsonValueKind.Number || !parts[index].TryGetInt32(out numbers[index])) return false;
        year = numbers[0];
        if (numbers.Length > 1) month = numbers[1];
        if (numbers.Length > 2) day = numbers[2];
        return true;
    }

    private static void WriteDatePart(Utf8JsonWriter writer, int year, int? month, int? day) {
        writer.WriteStartArray(); writer.WriteNumberValue(year); if (month.HasValue) writer.WriteNumberValue(month.Value); if (day.HasValue) writer.WriteNumberValue(day.Value); writer.WriteEndArray();
    }

    private static bool TryWriteRaw(Utf8JsonWriter writer, string raw) {
        try { using JsonDocument value = JsonDocument.Parse(raw, new JsonDocumentOptions { MaxDepth = 1024 }); value.RootElement.WriteTo(writer); return true; } catch (JsonException) { return false; }
    }
    private static bool WriteNativeValue(Utf8JsonWriter writer, BibliographyNativeField field) {
        string? raw = field.UnmodifiedRawValue;
        if (raw == null) { writer.WriteStringValue(field.Value); return true; }
        if (TryWriteRaw(writer, raw)) return true;
        writer.WriteStringValue(field.Value);
        return false;
    }

    private static string GetBoundedRawValue(JsonElement value, IList<BibliographyItem> items, BibliographyLimitGuard limits) {
        string raw = value.GetRawText();
        limits.CheckValueLength(items, raw, 0);
        return raw;
    }

    private static void GetJsonLocation(string source, JsonException exception, out int offset, out int line, out int column) {
        if (!exception.LineNumber.HasValue || !exception.BytePositionInLine.HasValue) { offset = -1; line = -1; column = -1; return; }
        int zeroBasedLine = checked((int)exception.LineNumber.Value);
        int lineStart = 0;
        for (int currentLine = 0; currentLine < zeroBasedLine && lineStart < source.Length; currentLine++) {
            int newLine = source.IndexOf('\n', lineStart);
            if (newLine < 0) { lineStart = source.Length; break; }
            lineStart = newLine + 1;
        }
        int lineEnd = source.IndexOf('\n', lineStart);
        if (lineEnd < 0) lineEnd = source.Length;
        if (lineEnd > lineStart && source[lineEnd - 1] == '\r') lineEnd--;
        long bytePosition = exception.BytePositionInLine.Value;
        int characterCount = 0;
        while (lineStart + characterCount < lineEnd) {
            int width = char.IsHighSurrogate(source[lineStart + characterCount]) && lineStart + characterCount + 1 < lineEnd && char.IsLowSurrogate(source[lineStart + characterCount + 1]) ? 2 : 1;
            char current = source[lineStart + characterCount];
            int bytes = width == 2 ? 4 : current <= 0x7F ? 1 : current <= 0x7FF ? 2 : 3;
            if (bytes > bytePosition) break;
            bytePosition -= bytes;
            characterCount += width;
        }
        offset = lineStart + characterCount;
        line = zeroBasedLine + 1;
        column = characterCount + 1;
    }
    private static string NormalizeLineEndings(string value, string lineEnding) => value.Replace("\r\n", "\n").Replace("\r", "\n").Replace("\n", lineEnding);
}
