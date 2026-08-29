using System.Text.Json;

namespace OfficeIMO.Bibliography;

internal static class CslJsonCodec {

    internal static IList<BibliographyItem> Parse(string source, BibliographyReadOptions options, List<BibliographyDiagnostic> diagnostics, out bool singleObjectRoot, CancellationToken cancellationToken) {
        var items = new List<BibliographyItem>();
        singleObjectRoot = false;
        var limits = new BibliographyLimitGuard(options);
        var diagnosticGuard = new BibliographyDiagnosticGuard(options, diagnostics, items);
        try {
            using JsonDocument json = JsonDocument.Parse(source, new JsonDocumentOptions { AllowTrailingCommas = true, CommentHandling = JsonCommentHandling.Skip, MaxDepth = options.MaximumNestingDepth });
            if (json.RootElement.ValueKind == JsonValueKind.Array) {
                foreach (JsonElement element in json.RootElement.EnumerateArray()) ParseItem(element, items, limits, diagnosticGuard, cancellationToken);
            } else if (json.RootElement.ValueKind == JsonValueKind.Object) {
                singleObjectRoot = true;
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
            bool preserveSingleObjectRoot = document.Items.Count == 1 && document.CslJsonSingleObjectRoot;
            if (!preserveSingleObjectRoot) writer.WriteStartArray();
            for (int itemIndex = 0; itemIndex < document.Items.Count; itemIndex++) {
                BibliographyItem item = document.Items[itemIndex];
                cancellationToken.ThrowIfCancellationRequested();
                writer.WriteStartObject();
                if (ShouldWriteTypedId(item)) writer.WriteString("id", CodecMappings.OutputKey(item, itemIndex));
                if (ShouldWriteTypedType(item)) writer.WriteString("type", item.Type == BibliographyItemType.Unknown && !string.IsNullOrWhiteSpace(item.NativeType) ? item.NativeType : CodecMappings.ToCslType(item.Type));
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

                HashSet<string> emitted = GetEmittedProperties(item);
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
            if (!preserveSingleObjectRoot) writer.WriteEndArray();
        }
        string text = Encoding.UTF8.GetString(stream.ToArray());
        if (document.CslJsonSingleObjectRoot && document.Items.Count != 1) report.Add("BIBCONV130", BibliographyDiagnosticSeverity.Warning, "The single-item CSL JSON object root cannot represent the current item count.", BibliographyConversionAction.Approximated, field: "root-shape");
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
        var seenProperties = new HashSet<string>(StringComparer.Ordinal);
        foreach (JsonProperty property in element.EnumerateObject()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!seenProperties.Add(property.Name)) {
                string duplicateRaw = GetBoundedRawValue(property.Value, items, limits);
                item.NativeFields.Add(BibliographyNativeField.FromParsedSource(BibliographyFormat.CslJson, property.Name, ScalarOrRaw(property.Value, duplicateRaw), duplicateRaw));
                continue;
            }
            switch (property.Name) {
                case "id": BindScalar(item, property, assigned => item.Key = assigned, items, limits); break;
                case "type": if (TryReadScalar(item, property, items, limits, out string type)) { item.NativeType = type; item.Type = CodecMappings.ParseCslType(type); } break;
                case "title": BindScalar(item, property, assigned => item.Title = assigned, items, limits); break;
                case "container-title": BindScalar(item, property, assigned => item.ContainerTitle = assigned, items, limits); break;
                case "collection-title": BindScalar(item, property, assigned => item.CollectionTitle = assigned, items, limits); break;
                case "publisher": BindScalar(item, property, assigned => item.Publisher = assigned, items, limits); break;
                case "publisher-place": BindScalar(item, property, assigned => item.PublisherPlace = assigned, items, limits); break;
                case "edition": BindScalar(item, property, assigned => item.Edition = assigned, items, limits); break;
                case "volume": BindScalar(item, property, assigned => item.Volume = assigned, items, limits); break;
                case "issue": BindScalar(item, property, assigned => item.Issue = assigned, items, limits); break;
                case "page": BindScalar(item, property, assigned => item.Pages = assigned, items, limits); break;
                case "abstract": BindScalar(item, property, assigned => item.Abstract = assigned, items, limits); break;
                case "language": BindScalar(item, property, assigned => item.Language = assigned, items, limits); break;
                case "URL": BindScalar(item, property, assigned => item.Url = assigned, items, limits); break;
                case "author": PreserveWrongShapedNames(item, property, BibliographyContributorRole.Author, items, limits); break;
                case "editor": PreserveWrongShapedNames(item, property, BibliographyContributorRole.Editor, items, limits); break;
                case "translator": PreserveWrongShapedNames(item, property, BibliographyContributorRole.Translator, items, limits); break;
                case "recipient": PreserveWrongShapedNames(item, property, BibliographyContributorRole.Recipient, items, limits); break;
                case "interviewer": PreserveWrongShapedNames(item, property, BibliographyContributorRole.Interviewer, items, limits); break;
                case "composer": PreserveWrongShapedNames(item, property, BibliographyContributorRole.Composer, items, limits); break;
                case "collection-editor": PreserveWrongShapedNames(item, property, BibliographyContributorRole.CollectionEditor, items, limits); break;
                case "issued": ParseDate(item, property, BibliographyDateRole.Issued, diagnostics, items, limits); break;
                case "accessed": ParseDate(item, property, BibliographyDateRole.Accessed, diagnostics, items, limits); break;
                case "submitted": ParseDate(item, property, BibliographyDateRole.Submitted, diagnostics, items, limits); break;
                case "original-date": ParseDate(item, property, BibliographyDateRole.Original, diagnostics, items, limits); break;
                case "event-date": ParseDate(item, property, BibliographyDateRole.Event, diagnostics, items, limits); break;
                case "DOI": BindIdentifier(item, property, "DOI", items, limits); break;
                case "ISBN": BindIdentifier(item, property, "ISBN", items, limits); break;
                case "ISSN": BindIdentifier(item, property, "ISSN", items, limits); break;
                case "PMID": BindIdentifier(item, property, "PMID", items, limits); break;
                case "PMCID": BindIdentifier(item, property, "PMCID", items, limits); break;
                case "keyword": if (TryReadScalar(item, property, items, limits, out string keyword) && !string.IsNullOrWhiteSpace(keyword)) item.Keywords.Add(keyword); break;
                case "note": if (TryReadScalar(item, property, items, limits, out string note)) item.Notes.Add(note); break;
                default:
                    string raw = GetBoundedRawValue(property.Value, items, limits);
                    item.NativeFields.Add(BibliographyNativeField.FromParsedSource(BibliographyFormat.CslJson, property.Name, ScalarOrRaw(property.Value, raw), raw));
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

    private static void BindScalar(BibliographyItem item, JsonProperty property, Action<string> assign, IList<BibliographyItem> items, BibliographyLimitGuard limits) {
        if (TryReadScalar(item, property, items, limits, out string scalar)) assign(scalar);
    }

    private static void BindIdentifier(BibliographyItem item, JsonProperty property, string scheme, IList<BibliographyItem> items, BibliographyLimitGuard limits) {
        if (TryReadScalar(item, property, items, limits, out string scalar)) CodecMappings.AddIdentifier(item, scheme, scalar);
    }

    private static bool TryReadScalar(BibliographyItem item, JsonProperty property, IList<BibliographyItem> items, BibliographyLimitGuard limits, out string scalar) {
        if (property.Value.ValueKind != JsonValueKind.Object && property.Value.ValueKind != JsonValueKind.Array) { scalar = Scalar(property.Value); return true; }
        string raw = GetBoundedRawValue(property.Value, items, limits);
        item.NativeFields.Add(BibliographyNativeField.FromParsedSource(BibliographyFormat.CslJson, property.Name, raw, raw));
        scalar = string.Empty;
        return false;
    }

    private static void PreserveWrongShapedNames(BibliographyItem item, JsonProperty property, BibliographyContributorRole role, IList<BibliographyItem> items, BibliographyLimitGuard limits) {
        if (ParseNames(item, property.Value, role, items, limits)) return;
        string raw = GetBoundedRawValue(property.Value, items, limits);
        item.NativeFields.Add(BibliographyNativeField.FromParsedSource(BibliographyFormat.CslJson, property.Name, ScalarOrRaw(property.Value, raw), raw));
    }

    private static bool ParseNames(BibliographyItem item, JsonElement value, BibliographyContributorRole role, IList<BibliographyItem> items, BibliographyLimitGuard limits) {
        if (value.ValueKind != JsonValueKind.Array) return false;
        foreach (JsonElement element in value.EnumerateArray()) {
            if (element.ValueKind != JsonValueKind.String && element.ValueKind != JsonValueKind.Object) return false;
            if (element.ValueKind == JsonValueKind.Object && element.EnumerateObject().Any(property => IsKnownNameProperty(property.Name) && (property.Value.ValueKind == JsonValueKind.Object || property.Value.ValueKind == JsonValueKind.Array))) return false;
        }
        foreach (JsonElement element in value.EnumerateArray()) {
            if (element.ValueKind == JsonValueKind.String) item.Contributors.Add(new BibliographyContributor(role, new BibliographyName { Literal = element.GetString() }));
            else if (element.ValueKind == JsonValueKind.Object) {
                var name = new BibliographyName();
                    var seenProperties = new HashSet<string>(StringComparer.Ordinal);
                foreach (JsonProperty property in element.EnumerateObject()) {
                    string scalar = Scalar(property.Value);
                    if (!seenProperties.Add(property.Name)) { string raw = GetBoundedRawValue(property.Value, items, limits); name.NativeFields.Add(BibliographyNativeField.FromParsedSource(BibliographyFormat.CslJson, property.Name, ScalarOrRaw(property.Value, raw), raw)); continue; }
                    switch (property.Name) {
                        case "given": name.Given = scalar; break; case "family": name.Family = scalar; break; case "literal": name.Literal = scalar; break;
                        case "suffix": name.Suffix = scalar; break; case "dropping-particle": name.DroppingParticle = scalar; break; case "non-dropping-particle": name.NonDroppingParticle = scalar; break;
                        default: string raw = GetBoundedRawValue(property.Value, items, limits); name.NativeFields.Add(BibliographyNativeField.FromParsedSource(BibliographyFormat.CslJson, property.Name, ScalarOrRaw(property.Value, raw), raw)); break;
                    }
                }
                item.Contributors.Add(new BibliographyContributor(role, name));
            }
        }
        return true;
    }

    private static bool IsKnownNameProperty(string name) {
        switch (name) {
            case "given": case "family": case "literal": case "suffix": case "dropping-particle": case "non-dropping-particle": return true;
            default: return false;
        }
    }

    private static void ParseDate(BibliographyItem item, JsonProperty sourceProperty, BibliographyDateRole role, BibliographyDiagnosticGuard diagnostics, IList<BibliographyItem> items, BibliographyLimitGuard limits) {
        JsonElement value = sourceProperty.Value;
        if (value.ValueKind != JsonValueKind.Object) {
            string raw = GetBoundedRawValue(value, items, limits);
            item.NativeFields.Add(BibliographyNativeField.FromParsedSource(BibliographyFormat.CslJson, sourceProperty.Name, ScalarOrRaw(value, raw), raw));
            return;
        }
        var date = new BibliographyDate { Role = role };
        var seenProperties = new HashSet<string>(StringComparer.Ordinal);
        foreach (JsonProperty property in value.EnumerateObject()) {
            if (!seenProperties.Add(property.Name)) {
                string duplicateRaw = GetBoundedRawValue(property.Value, items, limits);
                date.NativeFields.Add(BibliographyNativeField.FromParsedSource(BibliographyFormat.CslJson, property.Name, ScalarOrRaw(property.Value, duplicateRaw), duplicateRaw));
                continue;
            }
            if (string.Equals(property.Name, "literal", StringComparison.Ordinal)) {
                if (property.Value.ValueKind == JsonValueKind.Object || property.Value.ValueKind == JsonValueKind.Array) { string literalRaw = GetBoundedRawValue(property.Value, items, limits); date.NativeFields.Add(BibliographyNativeField.FromParsedSource(BibliographyFormat.CslJson, property.Name, literalRaw, literalRaw)); }
                else date.Literal = Scalar(property.Value);
            } else if (string.Equals(property.Name, "date-parts", StringComparison.Ordinal)) ParseDateParts(item, date, property.Value, role, diagnostics, items, limits);
            else { string raw = GetBoundedRawValue(property.Value, items, limits); date.NativeFields.Add(BibliographyNativeField.FromParsedSource(BibliographyFormat.CslJson, property.Name, ScalarOrRaw(property.Value, raw), raw)); }
        }
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
    private static HashSet<string> GetEmittedProperties(BibliographyItem item) {
        var emitted = new HashSet<string>(StringComparer.Ordinal);
        if (ShouldWriteTypedId(item)) emitted.Add("id");
        if (ShouldWriteTypedType(item)) emitted.Add("type");
        AddIfValue("title", item.Title); AddIfValue("container-title", item.ContainerTitle); AddIfValue("collection-title", item.CollectionTitle);
        AddIfValue("publisher", item.Publisher); AddIfValue("publisher-place", item.PublisherPlace); AddIfValue("edition", item.Edition);
        AddIfValue("volume", item.Volume); AddIfValue("issue", item.Issue); AddIfValue("page", item.Pages); AddIfValue("abstract", item.Abstract);
        AddIfValue("language", item.Language); AddIfValue("URL", item.Url);
        AddIfContributors(BibliographyContributorRole.Author, "author"); AddIfContributors(BibliographyContributorRole.Editor, "editor");
        AddIfContributors(BibliographyContributorRole.Translator, "translator"); AddIfContributors(BibliographyContributorRole.Recipient, "recipient");
        AddIfContributors(BibliographyContributorRole.Interviewer, "interviewer"); AddIfContributors(BibliographyContributorRole.Composer, "composer");
        AddIfContributors(BibliographyContributorRole.CollectionEditor, "collection-editor");
        AddIfDate(BibliographyDateRole.Issued, "issued"); AddIfDate(BibliographyDateRole.Accessed, "accessed");
        AddIfDate(BibliographyDateRole.Submitted, "submitted"); AddIfDate(BibliographyDateRole.Original, "original-date"); AddIfDate(BibliographyDateRole.Event, "event-date");
        foreach (BibliographyIdentifier identifier in item.Identifiers.Where(identifier => CodecMappings.IsCslIdentifierScheme(identifier.Scheme))) emitted.Add(identifier.Scheme.ToUpperInvariant());
        if (item.Keywords.Count > 0) emitted.Add("keyword");
        if (item.Notes.Count > 0) emitted.Add("note");
        return emitted;

        void AddIfValue(string name, string? value) { if (!string.IsNullOrWhiteSpace(value)) emitted.Add(name); }
        void AddIfContributors(BibliographyContributorRole role, string name) { if (item.Contributors.Any(contributor => contributor.Role == role)) emitted.Add(name); }
        void AddIfDate(BibliographyDateRole role, string name) { if (item.GetDate(role) != null) emitted.Add(name); }
    }

    private static bool ShouldWriteTypedId(BibliographyItem item) => !string.IsNullOrWhiteSpace(item.Key) || !HasNativeProperty(item, "id");
    private static bool ShouldWriteTypedType(BibliographyItem item) => item.Type != BibliographyItemType.Unknown || !string.IsNullOrWhiteSpace(item.NativeType) || !HasNativeProperty(item, "type");
    private static bool HasNativeProperty(BibliographyItem item, string property) => item.NativeFields.Any(field => field.Format == BibliographyFormat.CslJson && string.Equals(field.Name, property, StringComparison.Ordinal));
    private static void WriteNames(Utf8JsonWriter writer, BibliographyItem item, BibliographyContributorRole role, string property, BibliographyConversionReport report) {
        BibliographyContributor[] contributors = item.Contributors.Where(contributor => contributor.Role == role).ToArray();
        if (contributors.Length == 0) return;
        writer.WritePropertyName(property); writer.WriteStartArray();
        foreach (BibliographyContributor contributor in contributors) {
            writer.WriteStartObject(); WriteString(writer, "literal", contributor.Name.Literal); WriteString(writer, "given", contributor.Name.Given); WriteString(writer, "family", contributor.Name.Family);
            WriteString(writer, "suffix", contributor.Name.Suffix); WriteString(writer, "dropping-particle", contributor.Name.DroppingParticle); WriteString(writer, "non-dropping-particle", contributor.Name.NonDroppingParticle);
            var known = new HashSet<string>(new[] { "literal", "given", "family", "suffix", "dropping-particle", "non-dropping-particle" }, StringComparer.Ordinal);
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
        var emitted = new HashSet<string>(StringComparer.Ordinal);
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
