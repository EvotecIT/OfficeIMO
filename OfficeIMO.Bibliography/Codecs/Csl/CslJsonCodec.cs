using System.Buffers;
using System.Text.Json;

namespace OfficeIMO.Bibliography;

internal static class CslJsonCodec {
    internal const int NativeJsonMaximumDepth = 124;
    private const int JsonWriterMaximumDepth = 128;
    internal static JsonDocumentOptions NativeJsonDocumentOptions => new JsonDocumentOptions { AllowTrailingCommas = true, CommentHandling = JsonCommentHandling.Skip, MaxDepth = NativeJsonMaximumDepth };

    internal static IList<BibliographyItem> Parse(string source, BibliographyReadOptions options, List<BibliographyDiagnostic> diagnostics, out bool singleObjectRoot, CancellationToken cancellationToken) {
        var items = new List<BibliographyItem>();
        singleObjectRoot = false;
        var limits = new BibliographyLimitGuard(options);
        var diagnosticGuard = new BibliographyDiagnosticGuard(options, diagnostics, items);
        try {
            using JsonDocument json = ParseDocument(source, options, items, cancellationToken);
            if (json.RootElement.ValueKind == JsonValueKind.Array) {
                foreach (JsonElement element in json.RootElement.EnumerateArray()) ParseItem(element, items, limits, diagnosticGuard, cancellationToken);
            } else if (json.RootElement.ValueKind == JsonValueKind.Object) {
                singleObjectRoot = true;
                ParseItem(json.RootElement, items, limits, diagnosticGuard, cancellationToken);
            } else {
                diagnosticGuard.Add(new BibliographyDiagnostic("BIBCSL001", BibliographyDiagnosticSeverity.Error, "CSL JSON root must be an object or an array."));
            }
        } catch (JsonException exception) {
            GetJsonLocation(source, exception, cancellationToken, out int offset, out int line, out int column);
            diagnosticGuard.Add(new BibliographyDiagnostic("BIBCSL002", BibliographyDiagnosticSeverity.Error, exception.Message, offset, line, column));
        }
        return items;
    }

    internal static string Write(BibliographyDocument document, BibliographyWriteOptions options, BibliographyConversionReport report, CancellationToken cancellationToken) {
        using var stream = new MemoryStream();
        string[] outputKeys = CodecMappings.OutputKeys(document.Items, BibliographyFormat.CslJson, cancellationToken);
        using (var writer = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = true, MaxDepth = JsonWriterMaximumDepth })) {
            bool preserveSingleObjectRoot = document.Items.Count == 1 && document.CslJsonSingleObjectRoot;
            if (!preserveSingleObjectRoot) writer.WriteStartArray();
            for (int itemIndex = 0; itemIndex < document.Items.Count; itemIndex++) {
                BibliographyItem item = document.Items[itemIndex];
                cancellationToken.ThrowIfCancellationRequested();
                writer.WriteStartObject();
                if (ShouldWriteTypedId(item, cancellationToken)) WriteString(writer, "id", outputKeys[itemIndex], cancellationToken);
                if (ShouldWriteTypedType(item, cancellationToken)) WriteString(writer, "type", OutputType(document.SourceFormat, item), cancellationToken);
                WriteString(writer, "title", item.Title, cancellationToken); WriteString(writer, "container-title", item.ContainerTitle, cancellationToken); WriteString(writer, "collection-title", item.CollectionTitle, cancellationToken);
                WriteString(writer, "publisher", item.Publisher, cancellationToken); WriteString(writer, "publisher-place", item.PublisherPlace, cancellationToken); WriteString(writer, "edition", item.Edition, cancellationToken);
                WriteString(writer, "volume", item.Volume, cancellationToken); WriteString(writer, "issue", item.Issue, cancellationToken); WriteString(writer, "page", item.Pages, cancellationToken); WriteString(writer, "abstract", item.Abstract, cancellationToken);
                WriteString(writer, "language", item.Language, cancellationToken); WriteString(writer, "URL", item.Url, cancellationToken);
                WriteNames(writer, item, BibliographyContributorRole.Author, "author", report, cancellationToken); WriteNames(writer, item, BibliographyContributorRole.Editor, "editor", report, cancellationToken);
                WriteNames(writer, item, BibliographyContributorRole.Translator, "translator", report, cancellationToken); WriteNames(writer, item, BibliographyContributorRole.Recipient, "recipient", report, cancellationToken);
                WriteNames(writer, item, BibliographyContributorRole.Interviewer, "interviewer", report, cancellationToken); WriteNames(writer, item, BibliographyContributorRole.Composer, "composer", report, cancellationToken);
                WriteNames(writer, item, BibliographyContributorRole.CollectionEditor, "collection-editor", report, cancellationToken);
                foreach (BibliographyDateRole role in GetDistinctDateRoles(item, cancellationToken)) {
                    string? property = DateProperty(role);
                    if (property != null) WriteDate(writer, item, role, property, report, cancellationToken);
                }
                WriteIdentifiers(writer, item, cancellationToken);
                if (item.Keywords.Count > 0) WriteString(writer, "keyword", JoinValues(item.Keywords, ", ", cancellationToken), cancellationToken);
                if (item.Notes.Count > 0) WriteString(writer, "note", JoinValues(item.Notes, "; ", cancellationToken), cancellationToken);

                HashSet<string> emitted = GetEmittedProperties(item, cancellationToken);
                foreach (BibliographyNativeField field in item.NativeFields) {
                    cancellationToken.ThrowIfCancellationRequested();
                    bool changesOwner = field.Format == BibliographyFormat.CslJson && WouldBindTypedItemProperty(field, cancellationToken);
                    if (field.Format == BibliographyFormat.CslJson && !emitted.Contains(field.Name) && !changesOwner) {
                        WritePropertyName(writer, field.Name, cancellationToken);
                        bool exact = WriteNativeValue(writer, field, cancellationToken);
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
        foreach (BibliographyNativeEntry entry in document.NativeEntries) {
            cancellationToken.ThrowIfCancellationRequested();
            report.Add("BIBCONV121", BibliographyDiagnosticSeverity.Warning, $"Document-level {entry.Format} entry '{entry.Kind}' cannot be represented in CSL JSON.", BibliographyConversionAction.Omitted, field: entry.Name ?? entry.Kind);
        }
        return options.LineEnding == "\n" ? text + options.LineEnding : NormalizeLineEndings(text, options.LineEnding) + options.LineEnding;
    }

    private static void ParseItem(JsonElement element, IList<BibliographyItem> items, BibliographyLimitGuard limits, BibliographyDiagnosticGuard diagnostics, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (element.ValueKind != JsonValueKind.Object) { diagnostics.Add(new BibliographyDiagnostic("BIBCSL003", BibliographyDiagnosticSeverity.Warning, "Ignored a non-object CSL JSON array value.")); return; }
        limits.AddItem(items, 0);
        var item = new BibliographyItem();
        items.Add(item);
        CountValues(element, items, limits, cancellationToken);
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
                case "author": PreserveWrongShapedNames(item, property, BibliographyContributorRole.Author, items, limits, cancellationToken); break;
                case "editor": PreserveWrongShapedNames(item, property, BibliographyContributorRole.Editor, items, limits, cancellationToken); break;
                case "translator": PreserveWrongShapedNames(item, property, BibliographyContributorRole.Translator, items, limits, cancellationToken); break;
                case "recipient": PreserveWrongShapedNames(item, property, BibliographyContributorRole.Recipient, items, limits, cancellationToken); break;
                case "interviewer": PreserveWrongShapedNames(item, property, BibliographyContributorRole.Interviewer, items, limits, cancellationToken); break;
                case "composer": PreserveWrongShapedNames(item, property, BibliographyContributorRole.Composer, items, limits, cancellationToken); break;
                case "collection-editor": PreserveWrongShapedNames(item, property, BibliographyContributorRole.CollectionEditor, items, limits, cancellationToken); break;
                case "issued": ParseDate(item, property, BibliographyDateRole.Issued, diagnostics, items, limits, cancellationToken); break;
                case "accessed": ParseDate(item, property, BibliographyDateRole.Accessed, diagnostics, items, limits, cancellationToken); break;
                case "submitted": ParseDate(item, property, BibliographyDateRole.Submitted, diagnostics, items, limits, cancellationToken); break;
                case "original-date": ParseDate(item, property, BibliographyDateRole.Original, diagnostics, items, limits, cancellationToken); break;
                case "event-date": ParseDate(item, property, BibliographyDateRole.Event, diagnostics, items, limits, cancellationToken); break;
                case "DOI": BindIdentifier(item, property, "DOI", items, limits); break;
                case "ISBN": BindIdentifier(item, property, "ISBN", items, limits); break;
                case "ISSN": BindIdentifier(item, property, "ISSN", items, limits); break;
                case "PMID": BindIdentifier(item, property, "PMID", items, limits); break;
                case "PMCID": BindIdentifier(item, property, "PMCID", items, limits); break;
                case "keyword": if (TryReadScalar(item, property, items, limits, out string keyword)) item.Keywords.Add(keyword); break;
                case "note": if (TryReadScalar(item, property, items, limits, out string note)) item.Notes.Add(note); break;
                default:
                    string raw = GetBoundedRawValue(property.Value, items, limits);
                    item.NativeFields.Add(BibliographyNativeField.FromParsedSource(BibliographyFormat.CslJson, property.Name, ScalarOrRaw(property.Value, raw), raw));
                    break;
            }
        }
        if (string.IsNullOrWhiteSpace(item.Key)) diagnostics.Add(new BibliographyDiagnostic("BIBCSL004", BibliographyDiagnosticSeverity.Warning, "CSL JSON item has no id."));
    }

    private static void CountValues(JsonElement element, IList<BibliographyItem> items, BibliographyLimitGuard limits, CancellationToken cancellationToken) {
        switch (element.ValueKind) {
            case JsonValueKind.Object:
                foreach (JsonProperty property in element.EnumerateObject()) {
                    cancellationToken.ThrowIfCancellationRequested();
                    limits.AddValue(items, property.Name, 0);
                    if (property.Value.ValueKind == JsonValueKind.Object || property.Value.ValueKind == JsonValueKind.Array) CountValues(property.Value, items, limits, cancellationToken);
                    else limits.CheckValueLength(items, GetScalarValueLength(property.Value), 0);
                }
                break;
            case JsonValueKind.Array:
                foreach (JsonElement value in element.EnumerateArray()) {
                    cancellationToken.ThrowIfCancellationRequested();
                    limits.AddValue(items, null, 0);
                    if (value.ValueKind == JsonValueKind.Object || value.ValueKind == JsonValueKind.Array) CountValues(value, items, limits, cancellationToken);
                    else limits.CheckValueLength(items, GetScalarValueLength(value), 0);
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

    private static int GetScalarValueLength(JsonElement value) => value.ValueKind == JsonValueKind.String ? (value.GetString()?.Length ?? 0) : value.GetRawText().Length;

    private static string ScalarOrRaw(JsonElement value, string raw) => value.ValueKind == JsonValueKind.Object || value.ValueKind == JsonValueKind.Array ? raw : Scalar(value);

    private static void BindScalar(BibliographyItem item, JsonProperty property, Action<string> assign, IList<BibliographyItem> items, BibliographyLimitGuard limits) {
        if (TryReadScalar(item, property, items, limits, out string scalar)) assign(scalar);
    }

    private static void BindIdentifier(BibliographyItem item, JsonProperty property, string scheme, IList<BibliographyItem> items, BibliographyLimitGuard limits) {
        if (!TryReadScalar(item, property, items, limits, out string scalar)) return;
        if (!string.IsNullOrWhiteSpace(scalar)) CodecMappings.AddIdentifier(item, scheme, scalar);
        else item.NativeFields.Add(BibliographyNativeField.FromParsedSource(BibliographyFormat.CslJson, property.Name, scalar, property.Value.GetRawText()));
    }

    private static bool TryReadScalar(BibliographyItem item, JsonProperty property, IList<BibliographyItem> items, BibliographyLimitGuard limits, out string scalar) {
        if (property.Value.ValueKind == JsonValueKind.String) { scalar = property.Value.GetString() ?? string.Empty; return true; }
        string raw = GetBoundedRawValue(property.Value, items, limits);
        item.NativeFields.Add(BibliographyNativeField.FromParsedSource(BibliographyFormat.CslJson, property.Name, ScalarOrRaw(property.Value, raw), raw));
        scalar = string.Empty;
        return false;
    }

    private static void PreserveWrongShapedNames(BibliographyItem item, JsonProperty property, BibliographyContributorRole role, IList<BibliographyItem> items, BibliographyLimitGuard limits, CancellationToken cancellationToken) {
        if (ParseNames(item, property.Value, role, items, limits, cancellationToken)) return;
        string raw = GetBoundedRawValue(property.Value, items, limits);
        item.NativeFields.Add(BibliographyNativeField.FromParsedSource(BibliographyFormat.CslJson, property.Name, ScalarOrRaw(property.Value, raw), raw));
    }

    internal static bool ParseNames(BibliographyItem item, JsonElement value, BibliographyContributorRole role, IList<BibliographyItem> items, BibliographyLimitGuard limits, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (value.ValueKind != JsonValueKind.Array) return false;
        if (value.GetArrayLength() == 0) return false;
        foreach (JsonElement element in value.EnumerateArray()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (element.ValueKind != JsonValueKind.String && element.ValueKind != JsonValueKind.Object) return false;
            if (element.ValueKind == JsonValueKind.Object && HasWrongShapedKnownNameProperty(element, cancellationToken)) return false;
        }
        foreach (JsonElement element in value.EnumerateArray()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (element.ValueKind == JsonValueKind.String) item.Contributors.Add(new BibliographyContributor(role, new BibliographyName { Literal = element.GetString() }));
            else if (element.ValueKind == JsonValueKind.Object) {
                var name = new BibliographyName();
                var seenProperties = new HashSet<string>(StringComparer.Ordinal);
                foreach (JsonProperty property in element.EnumerateObject()) {
                    cancellationToken.ThrowIfCancellationRequested();
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
        cancellationToken.ThrowIfCancellationRequested();
        return true;
    }

    private static bool HasWrongShapedKnownNameProperty(JsonElement element, CancellationToken cancellationToken) {
        foreach (JsonProperty property in element.EnumerateObject()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (IsKnownNameProperty(property.Name) && property.Value.ValueKind != JsonValueKind.String) return true;
        }
        return false;
    }

    private static bool IsKnownNameProperty(string name) {
        switch (name) {
            case "given": case "family": case "literal": case "suffix": case "dropping-particle": case "non-dropping-particle": return true;
            default: return false;
        }
    }

    private static void ParseDate(BibliographyItem item, JsonProperty sourceProperty, BibliographyDateRole role, BibliographyDiagnosticGuard diagnostics, IList<BibliographyItem> items, BibliographyLimitGuard limits, CancellationToken cancellationToken) {
        JsonElement value = sourceProperty.Value;
        if (value.ValueKind != JsonValueKind.Object) {
            string raw = GetBoundedRawValue(value, items, limits);
            item.NativeFields.Add(BibliographyNativeField.FromParsedSource(BibliographyFormat.CslJson, sourceProperty.Name, ScalarOrRaw(value, raw), raw));
            return;
        }
        var date = new BibliographyDate { Role = role };
        var seenProperties = new HashSet<string>(StringComparer.Ordinal);
        foreach (JsonProperty property in value.EnumerateObject()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!seenProperties.Add(property.Name)) {
                string duplicateRaw = GetBoundedRawValue(property.Value, items, limits);
                date.NativeFields.Add(BibliographyNativeField.FromParsedSource(BibliographyFormat.CslJson, property.Name, ScalarOrRaw(property.Value, duplicateRaw), duplicateRaw));
                continue;
            }
            if (string.Equals(property.Name, "literal", StringComparison.Ordinal)) {
                if (property.Value.ValueKind != JsonValueKind.String) { string literalRaw = GetBoundedRawValue(property.Value, items, limits); date.NativeFields.Add(BibliographyNativeField.FromParsedSource(BibliographyFormat.CslJson, property.Name, ScalarOrRaw(property.Value, literalRaw), literalRaw)); }
                else date.Literal = property.Value.GetString();
            } else if (string.Equals(property.Name, "date-parts", StringComparison.Ordinal)) ParseDateParts(item, date, property.Value, role, diagnostics, items, limits, cancellationToken);
            else { string raw = GetBoundedRawValue(property.Value, items, limits); date.NativeFields.Add(BibliographyNativeField.FromParsedSource(BibliographyFormat.CslJson, property.Name, ScalarOrRaw(property.Value, raw), raw)); }
        }
        item.Dates.Add(date);
    }

    private static void ParseDateParts(BibliographyItem item, BibliographyDate date, JsonElement parts, BibliographyDateRole role, BibliographyDiagnosticGuard diagnostics, IList<BibliographyItem> items, BibliographyLimitGuard limits, CancellationToken cancellationToken) {
        List<JsonElement> ranges = TakeBoundedElements(parts, 2, cancellationToken, out bool tooManyRanges);
        int? year = null; int? month = null; int? day = null;
        int? endYear = null; int? endMonth = null; int? endDay = null;
        bool valid = !tooManyRanges && ranges.Count >= 1 && TryReadDatePart(ranges[0], cancellationToken, out year, out month, out day);
        if (valid && ranges.Count == 2) valid = TryReadDatePart(ranges[1], cancellationToken, out endYear, out endMonth, out endDay);
        if (valid) {
            date.Year = year; date.Month = month; date.Day = day;
            date.EndYear = endYear; date.EndMonth = endMonth; date.EndDay = endDay;
            return;
        }
        string raw = GetBoundedRawValue(parts, items, limits);
        date.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.CslJson, "date-parts", raw, raw));
        diagnostics.Add(new BibliographyDiagnostic("BIBCSL005", BibliographyDiagnosticSeverity.Warning, "CSL JSON date-parts could not be represented by the typed date model and were retained as native JSON.", itemKey: item.Key, field: role.ToString()));
    }

    private static void WriteString(Utf8JsonWriter writer, string name, string? value, CancellationToken cancellationToken) { if (value != null) writer.WriteString(name, SanitizeUtf16(value, cancellationToken)); }
    private static HashSet<string> GetEmittedProperties(BibliographyItem item, CancellationToken cancellationToken) {
        var emitted = new HashSet<string>(StringComparer.Ordinal);
        if (ShouldWriteTypedId(item, cancellationToken)) emitted.Add("id");
        if (ShouldWriteTypedType(item, cancellationToken)) emitted.Add("type");
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
        foreach (BibliographyIdentifier identifier in item.Identifiers) { cancellationToken.ThrowIfCancellationRequested(); if (CodecMappings.IsCslIdentifierScheme(identifier.Scheme)) emitted.Add(identifier.Scheme.ToUpperInvariant()); }
        if (item.Keywords.Count > 0) emitted.Add("keyword");
        if (item.Notes.Count > 0) emitted.Add("note");
        return emitted;

        void AddIfValue(string name, string? value) { if (value != null) emitted.Add(name); }
        void AddIfContributors(BibliographyContributorRole role, string name) { foreach (BibliographyContributor contributor in item.Contributors) { cancellationToken.ThrowIfCancellationRequested(); if (contributor.Role == role) { emitted.Add(name); break; } } }
        void AddIfDate(BibliographyDateRole role, string name) { if (FindDate(item, role, cancellationToken) != null) emitted.Add(name); }
    }

    private static bool ShouldWriteTypedId(BibliographyItem item, CancellationToken cancellationToken) => !string.IsNullOrWhiteSpace(item.Key) || !HasNativeProperty(item, "id", cancellationToken);
    private static bool ShouldWriteTypedType(BibliographyItem item, CancellationToken cancellationToken) => item.Type != BibliographyItemType.Unknown || !string.IsNullOrWhiteSpace(item.NativeType) || !HasNativeProperty(item, "type", cancellationToken);
    internal static bool HasNativeProperty(BibliographyItem item, string property, CancellationToken cancellationToken) {
        foreach (BibliographyNativeField field in item.NativeFields) {
            cancellationToken.ThrowIfCancellationRequested();
            if (field.Format == BibliographyFormat.CslJson && string.Equals(field.Name, property, StringComparison.Ordinal)) return true;
        }
        return false;
    }
    private static void WriteNames(Utf8JsonWriter writer, BibliographyItem item, BibliographyContributorRole role, string property, BibliographyConversionReport report, CancellationToken cancellationToken) {
        var contributors = new List<BibliographyContributor>();
        foreach (BibliographyContributor contributor in item.Contributors) { cancellationToken.ThrowIfCancellationRequested(); if (contributor.Role == role) contributors.Add(contributor); }
        if (contributors.Count == 0) return;
        writer.WritePropertyName(property); writer.WriteStartArray();
        foreach (BibliographyContributor contributor in contributors) {
            cancellationToken.ThrowIfCancellationRequested();
            writer.WriteStartObject(); WriteString(writer, "literal", contributor.Name.Literal, cancellationToken); WriteString(writer, "given", contributor.Name.Given, cancellationToken); WriteString(writer, "family", contributor.Name.Family, cancellationToken);
            WriteString(writer, "suffix", contributor.Name.Suffix, cancellationToken); WriteString(writer, "dropping-particle", contributor.Name.DroppingParticle, cancellationToken); WriteString(writer, "non-dropping-particle", contributor.Name.NonDroppingParticle, cancellationToken);
            var known = new HashSet<string>(new[] { "literal", "given", "family", "suffix", "dropping-particle", "non-dropping-particle" }, StringComparer.Ordinal);
            foreach (BibliographyNativeField field in contributor.Name.NativeFields) {
                cancellationToken.ThrowIfCancellationRequested();
                if (field.Format == BibliographyFormat.CslJson && !known.Contains(field.Name)) { WritePropertyName(writer, field.Name, cancellationToken); bool exact = WriteNativeValue(writer, field, cancellationToken); known.Add(field.Name); report.Add(exact ? "BIBCONV016" : "BIBCONV127", exact ? BibliographyDiagnosticSeverity.Information : BibliographyDiagnosticSeverity.Warning, exact ? $"Preserved native CSL JSON name property '{field.Name}'." : $"Native CSL JSON name property '{field.Name}' was emitted as a string because its raw JSON value was invalid or too deeply nested.", exact ? BibliographyConversionAction.PreservedExtension : BibliographyConversionAction.Approximated, item, property + "." + field.Name); }
                else report.Add("BIBCONV124", BibliographyDiagnosticSeverity.Warning, $"Native name property '{field.Name}' cannot be represented safely in CSL JSON.", BibliographyConversionAction.Omitted, item, property + "." + field.Name);
            }
            writer.WriteEndObject();
        }
        writer.WriteEndArray();
    }

    private static void WriteDate(Utf8JsonWriter writer, BibliographyItem item, BibliographyDateRole role, string property, BibliographyConversionReport report, CancellationToken cancellationToken) {
        BibliographyDate? date = FindDate(item, role, cancellationToken); if (date == null) return;
        writer.WritePropertyName(property); writer.WriteStartObject();
        var emitted = new HashSet<string>(StringComparer.Ordinal);
        if (date.Year.HasValue) {
            writer.WritePropertyName("date-parts"); writer.WriteStartArray();
            WriteDatePart(writer, date.Year.Value, date.Month, date.Day);
            if (date.EndYear.HasValue) WriteDatePart(writer, date.EndYear.Value, date.EndMonth, date.EndDay);
            writer.WriteEndArray(); emitted.Add("date-parts");
        }
        if (date.Literal != null) { WriteString(writer, "literal", date.Literal, cancellationToken); emitted.Add("literal"); }
        foreach (BibliographyNativeField field in date.NativeFields) {
            cancellationToken.ThrowIfCancellationRequested();
            if (field.Format == BibliographyFormat.CslJson && !emitted.Contains(field.Name) && !WouldBindTypedDateProperty(field, cancellationToken)) { WritePropertyName(writer, field.Name, cancellationToken); bool exact = WriteNativeValue(writer, field, cancellationToken); emitted.Add(field.Name); report.Add(exact ? "BIBCONV017" : "BIBCONV128", exact ? BibliographyDiagnosticSeverity.Information : BibliographyDiagnosticSeverity.Warning, exact ? $"Preserved native CSL JSON date property '{field.Name}'." : $"Native CSL JSON date property '{field.Name}' was emitted as a string because its raw JSON value was invalid or too deeply nested.", exact ? BibliographyConversionAction.PreservedExtension : BibliographyConversionAction.Approximated, item, property + "." + field.Name); }
            else report.Add("BIBCONV125", BibliographyDiagnosticSeverity.Warning, $"Native date property '{field.Name}' cannot be represented safely in CSL JSON.", BibliographyConversionAction.Omitted, item, property + "." + field.Name);
        }
        writer.WriteEndObject();
    }

    private static string? DateProperty(BibliographyDateRole role) {
        switch (role) {
            case BibliographyDateRole.Issued: return "issued";
            case BibliographyDateRole.Accessed: return "accessed";
            case BibliographyDateRole.Submitted: return "submitted";
            case BibliographyDateRole.Original: return "original-date";
            case BibliographyDateRole.Event: return "event-date";
            default: return null;
        }
    }

    private static IEnumerable<BibliographyDateRole> GetDistinctDateRoles(BibliographyItem item, CancellationToken cancellationToken) {
        var emitted = new HashSet<BibliographyDateRole>();
        foreach (BibliographyDate date in item.Dates) {
            cancellationToken.ThrowIfCancellationRequested();
            if (emitted.Add(date.Role)) yield return date.Role;
        }
    }

    private static BibliographyDate? FindDate(BibliographyItem item, BibliographyDateRole role, CancellationToken cancellationToken) {
        foreach (BibliographyDate date in item.Dates) {
            cancellationToken.ThrowIfCancellationRequested();
            if (date.Role == role) return date;
        }
        return null;
    }

    private static void WriteIdentifiers(Utf8JsonWriter writer, BibliographyItem item, CancellationToken cancellationToken) {
        var order = new List<string>();
        var values = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
        foreach (BibliographyIdentifier identifier in item.Identifiers) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!CodecMappings.IsCslIdentifierScheme(identifier.Scheme)) continue;
            string scheme = identifier.Scheme.ToUpperInvariant();
            if (!values.TryGetValue(scheme, out List<string>? group)) { group = new List<string>(); values.Add(scheme, group); order.Add(scheme); }
            group.Add(identifier.Value);
        }
        foreach (string scheme in order) {
            cancellationToken.ThrowIfCancellationRequested();
            WriteString(writer, scheme, JoinValues(values[scheme], "; ", cancellationToken), cancellationToken);
        }
    }

    private static string JoinValues(IEnumerable<string> values, string separator, CancellationToken cancellationToken) {
        var builder = new StringBuilder();
        foreach (string value in values) {
            cancellationToken.ThrowIfCancellationRequested();
            if (builder.Length > 0) builder.Append(separator);
            builder.Append(value);
        }
        return builder.ToString();
    }

    private static bool TryReadDatePart(JsonElement value, CancellationToken cancellationToken, out int? year, out int? month, out int? day) {
        year = null; month = null; day = null;
        List<JsonElement> parts = TakeBoundedElements(value, 3, cancellationToken, out bool tooManyParts);
        if (tooManyParts || parts.Count < 1) return false;
        var numbers = new int[parts.Count];
        for (int index = 0; index < parts.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (parts[index].ValueKind != JsonValueKind.Number || !parts[index].TryGetInt32(out numbers[index])) return false;
        }
        if (numbers.Length > 1 && (numbers[1] < 1 || numbers[1] > 12)) return false;
        if (numbers.Length > 2 && (numbers[2] < 1 || numbers[2] > 31)) return false;
        year = numbers[0];
        if (numbers.Length > 1) month = numbers[1];
        if (numbers.Length > 2) day = numbers[2];
        return true;
    }

    private static List<JsonElement> TakeBoundedElements(JsonElement value, int maximum, CancellationToken cancellationToken, out bool tooMany) {
        var elements = new List<JsonElement>(maximum);
        tooMany = false;
        if (value.ValueKind != JsonValueKind.Array) return elements;
        foreach (JsonElement element in value.EnumerateArray()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (elements.Count == maximum) { tooMany = true; break; }
            elements.Add(element);
        }
        cancellationToken.ThrowIfCancellationRequested();
        return elements;
    }

    private static void WriteDatePart(Utf8JsonWriter writer, int year, int? month, int? day) {
        writer.WriteStartArray(); writer.WriteNumberValue(year); if (month.HasValue) writer.WriteNumberValue(month.Value); if (day.HasValue) writer.WriteNumberValue(day.Value); writer.WriteEndArray();
    }

    private static bool TryWriteRaw(Utf8JsonWriter writer, string raw, CancellationToken cancellationToken) {
        try {
            cancellationToken.ThrowIfCancellationRequested();
            using JsonDocument value = JsonDocument.Parse(raw, NativeJsonDocumentOptions);
            cancellationToken.ThrowIfCancellationRequested();
            if (IsStrictNativeJson(raw, cancellationToken)) writer.WriteRawValue(raw, skipInputValidation: true);
            else value.RootElement.WriteTo(writer);
            cancellationToken.ThrowIfCancellationRequested();
            return true;
        } catch (JsonException) { return false; }
    }

    private static bool WouldBindTypedItemProperty(BibliographyNativeField field, CancellationToken cancellationToken) {
        using JsonDocument? output = GetNativeOutputJson(field, cancellationToken);
        JsonValueKind kind = output?.RootElement.ValueKind ?? JsonValueKind.String;
        if (IsTypedScalarProperty(field.Name)) {
            if (IsIdentifierProperty(field.Name) && kind == JsonValueKind.String) {
                string value = output == null ? field.Value : output.RootElement.GetString() ?? string.Empty;
                return !string.IsNullOrWhiteSpace(value);
            }
            return kind == JsonValueKind.String;
        }
        if (IsTypedNameProperty(field.Name)) return kind == JsonValueKind.Array && CanParseNames(output!.RootElement, cancellationToken);
        if (IsTypedDateProperty(field.Name)) return kind == JsonValueKind.Object;
        return false;
    }

    private static bool WouldBindTypedDateProperty(BibliographyNativeField field, CancellationToken cancellationToken) {
        using JsonDocument? output = GetNativeOutputJson(field, cancellationToken);
        JsonValueKind kind = output?.RootElement.ValueKind ?? JsonValueKind.String;
        if (string.Equals(field.Name, "literal", StringComparison.Ordinal)) return kind == JsonValueKind.String;
        return string.Equals(field.Name, "date-parts", StringComparison.Ordinal) && kind == JsonValueKind.Array && CanParseDateParts(output!.RootElement, cancellationToken);
    }

    private static JsonDocument? GetNativeOutputJson(BibliographyNativeField field, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        string? raw = field.UnmodifiedRawValue;
        if (raw != null) return TryParseNativeJson(raw, cancellationToken);
        JsonValueKind? originalKind = GetRawJsonKind(field.RawValue, cancellationToken);
        if (originalKind.HasValue && originalKind != JsonValueKind.String && originalKind != JsonValueKind.Null && originalKind != JsonValueKind.Undefined) {
            JsonDocument? edited = TryParseNativeJson(field.Value, cancellationToken);
            if (edited != null && edited.RootElement.ValueKind != JsonValueKind.String && edited.RootElement.ValueKind != JsonValueKind.Null && edited.RootElement.ValueKind != JsonValueKind.Undefined) return edited;
            edited?.Dispose();
        }
        return null;
    }

    private static JsonDocument? TryParseNativeJson(string raw, CancellationToken cancellationToken) {
        try {
            cancellationToken.ThrowIfCancellationRequested();
            JsonDocument result = JsonDocument.Parse(raw, NativeJsonDocumentOptions);
            if (cancellationToken.IsCancellationRequested) { result.Dispose(); cancellationToken.ThrowIfCancellationRequested(); }
            return result;
        }
        catch (JsonException) { return null; }
    }

    private static bool CanParseNames(JsonElement value, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (value.GetArrayLength() == 0) return false;
        foreach (JsonElement element in value.EnumerateArray()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (element.ValueKind != JsonValueKind.String && element.ValueKind != JsonValueKind.Object) return false;
            if (element.ValueKind == JsonValueKind.Object && HasWrongShapedKnownNameProperty(element, cancellationToken)) return false;
        }
        return true;
    }

    private static bool CanParseDateParts(JsonElement parts, CancellationToken cancellationToken) {
        List<JsonElement> ranges = TakeBoundedElements(parts, 2, cancellationToken, out bool tooManyRanges);
        if (tooManyRanges || ranges.Count < 1 || !TryReadDatePart(ranges[0], cancellationToken, out _, out _, out _)) return false;
        return ranges.Count != 2 || TryReadDatePart(ranges[1], cancellationToken, out _, out _, out _);
    }

    private static bool IsTypedScalarProperty(string name) {
        switch (name) {
            case "id": case "type": case "title": case "container-title": case "collection-title":
            case "publisher": case "publisher-place": case "edition": case "volume": case "issue": case "page":
            case "abstract": case "language": case "URL": case "DOI": case "ISBN": case "ISSN": case "PMID": case "PMCID":
            case "keyword": case "note": return true;
            default: return false;
        }
    }

    private static bool IsIdentifierProperty(string name) {
        switch (name) {
            case "DOI": case "ISBN": case "ISSN": case "PMID": case "PMCID": return true;
            default: return false;
        }
    }

    internal static bool CanPreserveNativeType(BibliographyFormat sourceFormat, BibliographyItem item) =>
        sourceFormat == BibliographyFormat.CslJson && !string.IsNullOrWhiteSpace(item.NativeType) && CodecMappings.ParseCslType(item.NativeType) == item.Type;

    internal static bool UsesNativeType(BibliographyFormat sourceFormat, BibliographyItem item) =>
        CanPreserveNativeType(sourceFormat, item) || item.Type == BibliographyItemType.Unknown && !string.IsNullOrWhiteSpace(item.NativeType);

    internal static bool CanRoundTripType(BibliographyFormat sourceFormat, BibliographyItem item) =>
        CodecMappings.ParseCslType(OutputType(sourceFormat, item)) == item.Type;

    internal static bool PreservesNativeType(BibliographyFormat sourceFormat, BibliographyItem item) =>
        sourceFormat != BibliographyFormat.CslJson ||
        item.NativeType == null ||
        string.Equals(OutputType(sourceFormat, item), item.NativeType, StringComparison.Ordinal);

    private static string OutputType(BibliographyFormat sourceFormat, BibliographyItem item) =>
        UsesNativeType(sourceFormat, item) ? item.NativeType! : CodecMappings.ToCslType(item.Type);

    private static bool IsTypedNameProperty(string name) {
        switch (name) {
            case "author": case "editor": case "translator": case "recipient": case "interviewer": case "composer": case "collection-editor": return true;
            default: return false;
        }
    }

    private static bool IsTypedDateProperty(string name) {
        switch (name) {
            case "issued": case "accessed": case "submitted": case "original-date": case "event-date": return true;
            default: return false;
        }
    }

    private static JsonDocument ParseDocument(string source, BibliographyReadOptions options, IList<BibliographyItem> partialItems, CancellationToken cancellationToken) {
        const int ChunkCharacters = 4096;
        using var stream = new MemoryStream(Math.Min(source.Length, 1024 * 1024));
        var encoder = new UTF8Encoding(false, true).GetEncoder();
        var characters = new char[Math.Min(ChunkCharacters, Math.Max(1, source.Length))];
        var bytes = new byte[Encoding.UTF8.GetMaxByteCount(characters.Length)];
        int position = 0;
        while (position < source.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            int characterCount = Math.Min(characters.Length, source.Length - position);
            source.CopyTo(position, characters, 0, characterCount);
            for (int index = 0; index < characterCount; index++) {
                int sourceIndex = position + index;
                if (characters[index] == '\r' && (sourceIndex + 1 >= source.Length || source[sourceIndex + 1] != '\n')) characters[index] = '\n';
            }
            bool flush = position + characterCount == source.Length;
            try {
                encoder.Convert(characters, 0, characterCount, bytes, 0, bytes.Length, flush, out int charactersUsed, out int bytesUsed, out _);
                stream.Write(bytes, 0, bytesUsed);
                position += charactersUsed;
            } catch (EncoderFallbackException exception) {
                throw new JsonException("CSL JSON input contains invalid UTF-16.", exception);
            }
        }
        cancellationToken.ThrowIfCancellationRequested();
        if (!stream.TryGetBuffer(out ArraySegment<byte> buffer) || buffer.Array == null) throw new InvalidOperationException("The CSL JSON input buffer is unavailable.");
        ValidateBeforeMaterialization(buffer.Array, buffer.Offset, checked((int)stream.Length), options, partialItems, cancellationToken);
        stream.Position = 0;
        return JsonDocument.ParseAsync(stream, new JsonDocumentOptions { AllowTrailingCommas = true, CommentHandling = JsonCommentHandling.Skip, MaxDepth = options.MaximumNestingDepth }, cancellationToken).GetAwaiter().GetResult();
    }

    private static void ValidateBeforeMaterialization(byte[] sourceBytes, int sourceOffset, int sourceLength, BibliographyReadOptions options, IList<BibliographyItem> partialItems, CancellationToken cancellationToken) {
        int utf16BaseOffset = 0;
        if (sourceLength >= 3 && sourceBytes[sourceOffset] == 0xEF && sourceBytes[sourceOffset + 1] == 0xBB && sourceBytes[sourceOffset + 2] == 0xBF) {
            sourceOffset += 3;
            sourceLength -= 3;
            utf16BaseOffset = 1;
        }
        ReadOnlySpan<byte> source = new ReadOnlySpan<byte>(sourceBytes, sourceOffset, sourceLength);
        var limits = new BibliographyLimitGuard(options);
        using var cancellationAwareSource = new CancellationAwareJsonSequence(sourceBytes, sourceOffset, sourceLength, cancellationToken);
        var reader = new Utf8JsonReader(cancellationAwareSource.Sequence, new JsonReaderOptions { AllowTrailingCommas = true, CommentHandling = JsonCommentHandling.Skip, MaxDepth = options.MaximumNestingDepth });
        var arrayContainers = new List<bool>();
        var containerOffsets = new List<int>();
        var containerUtf16Starts = new List<int>();
        var boundedContainers = new List<bool>();
        bool rootIsArray = false;
        int measuredByteOffset = 0;
        int measuredUtf16Offset = utf16BaseOffset;
        int tokenCount = 0;
        while (reader.Read()) {
            if ((tokenCount++ & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
            int tokenByteOffset = checked((int)reader.TokenStartIndex);
            AdvanceJsonUtf16Offset(source, tokenByteOffset, ref measuredByteOffset, ref measuredUtf16Offset, cancellationToken);
            int tokenUtf16Offset = measuredUtf16Offset;
            JsonTokenType token = reader.TokenType;
            bool parentIsArray = arrayContainers.Count > 0 && arrayContainers[arrayContainers.Count - 1];
            if (token == JsonTokenType.PropertyName) {
                limits.CheckValueLength(partialItems, GetJsonStringUtf16Length(ref reader, cancellationToken), tokenUtf16Offset);
                limits.AddValue(partialItems, null, tokenUtf16Offset);
                continue;
            }
            if (token == JsonTokenType.StartObject || token == JsonTokenType.StartArray || token == JsonTokenType.String || token == JsonTokenType.Number || token == JsonTokenType.True || token == JsonTokenType.False || token == JsonTokenType.Null) {
                bool rootArrayItemObject = rootIsArray && arrayContainers.Count == 1 && token == JsonTokenType.StartObject;
                if (parentIsArray && !rootArrayItemObject) limits.AddValue(partialItems, null, tokenUtf16Offset);
                if (token == JsonTokenType.StartObject && (arrayContainers.Count == 0 || rootIsArray && arrayContainers.Count == 1))
                    limits.AddItem(partialItems, tokenUtf16Offset);
            }
            if (token == JsonTokenType.String)
                limits.CheckValueLength(partialItems, GetJsonStringUtf16Length(ref reader, cancellationToken), tokenUtf16Offset);
            else if (token == JsonTokenType.Number || token == JsonTokenType.True || token == JsonTokenType.False || token == JsonTokenType.Null)
                limits.CheckValueLength(partialItems, checked((int)(reader.HasValueSequence ? reader.ValueSequence.Length : reader.ValueSpan.Length)), tokenUtf16Offset);
            if (token == JsonTokenType.StartArray || token == JsonTokenType.StartObject) {
                bool rootContainer = arrayContainers.Count == 0;
                bool itemContainer = token == JsonTokenType.StartObject && (rootContainer || rootIsArray && arrayContainers.Count == 1);
                containerOffsets.Add(tokenUtf16Offset);
                containerUtf16Starts.Add(measuredUtf16Offset);
                boundedContainers.Add(!rootContainer && !itemContainer);
            }
            if (token == JsonTokenType.StartArray) {
                if (arrayContainers.Count == 0) rootIsArray = true;
                arrayContainers.Add(true);
            } else if (token == JsonTokenType.StartObject) arrayContainers.Add(false);
            else if ((token == JsonTokenType.EndArray || token == JsonTokenType.EndObject) && arrayContainers.Count > 0) {
                int containerIndex = arrayContainers.Count - 1;
                if (boundedContainers[containerIndex]) {
                    int end = checked((int)reader.BytesConsumed);
                    AdvanceJsonUtf16Offset(source, end, ref measuredByteOffset, ref measuredUtf16Offset, cancellationToken);
                    limits.CheckValueLength(partialItems, measuredUtf16Offset - containerUtf16Starts[containerIndex], containerOffsets[containerIndex]);
                }
                arrayContainers.RemoveAt(containerIndex);
                containerOffsets.RemoveAt(containerIndex);
                containerUtf16Starts.RemoveAt(containerIndex);
                boundedContainers.RemoveAt(containerIndex);
            }
        }
        cancellationToken.ThrowIfCancellationRequested();
    }

    private static void AdvanceJsonUtf16Offset(ReadOnlySpan<byte> source, int targetByteOffset, ref int byteOffset, ref int utf16Offset, CancellationToken cancellationToken) {
        ReadOnlySpan<byte> value = source.Slice(byteOffset, targetByteOffset - byteOffset);
        int length = 0;
        int nextCancellationCheck = 0;
        for (int index = 0; index < value.Length;) {
            if (index >= nextCancellationCheck) { cancellationToken.ThrowIfCancellationRequested(); nextCancellationCheck = index + 4096; }
            byte current = value[index];
            if (current < 0x80) { length++; index++; }
            else if (current < 0xE0) { length++; index += 2; }
            else if (current < 0xF0) { length++; index += 3; }
            else { length += 2; index += 4; }
        }
        utf16Offset += length;
        byteOffset = targetByteOffset;
    }

    private static int GetJsonStringUtf16Length(ref Utf8JsonReader reader, CancellationToken cancellationToken) {
        if (!reader.HasValueSequence) return GetJsonStringUtf16Length(reader.ValueSpan, reader.ValueIsEscaped, cancellationToken);
        int length = 0;
        int processed = 0;
        int utf8ContinuationBytes = 0;
        int unicodeEscapeDigits = 0;
        bool escapeCodePending = false;
        foreach (ReadOnlyMemory<byte> memory in reader.ValueSequence) {
            ReadOnlySpan<byte> segment = memory.Span;
            for (int index = 0; index < segment.Length; index++) {
                if ((processed++ & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                byte current = segment[index];
                if (utf8ContinuationBytes > 0) { utf8ContinuationBytes--; continue; }
                if (unicodeEscapeDigits > 0) { unicodeEscapeDigits--; continue; }
                if (escapeCodePending) {
                    escapeCodePending = false;
                    if (current == (byte)'u') unicodeEscapeDigits = 4;
                    continue;
                }
                if (reader.ValueIsEscaped && current == (byte)'\\') {
                    length++;
                    escapeCodePending = true;
                } else if (current < 0x80) length++;
                else if (current < 0xE0) { length++; utf8ContinuationBytes = 1; }
                else if (current < 0xF0) { length++; utf8ContinuationBytes = 2; }
                else { length += 2; utf8ContinuationBytes = 3; }
            }
        }
        cancellationToken.ThrowIfCancellationRequested();
        return length;
    }

    private static int GetJsonStringUtf16Length(ReadOnlySpan<byte> value, bool escaped, CancellationToken cancellationToken) {
        int length = 0;
        for (int index = 0; index < value.Length;) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            byte current = value[index];
            if (escaped && current == (byte)'\\') {
                length++;
                index += index + 1 < value.Length && value[index + 1] == (byte)'u' ? 6 : 2;
            } else if (current < 0x80) {
                length++;
                index++;
            } else if (current < 0xE0) {
                length++;
                index += 2;
            } else if (current < 0xF0) {
                length++;
                index += 3;
            } else {
                length += 2;
                index += 4;
            }
        }
        cancellationToken.ThrowIfCancellationRequested();
        return length;
    }

    private sealed class CancellationAwareJsonSequence : IDisposable {
        private const int SegmentSize = 4096;
        private readonly List<CancellationAwareMemoryManager> _managers = new List<CancellationAwareMemoryManager>();

        internal CancellationAwareJsonSequence(byte[] source, int offset, int length, CancellationToken cancellationToken) {
            if (length == 0) { Sequence = ReadOnlySequence<byte>.Empty; return; }
            JsonSequenceSegment? first = null;
            JsonSequenceSegment? last = null;
            int end = checked(offset + length);
            for (int position = offset; position < end;) {
                cancellationToken.ThrowIfCancellationRequested();
                int count = Math.Min(SegmentSize, end - position);
                var manager = new CancellationAwareMemoryManager(source, position, count, cancellationToken);
                _managers.Add(manager);
                if (first == null) first = last = new JsonSequenceSegment(manager.Memory);
                else last = last!.Append(manager.Memory);
                position += count;
            }
            Sequence = new ReadOnlySequence<byte>(first!, 0, last!, last!.Memory.Length);
        }

        internal ReadOnlySequence<byte> Sequence { get; }

        public void Dispose() {
            foreach (CancellationAwareMemoryManager manager in _managers) ((IDisposable)manager).Dispose();
        }
    }

    private sealed class JsonSequenceSegment : ReadOnlySequenceSegment<byte> {
        internal JsonSequenceSegment(ReadOnlyMemory<byte> memory) => Memory = memory;
        internal JsonSequenceSegment Append(ReadOnlyMemory<byte> memory) {
            var segment = new JsonSequenceSegment(memory) { RunningIndex = RunningIndex + Memory.Length };
            Next = segment;
            return segment;
        }
    }

    private sealed class CancellationAwareMemoryManager : MemoryManager<byte> {
        private readonly byte[] _source;
        private readonly int _offset;
        private readonly int _length;
        private readonly CancellationToken _cancellationToken;

        internal CancellationAwareMemoryManager(byte[] source, int offset, int length, CancellationToken cancellationToken) {
            _source = source;
            _offset = offset;
            _length = length;
            _cancellationToken = cancellationToken;
        }

        public override Span<byte> GetSpan() {
            _cancellationToken.ThrowIfCancellationRequested();
            return new Span<byte>(_source, _offset, _length);
        }

        public override MemoryHandle Pin(int elementIndex = 0) => throw new NotSupportedException();
        public override void Unpin() { }
        protected override void Dispose(bool disposing) { }
    }
    private static bool WriteNativeValue(Utf8JsonWriter writer, BibliographyNativeField field, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        string? raw = field.UnmodifiedRawValue;
        if (raw != null) {
            if (TryWriteRaw(writer, raw, cancellationToken)) { cancellationToken.ThrowIfCancellationRequested(); return true; }
            writer.WriteStringValue(SanitizeUtf16(field.Value, cancellationToken));
            return false;
        }
        JsonValueKind? originalKind = GetRawJsonKind(field.RawValue, cancellationToken);
        if (originalKind == JsonValueKind.String || field.RawValue == null) {
            writer.WriteStringValue(SanitizeUtf16(field.Value, cancellationToken));
            return true;
        }
        if (originalKind.HasValue && TryWriteEditedRaw(writer, field.Value, cancellationToken)) { cancellationToken.ThrowIfCancellationRequested(); return true; }
        writer.WriteStringValue(SanitizeUtf16(field.Value, cancellationToken));
        return false;
    }

    private static JsonValueKind? GetRawJsonKind(string? raw, CancellationToken cancellationToken) {
        if (raw == null) return null;
        try { cancellationToken.ThrowIfCancellationRequested(); using JsonDocument value = JsonDocument.Parse(raw, NativeJsonDocumentOptions); cancellationToken.ThrowIfCancellationRequested(); return value.RootElement.ValueKind; }
        catch (JsonException) { return null; }
    }

    private static bool TryWriteEditedRaw(Utf8JsonWriter writer, string value, CancellationToken cancellationToken) {
        try {
            cancellationToken.ThrowIfCancellationRequested();
            using JsonDocument parsed = JsonDocument.Parse(value, NativeJsonDocumentOptions);
            cancellationToken.ThrowIfCancellationRequested();
            if (parsed.RootElement.ValueKind == JsonValueKind.String || parsed.RootElement.ValueKind == JsonValueKind.Null || parsed.RootElement.ValueKind == JsonValueKind.Undefined) return false;
            if (IsStrictNativeJson(value, cancellationToken)) writer.WriteRawValue(value, skipInputValidation: true);
            else parsed.RootElement.WriteTo(writer);
            cancellationToken.ThrowIfCancellationRequested();
            return true;
        } catch (JsonException) {
            return false;
        }
    }

    private static bool IsStrictNativeJson(string value, CancellationToken cancellationToken) {
        try {
            cancellationToken.ThrowIfCancellationRequested();
            using JsonDocument parsed = JsonDocument.Parse(value, new JsonDocumentOptions { MaxDepth = NativeJsonMaximumDepth });
            cancellationToken.ThrowIfCancellationRequested();
            return true;
        } catch (JsonException) {
            return false;
        }
    }

    private static string GetBoundedRawValue(JsonElement value, IList<BibliographyItem> items, BibliographyLimitGuard limits) {
        if (value.ValueKind == JsonValueKind.String) limits.CheckValueLength(items, value.GetString() ?? string.Empty, 0);
        string raw = value.GetRawText();
        if (value.ValueKind != JsonValueKind.String) limits.CheckValueLength(items, raw, 0);
        return raw;
    }

    internal static void GetJsonLocation(string source, JsonException exception, CancellationToken cancellationToken, out int offset, out int line, out int column) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!exception.LineNumber.HasValue || !exception.BytePositionInLine.HasValue) { offset = -1; line = -1; column = -1; return; }
        int zeroBasedLine = checked((int)exception.LineNumber.Value);
        int lineStart = 0;
        for (int currentLine = 0; currentLine < zeroBasedLine && lineStart < source.Length; currentLine++) {
            cancellationToken.ThrowIfCancellationRequested();
            while (lineStart < source.Length && source[lineStart] != '\r' && source[lineStart] != '\n') {
                if ((lineStart & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                lineStart++;
            }
            if (lineStart >= source.Length) break;
            if (source[lineStart++] == '\r' && lineStart < source.Length && source[lineStart] == '\n') lineStart++;
        }
        int lineEnd = lineStart;
        while (lineEnd < source.Length && source[lineEnd] != '\r' && source[lineEnd] != '\n') {
            if ((lineEnd & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            lineEnd++;
        }
        long bytePosition = exception.BytePositionInLine.Value;
        int characterCount = 0;
        while (lineStart + characterCount < lineEnd) {
            if ((characterCount & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
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
        cancellationToken.ThrowIfCancellationRequested();
    }
    internal static bool ContainsInvalidUtf16(string value, CancellationToken cancellationToken) {
        for (int index = 0; index < value.Length; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (char.IsHighSurrogate(value[index])) {
                if (index + 1 < value.Length && char.IsLowSurrogate(value[index + 1])) { index++; continue; }
                return true;
            }
            if (char.IsLowSurrogate(value[index])) return true;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return false;
    }
    private static string SanitizeUtf16(string value, CancellationToken cancellationToken) {
        if (!ContainsInvalidUtf16(value, cancellationToken)) return value;
        var builder = new StringBuilder(value.Length);
        for (int index = 0; index < value.Length; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            char current = value[index];
            if (char.IsHighSurrogate(current) && index + 1 < value.Length && char.IsLowSurrogate(value[index + 1])) builder.Append(current).Append(value[++index]);
            else builder.Append(char.IsSurrogate(current) ? '\uFFFD' : current);
        }
        cancellationToken.ThrowIfCancellationRequested();
        return builder.ToString();
    }
    private static void WritePropertyName(Utf8JsonWriter writer, string value, CancellationToken cancellationToken) => writer.WritePropertyName(SanitizeUtf16(value, cancellationToken));
    private static string NormalizeLineEndings(string value, string lineEnding) => value.Replace("\r\n", "\n").Replace("\r", "\n").Replace("\n", lineEnding);
}
