using System.Globalization;
using System.Text.Json;
using OfficeIMO.Reader;

namespace OfficeIMO.AI;

public sealed partial class OfficeAiEngine {
    private sealed record ParsedBatch(IReadOnlyList<OfficeAiClaim> Claims, IReadOnlyList<OfficeAiField> Fields,
        IReadOnlyList<OfficeAiBlock> Blocks, IReadOnlyList<OfficeAiTable> Tables);

    private static ParsedBatch Parse(OfficeAiExecutionResponse response, Batch batch, OfficeAiRequest request) {
        if (response is null || !response.IsComplete || string.IsNullOrWhiteSpace(response.Json)
            || response.Json.Length > request.Limits.MaxResponseCharacters)
            throw Invalid();
        try {
            using JsonDocument json = JsonDocument.Parse(response.Json, new JsonDocumentOptions { MaxDepth = 12 });
            JsonElement root = json.RootElement;
            CheckObject(root, "status", "claims", "fields", "blocks", "tables");
            string status = Text(root.GetProperty("status"));
            if (status is not ("ok" or "insufficient")) throw Invalid();
            int maximum = Math.Min(200, request.Limits.MaxResultItems);
            JsonElement[] claimItems = Items(root.GetProperty("claims"), maximum);
            JsonElement[] fieldItems = Items(root.GetProperty("fields"), maximum);
            JsonElement[] blockItems = Items(root.GetProperty("blocks"), maximum);
            JsonElement[] tableItems = Items(root.GetProperty("tables"), maximum);
            bool reasoning = request.Operation is OfficeAiOperation.Ask or OfficeAiOperation.Explain or OfficeAiOperation.Summarize;
            if ((!reasoning && claimItems.Length != 0) || (request.Operation != OfficeAiOperation.ExtractFields && fieldItems.Length != 0)
                || (request.Operation != OfficeAiOperation.Parse && (blockItems.Length != 0 || tableItems.Length != 0))) throw Invalid();
            if (status == "insufficient" && (claimItems.Length + blockItems.Length + tableItems.Length > 0)) throw Invalid();
            var claims = new List<OfficeAiClaim>();
            foreach (JsonElement item in claimItems) {
                CheckObject(item, "text", "evidence");
                claims.Add(new(Text(item.GetProperty("text")), Citations(item.GetProperty("evidence"), batch, required: true)));
            }
            var fields = new List<OfficeAiField>();
            var definitions = request.Fields.ToDictionary(field => field.Name, StringComparer.Ordinal);
            foreach (JsonElement item in fieldItems) {
                CheckObject(item, "name", "status", "rawValue", "evidence");
                string name = Text(item.GetProperty("name"));
                if (!definitions.Remove(name, out OfficeAiFieldDefinition? definition)) throw Invalid();
                OfficeAiFieldStatus fieldStatus = Text(item.GetProperty("status")) switch {
                    "present" => OfficeAiFieldStatus.Present, "missing" => OfficeAiFieldStatus.Missing,
                    "ambiguous" => OfficeAiFieldStatus.Ambiguous, "conflicting" => OfficeAiFieldStatus.Conflicting, _ => throw Invalid()
                };
                JsonElement rawElement = item.GetProperty("rawValue");
                string? raw = rawElement.ValueKind == JsonValueKind.Null ? null : Text(rawElement);
                IReadOnlyList<OfficeAiCitation> citations = Citations(item.GetProperty("evidence"), batch, fieldStatus != OfficeAiFieldStatus.Missing);
                if (fieldStatus == OfficeAiFieldStatus.Missing && (raw is not null || citations.Count != 0)) throw Invalid();
                if (fieldStatus == OfficeAiFieldStatus.Present && raw is null) throw Invalid();
                if (status == "insufficient" && fieldStatus != OfficeAiFieldStatus.Missing) throw Invalid();
                // A text-only field value must actually occur in its quoted source. Visual readings stay review candidates.
                if (raw is not null && !citations.Any(citation => citation.Quote?.Contains(raw, StringComparison.Ordinal) == true
                    || batch.Images.ContainsKey(citation.EvidenceId))) throw Invalid();
                string? normalized = null;
                if (fieldStatus == OfficeAiFieldStatus.Present && !TryNormalize(raw!, definition, request.Culture, out normalized))
                    fieldStatus = OfficeAiFieldStatus.Invalid;
                fields.Add(new(name, definition.Type, fieldStatus, raw, normalized, citations));
            }
            if (definitions.Count != 0) throw Invalid();
            var blocks = new List<OfficeAiBlock>();
            foreach (JsonElement item in blockItems) {
                CheckObject(item, "kind", "text", "evidence");
                string kind = Text(item.GetProperty("kind"));
                if (kind is not ("heading" or "paragraph" or "list-item" or "caption")) throw Invalid();
                IReadOnlyList<OfficeAiCitation> citations = Citations(item.GetProperty("evidence"), batch, required: true);
                blocks.Add(new(new OfficeDocumentBlock {
                    Id = batch.Request.RequestId + "-block-" + (blocks.Count + 1), Kind = kind,
                    Text = Text(item.GetProperty("text")), Location = new ReaderLocation { Page = CommonPage(citations) }
                }, citations));
            }
            var tables = new List<OfficeAiTable>();
            foreach (JsonElement item in tableItems) {
                CheckObject(item, "title", "columns", "rows", "evidence");
                string title = Text(item.GetProperty("title"), allowEmpty: true);
                string[] columns = Items(item.GetProperty("columns"), 100).Select(value => Text(value, allowEmpty: true)).ToArray();
                if (columns.Length == 0) throw Invalid();
                JsonElement[] rows = Items(item.GetProperty("rows"), maximum);
                if ((long)rows.Length * columns.Length > request.Limits.MaxTableCells) throw Invalid();
                var values = new List<IReadOnlyList<string>>();
                foreach (JsonElement row in rows) {
                    string[] cells = Items(row, 100).Select(value => Text(value, allowEmpty: true)).ToArray();
                    if (cells.Length != columns.Length) throw Invalid();
                    values.Add(Array.AsReadOnly(cells));
                }
                IReadOnlyList<OfficeAiCitation> citations = Citations(item.GetProperty("evidence"), batch, required: true);
                tables.Add(new(new ReaderTable { Title = title, Kind = "ai-proposed", Columns = Array.AsReadOnly(columns),
                    Rows = values.AsReadOnly(), Location = new ReaderLocation { Page = CommonPage(citations) } }, citations));
            }
            return new(claims.AsReadOnly(), fields.AsReadOnly(), blocks.AsReadOnly(), tables.AsReadOnly());
        } catch (JsonException) { throw Invalid(); }
          catch (InvalidOperationException) { throw Invalid(); }
          catch (KeyNotFoundException) { throw Invalid(); }
    }

    private static IReadOnlyList<OfficeAiCitation> Citations(JsonElement element, Batch batch, bool required) {
        JsonElement[] items = Items(element, 32);
        if (required && items.Length == 0) throw Invalid();
        var citations = new List<OfficeAiCitation>();
        var unique = new HashSet<(string, string?)>();
        foreach (JsonElement item in items) {
            CheckObject(item, "id", "quote");
            string id = Text(item.GetProperty("id"));
            JsonElement quoteElement = item.GetProperty("quote");
            string? quote = quoteElement.ValueKind == JsonValueKind.Null ? null : Text(quoteElement);
            if (!unique.Add((id, quote))) throw Invalid();
            if (batch.Evidence.TryGetValue(id, out OfficeAiEvidence? observation)) {
                if (quote is null || !observation.Text.Contains(quote, StringComparison.Ordinal)) throw Invalid();
                citations.Add(new(id, observation.Page, quote, true));
            } else if (batch.Images.TryGetValue(id, out OfficeAiImage? image)) {
                if (quote is not null) throw Invalid();
                citations.Add(new(id, image.Page, null, false));
            } else throw Invalid();
        }
        return citations.AsReadOnly();
    }

    private static int? CommonPage(IReadOnlyList<OfficeAiCitation> citations) {
        int? first = citations[0].Page;
        return citations.All(citation => citation.Page == first) ? first : null;
    }

    private static void CheckObject(JsonElement item, params string[] names) {
        if (item.ValueKind != JsonValueKind.Object) throw Invalid();
        var expected = names.ToHashSet(StringComparer.Ordinal);
        foreach (JsonProperty property in item.EnumerateObject()) if (!expected.Remove(property.Name)) throw Invalid();
        if (expected.Count != 0) throw Invalid();
    }

    private static JsonElement[] Items(JsonElement item, int maximum) {
        if (item.ValueKind != JsonValueKind.Array || item.GetArrayLength() > maximum) throw Invalid();
        return item.EnumerateArray().ToArray();
    }

    private static string Text(JsonElement element, bool allowEmpty = false) {
        if (element.ValueKind != JsonValueKind.String) throw Invalid();
        string value = element.GetString()!;
        if ((!allowEmpty && string.IsNullOrWhiteSpace(value)) || value.Length > 32_000 || value.Contains('\0')) throw Invalid();
        return value;
    }

    private static InvalidDataException Invalid() => new("The provider response does not satisfy the document result contract.");

    private static bool TryNormalize(string raw, OfficeAiFieldDefinition definition, string cultureName, out string? normalized) {
        CultureInfo culture = CultureInfo.GetCultureInfo(cultureName);
        normalized = null;
        switch (definition.Type) {
            case OfficeAiFieldType.String: normalized = raw; return true;
            case OfficeAiFieldType.Decimal:
                if (!ValidGrouping(raw, culture.NumberFormat) || !decimal.TryParse(raw,
                    NumberStyles.AllowLeadingWhite | NumberStyles.AllowTrailingWhite | NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint | NumberStyles.AllowThousands,
                    culture, out decimal number)) return false;
                normalized = number.ToString(CultureInfo.InvariantCulture); return true;
            case OfficeAiFieldType.Integer:
                if (!long.TryParse(raw, NumberStyles.Integer, culture, out long integer)) return false;
                normalized = integer.ToString(CultureInfo.InvariantCulture); return true;
            case OfficeAiFieldType.Boolean:
                if (!bool.TryParse(raw, out bool boolean)) return false;
                normalized = boolean ? "true" : "false"; return true;
            case OfficeAiFieldType.Date:
                if (!DateOnly.TryParseExact(raw, definition.DateFormat, culture, DateTimeStyles.None, out DateOnly date)) return false;
                normalized = date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture); return true;
            default: return false;
        }
    }

    private static bool ValidGrouping(string raw, NumberFormatInfo format) {
        string value = raw.Trim();
        if (value.StartsWith(format.NegativeSign, StringComparison.Ordinal)) value = value[format.NegativeSign.Length..];
        else if (format.PositiveSign.Length > 0 && value.StartsWith(format.PositiveSign, StringComparison.Ordinal)) value = value[format.PositiveSign.Length..];
        string separator = format.NumberGroupSeparator;
        // .NET accepts an ordinary space for a locale's non-breaking-space group separator.
        if (separator == "\u00a0" || separator == "\u202f") value = value.Replace(" ", separator, StringComparison.Ordinal);
        string integer = value.Split(format.NumberDecimalSeparator, StringSplitOptions.None)[0];
        if (separator.Length == 0 || !integer.Contains(separator, StringComparison.Ordinal)) return true;
        string[] groups = integer.Split(separator, StringSplitOptions.None);
        int[] sizes = format.NumberGroupSizes;
        int sizeIndex = 0;
        for (int index = groups.Length - 1; index > 0; index--) {
            int required = sizes.Length == 0 ? 0 : sizes[sizeIndex];
            if (required == 0 || groups[index].Length != required) return false;
            if (sizeIndex + 1 < sizes.Length) sizeIndex++;
        }
        int leadingMaximum = sizes.Length == 0 ? 0 : sizes[sizeIndex];
        return groups[0].Length > 0 && (leadingMaximum == 0 || groups[0].Length <= leadingMaximum);
    }

    private static IReadOnlyList<OfficeAiField> MergeFields(List<OfficeAiField> fields, IReadOnlyList<OfficeAiFieldDefinition> definitions) {
        var results = new List<OfficeAiField>();
        foreach (OfficeAiFieldDefinition definition in definitions) {
            OfficeAiField[] candidates = fields.Where(field => field.Name == definition.Name && field.Status != OfficeAiFieldStatus.Missing).ToArray();
            if (candidates.Length == 0) {
                results.Add(new(definition.Name, definition.Type, OfficeAiFieldStatus.Missing, null, null, Array.Empty<OfficeAiCitation>()));
                continue;
            }
            OfficeAiField first = candidates[0];
            bool conflict = candidates.Any(field => field.Status == OfficeAiFieldStatus.Conflicting)
                || candidates.Where(field => field.RawValue is not null).Select(field => field.NormalizedValue ?? field.RawValue).Distinct(StringComparer.Ordinal).Count() > 1;
            OfficeAiFieldStatus status = conflict ? OfficeAiFieldStatus.Conflicting
                : candidates.Any(field => field.Status == OfficeAiFieldStatus.Ambiguous) ? OfficeAiFieldStatus.Ambiguous
                : candidates.Any(field => field.Status == OfficeAiFieldStatus.Invalid) ? OfficeAiFieldStatus.Invalid : first.Status;
            results.Add(first with { Status = status, NormalizedValue = status == OfficeAiFieldStatus.Present ? first.NormalizedValue : null,
                RawValue = conflict ? null : first.RawValue,
                Citations = Array.AsReadOnly(candidates.SelectMany(field => field.Citations).Distinct().ToArray()) });
        }
        return results.AsReadOnly();
    }
}
