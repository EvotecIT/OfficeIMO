using System.Text.Json;
using System.Xml.Linq;

namespace OfficeIMO.Bibliography;

internal sealed partial class CslEvaluator {
    private CslText Date(XElement element, CslContext context) {
        string variable = Attr(element, "variable") ?? string.Empty;
        if (context.Suppressed.Contains(variable)) return CslText.Literal(string.Empty, true);
        JsonElement date = context.Record.Value(variable);
        if (date.ValueKind != JsonValueKind.Object) return CslText.Literal(string.Empty, true);
        string literal = Scalar(date, "literal");
        if (literal.Length > 0) return CslText.Rich(literal, cancellationToken: _token);
        int?[] start;
        int?[]? end = null;
        if (date.TryGetProperty("date-parts", out JsonElement parts) && parts.ValueKind == JsonValueKind.Array && parts.GetArrayLength() > 0) {
            start = DateParts(parts[0]);
            end = parts.GetArrayLength() > 1 ? DateParts(parts[1]) : null;
        } else {
            string raw = Scalar(date, "raw");
            if (raw.Length == 0) return CslText.Literal(string.Empty, true);
            string[] range = raw.Split('/');
            if (!TryRawDate(range[0], out start)) return CslText.Rich(raw, cancellationToken: _token);
            if (range.Length == 2 && TryRawDate(range[1], out int?[] ending)) end = ending;
        }
        if (!start[1].HasValue && date.TryGetProperty("season", out JsonElement season) &&
            int.TryParse(season.ValueKind == JsonValueKind.String ? season.GetString() : season.GetRawText(), out int seasonNumber) && seasonNumber >= 1 && seasonNumber <= 4) start[1] = seasonNumber + 12;
        XElement? localized = Attr(element, "form") == null ? null : _locale.Date(Attr(element, "form")!);
        XElement[] definitions = (localized ?? element).Elements(CslStyle.Namespace + "date-part").Select(part => new XElement(part)).ToArray();
        string requestedParts = Attr(element, "date-parts") ?? "year-month-day";
        definitions = definitions.Where(part => requestedParts.Split('-').Contains(Attr(part, "name"))).ToArray();
        if (context.Sorting) {
            string[] included = definitions.Length == 0 ? new[] { "year", "month", "day" } : definitions.Select(part => Attr(part, "name") ?? string.Empty).ToArray();
            string key = SortDate(start, included) + (end == null ? "0" : "1" + SortDate(end, included));
            return CslText.Literal(key, true);
        }
        if (context.SuppressYear) return CslText.Literal(context.AutomaticYearSuffix ? context.Record.YearSuffix : string.Empty, true);
        if (localized != null) {
            foreach (XElement part in definitions) {
                XElement? custom = element.Elements(CslStyle.Namespace + "date-part").FirstOrDefault(candidate => Attr(candidate, "name") == Attr(part, "name"));
                if (custom == null) continue;
                foreach (XAttribute attribute in custom.Attributes()) part.SetAttributeValue(attribute.Name, attribute.Value);
            }
        }
        string delimiter = Attr(element, "delimiter") ?? Attr(localized ?? element, "delimiter") ?? string.Empty;
        bool rendersYear = definitions.Any(part => Attr(part, "name") == "year");
        bool suffixAvailable = context.AutomaticYearSuffix && !context.YearRendered;
        bool suffixOnStart = suffixAvailable && start[0].HasValue;
        bool suffixOnEnd = suffixAvailable && !start[0].HasValue;
        CslText[] startParts = definitions.Select(part => DatePart(part, start, context, suffixOnStart)).ToArray();
        CslText output;
        if (end == null) output = Join(startParts, delimiter);
        else {
            CslText[] endParts = definitions.Select(part => DatePart(part, end, context, suffixOnEnd)).ToArray();
            int first = -1, last = -1;
            for (int index = 0; index < definitions.Length; index++) {
                int dateIndex = DatePartIndex(definitions[index]);
                if (start[dateIndex] == end[dateIndex]) continue;
                if (first < 0) first = index;
                last = index;
            }
            if (first < 0) output = Join(startParts, delimiter);
            else {
                var closingStartPart = new XElement(definitions[last]);
                closingStartPart.Attribute("suffix")?.Remove();
                startParts[last] = DatePart(closingStartPart, start, context, suffixOnStart);
                int largest = Enumerable.Range(first, last - first + 1).OrderBy(index => Attr(definitions[index], "name") == "year" ? 0 : Attr(definitions[index], "name") == "month" ? 1 : 2).First();
                string range = Attr(definitions[largest], "range-delimiter") ?? _locale.Term("year-range-delimiter");
                if (range.Length == 0) range = "–";
                CslText ranged = Join(new[] { Join(startParts.Skip(first).Take(last - first + 1), delimiter), Join(endParts.Skip(first).Take(last - first + 1), delimiter) }, range);
                output = Join(new[] { Join(startParts.Take(first), delimiter), ranged, Join(startParts.Skip(last + 1), delimiter) }, delimiter);
            }
        }
        if (rendersYear && (start[0].HasValue || end != null && end[0].HasValue)) context.YearRendered = true;
        return new CslText(output.Plain, output.Html, 1, output.IsEmpty ? 0 : 1);
    }

    private static int DatePartIndex(XElement part) => Attr(part, "name") == "month" ? 1 : Attr(part, "name") == "day" ? 2 : 0;

    private CslText DatePart(XElement part, int?[] values, CslContext context, bool appendYearSuffix) {
        string name = Attr(part, "name") ?? "year";
        int index = name == "year" ? 0 : name == "month" ? 1 : 2;
        int? number = values[index];
        if (!number.HasValue) return CslText.Empty;
        string form = Attr(part, "form") ?? "long";
        string value;
        if (name == "month") {
            if (number >= 13 && number <= 16) value = _locale.Term("season-" + (number - 12)!.Value.ToString("00", CultureInfo.InvariantCulture), form == "short" ? "short" : "long");
            else if (number >= 1 && number <= 12 && (form == "long" || form == "short")) value = _locale.Term("month-" + number.Value.ToString("00", CultureInfo.InvariantCulture), form);
            else value = number.Value.ToString(form == "numeric-leading-zeros" ? "00" : "0", CultureInfo.InvariantCulture);
        } else if (name == "day") {
            bool ordinal = form == "ordinal" && (!_locale.Option("limit-day-ordinals-to-day-1") || number == 1);
            string? gender = values[1] >= 1 && values[1] <= 12 ?
                _locale.Gender("month-" + values[1]!.Value.ToString("00", CultureInfo.InvariantCulture)) : null;
            value = ordinal ? _locale.Ordinal(number.Value, gender: gender) : number.Value.ToString(form == "numeric-leading-zeros" ? "00" : "0", CultureInfo.InvariantCulture);
        } else {
            value = Math.Abs((long)number.Value).ToString(form == "short" ? "00" : "0", CultureInfo.InvariantCulture);
            if (form == "short" && value.Length > 2) value = value.Substring(value.Length - 2);
            if (number < 0) value += _locale.Term("bc");
            else if (number < 1000) value += _locale.Term("ad");
            if (appendYearSuffix) value += context.Record.YearSuffix;
        }
        return CslText.Literal(value).Decorate(part, _locale, context.Sorting, cancellationToken: _token);
    }

    private static int?[] DateParts(JsonElement data) {
        var result = new int?[3];
        if (data.ValueKind == JsonValueKind.Array) {
            for (int index = 0; index < Math.Min(3, data.GetArrayLength()); index++) {
                if (data[index].ValueKind == JsonValueKind.Number && data[index].TryGetInt32(out int number)) result[index] = number;
                else if (data[index].ValueKind == JsonValueKind.String && int.TryParse(data[index].GetString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out number)) result[index] = number;
            }
        }
        return result;
    }

    private static string SortDate(int?[] parts, string[] included) =>
        (included.Contains("year") && parts[0].HasValue ? ((long)parts[0]!.Value + int.MaxValue).ToString("D10", CultureInfo.InvariantCulture) : "0000000000") +
        (included.Contains("month") && parts[1] >= 1 && parts[1] <= 12 ? parts[1]!.Value.ToString("D2", CultureInfo.InvariantCulture) : "00") +
        (included.Contains("day") && parts[2].HasValue ? parts[2]!.Value.ToString("D2", CultureInfo.InvariantCulture) : "00");

    private bool TryRawDate(string source, out int?[] result) {
        result = new int?[3];
        _token.ThrowIfCancellationRequested();
        // Preserve the previous anchored date grammar, including its optional final LF.
        int end = source.Length;
        if (end > 0 && source[end - 1] == '\n') end--;
        int position = 0;
        for (int index = 0; index < 3; index++) {
            int start = position;
            if (index == 0 && position < end && source[position] == '-') position++;
            int digits = position;
            while (position < end && char.IsDigit(source[position])) { CslNumberSyntax.Check(position, _token); position++; }
            if (position == digits || index > 0 && position - digits > 2) return false;
            if (int.TryParse(source.Substring(start, position - start), NumberStyles.Integer, CultureInfo.InvariantCulture, out int number)) result[index] = number;
            if (position == end) return result[0].HasValue;
            if (source[position++] != '-') return false;
        }
        return false;
    }
}
