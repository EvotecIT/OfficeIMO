using System.Text.Json;
using System.Xml.Linq;

namespace OfficeIMO.Bibliography;

public sealed partial class CslProcessor {
    private CslRecord[] Sort(CslRecord[] records, XElement section, CslEvaluator evaluator, XElementScope scope) {
        XElement[] keys = section.Element(CslStyle.Namespace + "sort")?.Elements(CslStyle.Namespace + "key").ToArray() ?? Array.Empty<XElement>();
        if (keys.Length == 0) return records;
        var values = records.Select((record, index) => new SortRow(record, index, keys.Select(key => SortKey(record, key, evaluator, scope)).ToArray())).ToArray();
        CultureInfo culture;
        try { culture = CultureInfo.GetCultureInfo(_options.Locale ?? _style.DefaultLocale); } catch (CultureNotFoundException) { culture = CultureInfo.InvariantCulture; }
        try { Array.Sort(values, (left, right) => {
            evaluator.PerformOperation();
            for (int index = 0; index < keys.Length; index++) {
                string a = left.Keys[index], b = right.Keys[index];
                if (a.Length == 0 && b.Length > 0) return 1;
                if (b.Length == 0 && a.Length > 0) return -1;
                int order = NaturalCompare(a, b, culture, evaluator);
                if (order != 0) return (string?)keys[index].Attribute("sort") == "descending" ? -order : order;
            }
            return left.Index.CompareTo(right.Index);
        }); } catch (InvalidOperationException exception) when (exception.InnerException is OperationCanceledException || exception.InnerException is InvalidDataException) {
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(exception.InnerException!).Throw();
            throw;
        }
        return values.Select(value => value.Record).ToArray();
    }

    private string SortKey(CslRecord record, XElement key, CslEvaluator evaluator, XElementScope scope) {
        var context = new CslContext(record, scope) { Sorting = true,
            SortNamesMinimum = ParseSortCount(key, "names-min"), SortNamesFirst = ParseSortCount(key, "names-use-first"),
            SortNamesLast = key.Attribute("names-use-last") == null ? null : (string?)key.Attribute("names-use-last") == "true" };
        string? macro = (string?)key.Attribute("macro");
        if (macro != null) return evaluator.Evaluate(_style.Macros[macro], context).Plain;
        string variable = (string?)key.Attribute("variable") ?? string.Empty;
        JsonElement value = record.Value(variable);
        if (value.ValueKind == JsonValueKind.Array) return evaluator.Evaluate(new XElement(CslStyle.Namespace + "names", new XAttribute("variable", variable),
            new XElement(CslStyle.Namespace + "name", new XAttribute("form", "long"), new XAttribute("name-as-sort-order", "all"))), context).Plain;
        if (value.ValueKind == JsonValueKind.Object) {
            XElement date = new XElement(CslStyle.Namespace + "date", new XAttribute("variable", variable));
            return evaluator.Evaluate(date, context).Plain;
        }
        return evaluator.Evaluate(new XElement(CslStyle.Namespace + "text", new XAttribute("variable", variable)), context).Plain;
    }

    private static int? ParseSortCount(XElement key, string name) => int.TryParse((string?)key.Attribute(name), NumberStyles.None, CultureInfo.InvariantCulture, out int count) ? count : null;

    private CslCitationItem[] SortCites(IList<CslCitationItem> cites, IDictionary<string, CslRecord> records, XElement section, CslEvaluator evaluator) {
        if (section.Element(CslStyle.Namespace + "sort") == null) return cites.ToArray();
        CslRecord[] sorted = Sort(cites.Select(item => records[item.Key]).ToArray(), section, evaluator, XElementScope.Citation);
        var queues = cites.GroupBy(cite => cite.Key, StringComparer.Ordinal).ToDictionary(group => group.Key, group => new Queue<CslCitationItem>(group), StringComparer.Ordinal);
        return sorted.Select(record => queues[record.Key].Dequeue()).ToArray();
    }

    private static string AlphabeticSuffix(int index) {
        var result = new StringBuilder();
        do { result.Insert(0, (char)('a' + index % 26)); index = index / 26 - 1; } while (index >= 0);
        return result.ToString();
    }

    private static int NaturalCompare(string a, string b, CultureInfo culture, CslEvaluator evaluator) {
        string[] left = SortTokens(a, culture, evaluator), right = SortTokens(b, culture, evaluator);
        for (int index = 0; index < Math.Min(left.Length, right.Length); index++) {
            evaluator.PerformOperation();
            int order;
            if (CslNameSortKey.IsSeparator(left[index]) || CslNameSortKey.IsSeparator(right[index])) {
                order = string.CompareOrdinal(left[index], right[index]);
            } else if (left[index].Length > 0 && right[index].Length > 0 && char.IsDigit(left[index][0]) && char.IsDigit(right[index][0])) {
                string leftNumber = left[index].TrimStart('0'), rightNumber = right[index].TrimStart('0');
                order = leftNumber.Length.CompareTo(rightNumber.Length);
                if (order == 0) order = string.CompareOrdinal(leftNumber, rightNumber);
            } else order = culture.CompareInfo.Compare(left[index], right[index], SortComparison);
            if (order != 0) return order;
        }
        return left.Length.CompareTo(right.Length);
    }

    private const CompareOptions SortComparison = CompareOptions.IgnoreCase | CompareOptions.IgnoreNonSpace;

    private static string[] SortTokens(string value, CultureInfo culture, CslEvaluator evaluator) {
        var result = new List<string>();
        foreach (string part in SortParts(value, evaluator.CancellationToken)) {
            evaluator.PerformOperation();
            // Keep structural name boundaries and numeric runs; punctuation-only
            // text between them must not introduce another comparison position.
            if (CslNameSortKey.IsSeparator(part)) { result.Add(part); continue; }
            string text = SortText(part, evaluator);
            if (culture.CompareInfo.Compare(text, string.Empty, SortComparison) != 0) result.Add(text);
        }
        return result.ToArray();
    }

    private static IEnumerable<string> SortParts(string value, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        int copied = 0, position = 0;
        while (position < value.Length) {
            CslNumberSyntax.Check(position, token);
            int start = position;
            if (char.IsDigit(value[position])) {
                do { CslNumberSyntax.Check(position, token); position++; } while (position < value.Length && char.IsDigit(value[position]));
            } else if (value[position] == '\u001e' || value[position] == '\u001f') position++;
            else { position++; continue; }
            yield return value.Substring(copied, start - copied);
            yield return value.Substring(start, position - start);
            copied = position;
        }
        yield return value.Substring(copied);
    }

    private static string SortText(string value, CslEvaluator evaluator) {
        var result = new StringBuilder(value.Length);
        int nextCheck = 0;
        for (int index = 0; index < value.Length;) {
            if (index >= nextCheck) { evaluator.PerformOperation(); nextCheck = index + 256; }
            int width = char.IsHighSurrogate(value[index]) && index + 1 < value.Length && char.IsLowSurrogate(value[index + 1]) ? 2 : 1;
            // IgnoreSymbols also removes word spaces, changing alphabetic word
            // order. Remove only punctuation and symbols from each text run.
            if (!char.IsPunctuation(value, index) && !char.IsSymbol(value, index)) result.Append(value, index, width);
            index += width;
        }
        return result.ToString().Trim();
    }

    private sealed class SortRow {
        internal SortRow(CslRecord record, int index, string[] keys) { Record = record; Index = index; Keys = keys; }
        internal CslRecord Record { get; }
        internal int Index { get; }
        internal string[] Keys { get; }
    }
}
