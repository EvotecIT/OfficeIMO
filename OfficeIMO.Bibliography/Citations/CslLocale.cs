using System.Xml.Linq;

namespace OfficeIMO.Bibliography;

internal sealed class CslLocale {
    private readonly List<XElement> _sources = new List<XElement>();
    private readonly Dictionary<XElement, string> _termValues = new Dictionary<XElement, string>();
    private readonly CancellationToken _token;
    internal CultureInfo Culture { get; }
    internal string Language { get; }

    internal CslLocale(CslStyle style, CslRenderOptions options, CancellationToken token) {
        _token = token;
        string selection = options.Locale ?? style.DefaultLocale;
        string primary = selection.Split('-')[0];
        string primaryDialect = CslLocaleCatalogue.PrimaryDialect(selection);
        Language = selection.IndexOf('-') < 0 && string.Equals(primaryDialect.Split('-')[0], primary, StringComparison.OrdinalIgnoreCase) ? primaryDialect : selection;
        try { Culture = CultureInfo.GetCultureInfo(Language); }
        catch (CultureNotFoundException) { Culture = CultureInfo.InvariantCulture; }
        // Higher precedence sources come first. Each individual term/date/option
        // falls back independently rather than replacing a whole locale.
        XElement[] embedded = style.Root.Elements(CslStyle.Namespace + "locale").ToArray();
        AddMatching(embedded, Language);
        if (!string.Equals(Language, primary, StringComparison.OrdinalIgnoreCase)) AddMatching(embedded, primary);
        AddMatching(embedded, string.Empty);
        if (options.Locales.TryGetValue(Language, out string? dialect)) AddExternal(dialect, token);
        if (!string.Equals(Language, primary, StringComparison.OrdinalIgnoreCase) && options.Locales.TryGetValue(primary, out string? generic)) AddExternal(generic, token);
        if (options.Locales.TryGetValue(primaryDialect, out string? primaryXml) && !string.Equals(primaryDialect, Language, StringComparison.OrdinalIgnoreCase)) AddExternal(primaryXml, token);
        if (CslLocaleCatalogue.Get(Language) is XElement requested) _sources.Add(requested);
        if (CslLocaleCatalogue.Get(primaryDialect) is XElement standard) _sources.Add(standard);
        _sources.Add(CslLocaleCatalogue.Get("en-US")!);
    }

    private void AddMatching(IEnumerable<XElement> embedded, string language) {
        _sources.AddRange(embedded.Where(element => string.Equals((string?)element.Attribute(XNamespace.Xml + "lang") ?? string.Empty, language, StringComparison.OrdinalIgnoreCase)));
    }

    private void AddExternal(string xml, CancellationToken token) {
        XElement root = CslStyle.ReadXml(xml, 4 * 1024 * 1024, 64, token);
        if (root.Name != CslStyle.Namespace + "locale") throw new InvalidDataException("A CSL locale in the CSL namespace is required.");
        CslStyleValidation.ValidateLocale(root, token);
        _sources.Add(root);
    }

    internal string Term(string name, string form = "long", bool plural = false, string? gender = null) =>
        TryTerm(name, out string value, form, plural, gender) ? value : string.Empty;

    private bool TryTerm(string name, out string value, string form = "long", bool plural = false, string? gender = null) {
        foreach (string candidateForm in Forms(form)) {
            foreach (XElement source in _sources) {
                XElement[] terms = source.Element(CslStyle.Namespace + "terms")?.Elements(CslStyle.Namespace + "term")
                    .Where(term => ((string?)term.Attribute("name") == name || name == "editor-translator" && (string?)term.Attribute("name") == "editortranslator") &&
                        ((string?)term.Attribute("form") ?? "long") == candidateForm).ToArray() ?? Array.Empty<XElement>();
                XElement? term = terms.FirstOrDefault(value => (string?)value.Attribute("gender-form") == gender) ?? terms.FirstOrDefault(value => value.Attribute("gender-form") == null);
                if (term == null) continue;
                XElement? specific = term.Element(CslStyle.Namespace + (plural ? "multiple" : "single"));
                value = TermValue(specific ?? term);
                return true;
            }
        }
        value = string.Empty;
        return false;
    }

    private string TermValue(XElement? element) {
        if (element == null) return string.Empty;
        if (_termValues.TryGetValue(element, out string? cached)) return cached;
        string value = element.Value;
        // Normalize only logically empty XML whitespace, never individual text
        // nodes: comments and CDATA may separate meaningful padding or words.
        bool whitespaceOnly = true;
        for (int index = 0; index < value.Length; index++) {
            if ((index & 1023) == 0) _token.ThrowIfCancellationRequested();
            char character = value[index];
            if (character != ' ' && character != '\t' && character != '\r' && character != '\n') {
                whitespaceOnly = false;
                break;
            }
        }
        string? space = element.AncestorsAndSelf().Select(parent => (string?)parent.Attribute(XNamespace.Xml + "space")).FirstOrDefault(scope => scope != null);
        if (whitespaceOnly && space != "preserve") value = string.Empty;
        _termValues.Add(element, value);
        return value;
    }

    internal XElement? Date(string form) => _sources.SelectMany(source => source.Elements(CslStyle.Namespace + "date")).FirstOrDefault(date => (string?)date.Attribute("form") == form);

    internal bool Option(string name, bool fallback = false) {
        foreach (XElement source in _sources) {
            string? value = (string?)source.Element(CslStyle.Namespace + "style-options")?.Attribute(name);
            if (value != null) return value == "true";
        }
        return fallback;
    }

    internal string? Gender(string name) {
        foreach (XElement source in _sources) {
            XElement? term = source.Element(CslStyle.Namespace + "terms")?.Elements(CslStyle.Namespace + "term")
                .FirstOrDefault(value => (string?)value.Attribute("name") == name && ((string?)value.Attribute("form") ?? "long") == "long");
            if (term != null) return (string?)term.Attribute("gender");
        }
        return null;
    }

    internal string Ordinal(int value, bool longForm = false, string? gender = null) {
        long positive = Math.Abs((long)value);
        if (longForm && positive <= 10) {
            if (TryTerm("long-ordinal-" + positive.ToString("00", CultureInfo.InvariantCulture), out string term, gender: gender)) return term;
        }
        // Ordinal suffixes are one locale unit: defining any suffix replaces
        // the complete inherited set, unlike ordinary per-term fallback.
        foreach (XElement source in _sources) {
            XElement[] all = source.Element(CslStyle.Namespace + "terms")?.Elements(CslStyle.Namespace + "term")
                .Where(term => (string?)term.Attribute("name") == "ordinal" ||
                    ((string?)term.Attribute("name"))?.StartsWith("ordinal-", StringComparison.Ordinal) == true).ToArray() ?? Array.Empty<XElement>();
            if (all.Length == 0) continue;
            XElement[] terms = all.Where(term => term.Attribute("gender-form") == null || (string?)term.Attribute("gender-form") == gender)
                .OrderByDescending(term => gender != null && (string?)term.Attribute("gender-form") == gender).ToArray();
            bool legacy = !all.Any(term => (string?)term.Attribute("name") == "ordinal") &&
                Enumerable.Range(1, 4).All(index => all.Any(term => (string?)term.Attribute("name") == "ordinal-0" + index));
            if (legacy) {
                long last = positive % 100;
                int index = last >= 11 && last <= 13 || positive % 10 < 1 || positive % 10 > 3 ? 4 : (int)(positive % 10);
                XElement? old = terms.FirstOrDefault(term => (string?)term.Attribute("name") == "ordinal-0" + index);
                return value.ToString(CultureInfo.InvariantCulture) + TermValue(old);
            }
            foreach (string match in new[] { "whole-number", "last-two-digits", "last-digit" }) {
                long target = match == "whole-number" ? positive : match == "last-two-digits" ? positive % 100 : positive % 10;
                XElement? term = terms.FirstOrDefault(element => {
                    string name = (string?)element.Attribute("name") ?? string.Empty;
                    if (name.Length <= 8 || !int.TryParse(name.Substring(8), out int number)) return false;
                    string rule = (string?)element.Attribute("match") ?? (number < 10 ? "last-digit" : "last-two-digits");
                    return rule == match && number == target;
                });
                if (term != null) return value.ToString(CultureInfo.InvariantCulture) + TermValue(term);
            }
            return value.ToString(CultureInfo.InvariantCulture) + TermValue(terms.FirstOrDefault(term => (string?)term.Attribute("name") == "ordinal"));
        }
        return value.ToString(CultureInfo.InvariantCulture);
    }

    private static IEnumerable<string> Forms(string form) {
        yield return form;
        if (form == "symbol") yield return "short";
        if (form == "verb-short") yield return "verb";
        if (form != "long") yield return "long";
    }
}
