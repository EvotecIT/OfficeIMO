using System.Xml.Linq;

namespace OfficeIMO.Bibliography;

internal sealed partial class CslEvaluator {
    private readonly CslStyle _style;
    private readonly CslLocale _locale;
    private readonly CslRenderOptions _options;
    private readonly CancellationToken _token;
    private int _operations;
    internal CslEvaluator(CslStyle style, CslLocale locale, CslRenderOptions options, CancellationToken token) {
        _style = style; _locale = locale; _options = options; _token = token;
    }
    internal void ThrowIfCancellationRequested() => _token.ThrowIfCancellationRequested();
    internal CancellationToken CancellationToken => _token;
    internal void PerformOperation() {
        _token.ThrowIfCancellationRequested();
        if (_operations++ >= _options.MaximumRenderingOperations) throw new InvalidDataException("CSL rendering exceeds MaximumRenderingOperations.");
    }
    private CslText Join(IEnumerable<CslText> values, string delimiter, bool suppressEmptyVariables = false) =>
        CslText.Join(values, delimiter, suppressEmptyVariables, _options.MaximumIntermediateCharacters);

    internal CslText Evaluate(XElement element, CslContext context, int depth = 0) {
        PerformOperation();
        if (depth >= _style.MaximumDepth) throw new InvalidDataException("CSL rendering exceeds MaximumNestingDepth.");
        CslText value;
        bool captureNames = element.Name.LocalName == "names" && context.NamesDepth == 0 && context.Cite?.AuthorOnly == true && context.NarrativeNames == null;
        switch (element.Name.LocalName) {
            case "layout":
                value = Layout(element, context, depth);
                break;
            case "macro": case "group":
                value = Children(element, context, depth, suppressEmptyVariables: true);
                if (!value.IsEmpty) value = new CslText(value.Plain, value.Html, 1, 1);
                break;
            case "text": value = Text(element, context, depth); break;
            case "number": value = Number(element, context); break;
            case "label": value = Label(element, context); break;
            case "names":
                context.NamesDepth++;
                try { value = Names(element, context, depth); }
                finally { context.NamesDepth--; }
                break;
            case "date": value = Date(element, context); break;
            case "choose":
                XElement? branch = element.Elements().FirstOrDefault(child => child.Name.LocalName == "else" || Matches(child, context));
                value = branch == null ? CslText.Empty : Join(branch.Elements().Select(child => Evaluate(child, context, depth + 1)), context.InheritedDelimiter);
                break;
            default: throw new InvalidDataException("Unsupported CSL rendering element '" + element.Name.LocalName + "'.");
        }
        XElement formatting = element;
        string language = context.Record.Scalar("language");
        if (Attr(element, "text-case") == "title") {
            if (language.Length == 0) language = _style.DefaultLocale;
            if (!language.StartsWith("en", StringComparison.OrdinalIgnoreCase)) { formatting = new XElement(element); formatting.Attribute("text-case")?.Remove(); }
        }
        CultureInfo? textCulture = null;
        if (language.Length > 0) {
            try { textCulture = CultureInfo.GetCultureInfo(language); }
            catch (CultureNotFoundException) { }
        }
        value = LinkBibliographyText(value, element, context, ref formatting);
        CslText decorated = value.Decorate(formatting, _locale, context.Sorting, textCulture: textCulture, cancellationToken: _token);
        if (decorated.Html.Length > _options.MaximumIntermediateCharacters || decorated.Plain.Length > _options.MaximumIntermediateCharacters)
            throw new InvalidDataException("CSL rendering exceeds MaximumIntermediateCharacters.");
        if (!decorated.IsEmpty && (element.Name.LocalName == "text" || element.Name.LocalName == "number" || element.Name.LocalName == "date") &&
            Attr(element, "variable") is string renderedVariable)
            context.RegisterRenderedVariable(renderedVariable);
        if (captureNames && !decorated.IsEmpty) context.NarrativeNames = decorated;
        return decorated;
    }

    private CslText Layout(XElement element, CslContext context, int depth) {
        CslText[] fields = element.Elements().Select(child => Evaluate(child, context, depth + 1)).ToArray();
        if (context.Scope != XElementScope.Bibliography || context.Sorting ||
            _style.Root.Element(CslStyle.Namespace + "bibliography")?.Attribute("second-field-align") == null)
            return Join(fields, string.Empty);

        int first = Array.FindIndex(fields, field => !field.IsEmpty);
        if (first < 0) return CslText.Empty;
        CslText remaining = Join(fields.Skip(first + 1), string.Empty);
        if (remaining.IsEmpty) return fields[first];
        return Join(new[] { fields[first].MarkFirstAlignedField(), remaining }, string.Empty);
    }

    private CslText Children(XElement element, CslContext context, int depth, bool suppressEmptyVariables = false) {
        string prior = context.InheritedDelimiter;
        string delimiter = Attr(element, "delimiter") ?? string.Empty;
        context.InheritedDelimiter = delimiter;
        try { return Join(element.Elements().Select(child => Evaluate(child, context, depth + 1)), delimiter, suppressEmptyVariables); }
        finally { context.InheritedDelimiter = prior; }
    }

    private CslText Text(XElement element, CslContext context, int depth) {
        string? macro = Attr(element, "macro");
        if (macro != null) {
            string prior = context.MacroPath;
            context.MacroPath += (element.Annotation<CslElementIdentity>()?.Key ?? string.Empty) + "/";
            try { return Evaluate(_style.Macros[macro], context, depth + 1); }
            finally { context.MacroPath = prior; }
        }
        string? variable = Attr(element, "variable");
        if (variable != null) {
            if (context.Suppressed.Contains(variable)) return CslText.Literal(string.Empty, true);
            string value = context.Variable(variable);
            if (Attr(element, "form") == "short") {
                string shortValue = context.Variable(variable + "-short");
                if (shortValue.Length > 0) value = shortValue;
                else if (_options.Abbreviations.TryGetValue(variable, out IDictionary<string, string>? abbreviations) && abbreviations.TryGetValue(value, out string? abbreviation)) value = abbreviation;
            }
            if (variable == "page") value = PageRange(value, Attr(_style.Root, "page-range-format"));
            else if (variable == "locator") value = value.Replace("-", "–");
            else if ((variable == "volume" || variable == "issue" || variable == "edition" || variable == "number") && IsNumeric(value))
                value = CslNumberSyntax.NormalizeNumericHyphens(value, _token);
            if (variable == "locator") value = LocalizeNumberConnectors(value);
            return CslText.Rich(value, cancellationToken: _token);
        }
        string? term = Attr(element, "term");
        if (term != null) return CslText.Literal(_locale.Term(term, Attr(element, "form") ?? "long", Attr(element, "plural") == "true"));
        return CslText.Rich(Attr(element, "value") ?? string.Empty, false, _token);
    }

    private CslText Number(XElement element, CslContext context) {
        string variable = Attr(element, "variable") ?? string.Empty;
        string value = context.Variable(variable);
        if (!IsNumeric(value)) return CslText.Literal(variable == "locator" ? value.Replace("-", "–") : value, true);
        value = CslNumberSyntax.NormalizeSeparators(value, _options.MaximumIntermediateCharacters, _token);
        string form = Attr(element, "form") ?? "numeric";
        string ConvertNumber(string digits) {
            if (!int.TryParse(digits, NumberStyles.None, CultureInfo.InvariantCulture, out int number)) return digits;
            switch (form) {
                case "ordinal": return _locale.Ordinal(number, gender: Gender(NumberTerm(variable, context)));
                case "long-ordinal": return _locale.Ordinal(number, true, Gender(NumberTerm(variable, context)));
                case "roman": return Roman(number);
                default: return digits;
            }
        }
        // Convert resolved source endpoints before inserting literal locale delimiters or generated suffixes.
        if (variable == "page" && form != "numeric" && Attr(_style.Root, "page-range-format") != null) {
            string pages = CslPageRangeFormatter.FormatNumbers(value, Attr(_style.Root, "page-range-format")!, _locale.Term("page-range-delimiter"), ConvertNumber, LocalizeNumberConnectors, _token, _options.MaximumIntermediateCharacters);
            return CslText.Literal(pages, true);
        }
        if (variable == "page" && form == "numeric") return CslText.Literal(PageRange(value, Attr(_style.Root, "page-range-format"), true), true);
        string rendered = CslNumberFormatter.Format(value, ConvertNumber, _locale.Term("and", "symbol"), variable == "locator",
            variable != "page" && form == "numeric", _options.MaximumIntermediateCharacters, _token);
        return CslText.Literal(rendered, true);
    }

    private CslText Label(XElement element, CslContext context) {
        string variable = Attr(element, "variable") ?? string.Empty;
        string value = context.Variable(variable);
        if (value.Length == 0) return CslText.Empty;
        string term = NumberTerm(variable, context);
        bool plural = Attr(element, "plural") == "always" || Attr(element, "plural") != "never" &&
            (variable.StartsWith("number-of-", StringComparison.Ordinal) ? long.TryParse(value, out long numeric) && numeric > 1 :
                CslNumberFormatter.IsPlural(value, _token));
        return CslText.Literal(_locale.Term(term, Attr(element, "form") ?? "long", plural));
    }

    private static string NumberTerm(string variable, CslContext context) => variable == "locator" ? context.Cite?.LocatorType ?? "page" :
        variable == "number-of-pages" ? "page" : variable == "number-of-volumes" ? "volume" : variable;

    private bool Matches(XElement element, CslContext context) {
        var conditions = new List<bool>();
        foreach (XAttribute attribute in element.Attributes()) {
            string[] values = attribute.Value.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            switch (attribute.Name.LocalName) {
                case "type": conditions.AddRange(values.Select(value => context.Record.Type == value)); break;
                case "variable": conditions.AddRange(values.Select(value => HasVariable(context, value))); break;
                case "is-numeric": conditions.AddRange(values.Select(value => IsNumeric(context.Variable(value, includeSuppressed: true)))); break;
                case "position": conditions.AddRange(values.Select(value => !context.Sorting && context.Scope == XElementScope.Citation &&
                    (value == context.Position || value == "subsequent" && context.Position != "first" || value == "ibid" && context.Position == "ibid-with-locator" || value == "near-note" && context.NearNote))); break;
                case "locator": conditions.AddRange(values.Select(value => context.Cite != null && !string.IsNullOrEmpty(context.Cite.Locator) && context.Cite.LocatorType == (value == "sub-verbo" ? "sub verbo" : value))); break;
                case "disambiguate": conditions.Add(ConditionalDetail(element, context) == (attribute.Value == "true")); break;
                case "is-uncertain-date": conditions.AddRange(values.Select(value => context.Record.Value(value).ValueKind == System.Text.Json.JsonValueKind.Object &&
                    context.Record.Value(value).TryGetProperty("circa", out var circa) && (circa.ValueKind == System.Text.Json.JsonValueKind.True || circa.ValueKind == System.Text.Json.JsonValueKind.Number && circa.GetRawText() == "1"))); break;
            }
        }
        return (Attr(element, "match") ?? "all") switch {
            "any" => conditions.Any(value => value),
            "none" => conditions.All(value => !value),
            _ => conditions.All(value => value)
        };
    }

    private bool ConditionalDetail(XElement element, CslContext context) {
        // Bibliography branches reflect whether the work needed citation detail.
        // Citation branches distinguish each reusable form and macro call site.
        if (context.Scope == XElementScope.Bibliography) return context.Record.ActiveConditions.Count > 0;
        string form = context.Position == "first" || !_style.HasSubsequentForm ? "0/" :
            _style.IsNoteStyle && _style.HasNearNoteCondition && context.NearNote ? "2/" : "1/";
        string locator = context.LocatorDisambiguationKey ??= _style.HasLocatorConditions && !string.IsNullOrEmpty(context.Cite?.Locator) ?
            "L" + context.Cite!.LocatorType.Length.ToString(CultureInfo.InvariantCulture) + ":" + context.Cite.LocatorType + ":" +
            (CslNumberSyntax.IsNumeric(context.Cite.Locator!, _token) ? "N/" : "T/") : string.Empty;
        string key = form + locator + context.MacroPath + element.Annotation<CslElementIdentity>()!.Key;
        context.ObservedConditions?.Add(key);
        return context.Record.ActiveConditions.Contains(key);
    }

    // Substitution suppresses repeated output; conditions still inspect source values.
    private static bool HasVariable(CslContext context, string variable) =>
        variable == "locator" ? !string.IsNullOrEmpty(context.Cite?.Locator) : variable == "year-suffix" ? context.Record.YearSuffix.Length > 0 : variable == "first-reference-note-number" ? context.Variable(variable, includeSuppressed: true).Length > 0 : variable == "citation-number" || context.Record.Has(variable);

    private bool IsNumeric(string value) => CslNumberSyntax.IsNumeric(value, _token);
    private string LocalizeNumberConnectors(string value) => value.IndexOf('&') < 0 ? value :
        CslNumberSyntax.LocalizeAmpersands(value, _locale.Term("and", "symbol"), _options.MaximumIntermediateCharacters, _token);
    private string? Gender(string variable) => _locale.Gender(variable);
    internal static string? Attr(XElement element, string name) => (string?)element.Attribute(name);

    internal string? NameOption(XElement element, CslContext context, string name) {
        string inherited = name == "form" ? "name-form" : name == "delimiter" ? "name-delimiter" : name;
        XElement? section = _style.Root.Element(CslStyle.Namespace + (context.Scope == XElementScope.Citation ? "citation" : "bibliography"));
        return Attr(element, name) ?? (section == null ? null : Attr(section, inherited)) ?? Attr(_style.Root, inherited);
    }

    private static string Roman(int number) {
        if (number < 1 || number > 3999) return number.ToString(CultureInfo.InvariantCulture);
        var output = new StringBuilder();
        foreach (var pair in new[] { (1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"), (50, "l"), (40, "xl"), (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i") })
            while (number >= pair.Item1) { output.Append(pair.Item2); number -= pair.Item1; }
        return output.ToString();
    }

    private string PageRange(string value, string? format, bool literal = false) {
        if (format == null) return LocalizeNumberConnectors(value);
        return CslPageRangeFormatter.Format(value, format, _locale.Term("page-range-delimiter"), _token,
            value.IndexOf('&') < 0 ? null : LocalizeNumberConnectors, literal ? _options.MaximumIntermediateCharacters : int.MaxValue);
    }
}
