using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace OfficeIMO.Bibliography;

internal sealed partial class CslEvaluator {
    private CslText Names(XElement element, CslContext context, int depth) {
        string[] variables = (Attr(element, "variable") ?? string.Empty).Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        int selectedVariableCount = variables.Length;
        variables = variables.Where(variable => !context.Suppressed.Contains(variable)).ToArray();
        if (selectedVariableCount == 2 && variables.Contains("editor") && variables.Contains("translator") && CslRecord.SameNames(context.Record.Names("editor"), context.Record.Names("translator"), _token))
            variables = variables.Select(variable => variable == "editor" ? "editor-translator" : variable).Where(variable => variable != "translator").ToArray();
        bool firstNames = !context.FirstNamesHandled;
        XElement name = element.Element(CslStyle.Namespace + "name") ?? new XElement(CslStyle.Namespace + "name");
        bool countForm = NameOption(name, context, "form") == "count";
        var lists = new List<CslText>();
        int totalCount = 0;
        bool hasNames = false;
        foreach (string variable in variables) {
            if (context.Suppressed.Contains(variable) || variable == "author" && context.Cite?.SuppressAuthor == true) continue;
            JsonElement[] names = context.Record.Names(variable, _token);
            if (names.Length == 0) continue;
            hasNames = true;
            if (countForm) {
                NameSelection selection = SelectNames(name, names.Length, context);
                totalCount += selection.Count + (selection.IncludeLast ? 1 : 0);
                RegisterNameVariable(context, variable);
                continue;
            }
            int previousReplacements = context.EmptyNamesReplacementCount;
            CslText rendered = RenderNames(name, element, names, context, variable, out bool completeListReplaced);
            if (!completeListReplaced) rendered = rendered.Affix(Attr(name, "prefix") ?? string.Empty, Attr(name, "suffix") ?? string.Empty);
            XElement? label = element.Element(CslStyle.Namespace + "label");
            if (label != null && !context.Sorting) {
                bool plural = Attr(label, "plural") == "always" || Attr(label, "plural") != "never" && names.Length > 1;
                CslText role = CslText.Literal(_locale.Term(variable, Attr(label, "form") ?? "long", plural)).Decorate(label, _locale, context.Sorting, cancellationToken: _token);
                // Without an explicit cs:name, the default name occupies the first position.
                bool labelFirst = label.ElementsAfterSelf().Any(sibling => sibling.Name == CslStyle.Namespace + "name");
                rendered = Join(labelFirst ? new[] { role, rendered } : new[] { rendered, role }, string.Empty);
            }
            lists.Add(rendered);
            if (!rendered.IsEmpty || context.EmptyNamesReplacementCount != previousReplacements) {
                RegisterNameVariable(context, variable);
            }
        }
        if (countForm && hasNames) return CslText.Literal(totalCount.ToString(CultureInfo.InvariantCulture), true);
        if (lists.Count > 0) {
            CslText result = Join(lists, Attr(element, "delimiter") ?? NameOption(element, context, "names-delimiter") ?? string.Empty);
            return FinishNames(result, context, firstNames);
        }
        XElement? substitute = element.Element(CslStyle.Namespace + "substitute");
        if (substitute != null) {
            foreach (XElement replacement in substitute.Elements()) {
                XElement effective = replacement;
                if (replacement.Name == CslStyle.Namespace + "names") {
                    effective = CslElementIdentity.Copy(replacement);
                    foreach (XElement inherited in element.Elements().Where(child => child.Name.LocalName == "name" || child.Name.LocalName == "et-al" || child.Name.LocalName == "label"))
                        if (effective.Element(inherited.Name) == null) effective.Add(CslElementIdentity.Copy(inherited));
                }
                CslSubstitutionScope scope = context.BeginSubstitution();
                bool successful = false;
                try {
                    CslText result = Evaluate(effective, context, depth + 1);
                    if (result.IsEmpty && context.EmptyNamesReplacementCount == scope.EmptyNamesReplacementCount) continue;
                    successful = true;
                    return countForm ? result : FinishNames(result, context, firstNames);
                } finally { context.EndSubstitution(scope, successful); }
            }
        }
        return CslText.Literal(string.Empty, true);
    }

    private CslText RenderNames(XElement name, XElement namesElement, JsonElement[] source, CslContext context, string variable, out bool completeListReplaced) {
        completeListReplaced = false;
        string delimiter = NameOption(name, context, "delimiter") ?? NameOption(namesElement, context, "name-delimiter") ?? ", ";
        string form = NameOption(name, context, "form") ?? "long";
        string? conjunction = NameOption(name, context, "and");
        string and = conjunction == "symbol" ? _locale.Term("and", "symbol") : conjunction == "text" ? _locale.Term("and") : string.Empty;
        NameSelection selection = SelectNames(name, source.Length, context);
        bool etAl = selection.Abbreviated;
        int count = selection.Count;
        var values = new List<CslText>();
        var formatting = new XElement(name);
        formatting.Attribute("prefix")?.Remove(); formatting.Attribute("suffix")?.Remove();
        for (int index = 0; index < count; index++) {
            _token.ThrowIfCancellationRequested();
            CslText value = RenderName(name, source[index], index, context, form, variable).Decorate(formatting, _locale, context.Sorting, cancellationToken: _token);
            values.Add(value);
            if (!context.Sorting) context.ObservedNames.Add(new CslNameOccurrence(context.Record, variable, index, source[index], name, form, value.Plain, context.ObservedNames.Count == 0));
        }
        bool includeLast = selection.IncludeLast;
        CslText lastName = includeLast ? RenderName(name, source[source.Length - 1], source.Length - 1, context, form, variable).Decorate(formatting, _locale, context.Sorting, cancellationToken: _token) : CslText.Empty;
        if (context.Sorting) return Join(includeLast ? values.Concat(new[] { lastName }) : values, CslNameSortKey.ListDelimiter);
        if (!context.FirstNamesHandled) {
            context.FirstNamesHandled = true;
            context.FirstNames = (includeLast ? values.Concat(new[] { lastName }) : values).Select(value => value.Plain).ToArray();
            context.CompleteNamesMatch = context.PreviousNames.Length > 0 && context.FirstNames.SequenceEqual(context.PreviousNames, StringComparer.Ordinal);
            if (context.NamesReplacement != null && context.NamesReplacementRule == "complete-all" && context.CompleteNamesMatch) {
                if (context.NamesReplacement.Length == 0) context.EmptyNamesReplacementCount++;
                // Replace the entire name list and its affixes before adding its contributor label.
                completeListReplaced = true;
                return CslText.Literal(context.NamesReplacement);
            }
            if (context.NamesReplacement != null && context.NamesReplacementRule != "complete-all") {
                bool completeEach = context.NamesReplacementRule == "complete-each";
                for (int index = 0; index < values.Count && index < context.PreviousNames.Length; index++) {
                    if (completeEach && !context.CompleteNamesMatch || !string.Equals(values[index].Plain, context.PreviousNames[index], StringComparison.Ordinal)) break;
                    if (context.NamesReplacement.Length == 0) context.EmptyNamesReplacementCount++;
                    values[index] = CslText.Literal(context.NamesReplacement).Decorate(formatting, _locale, context.Sorting, cancellationToken: _token);
                    if (context.NamesReplacementRule == "partial-first") break;
                }
                if (includeLast && completeEach && context.CompleteNamesMatch) {
                    if (context.NamesReplacement.Length == 0) context.EmptyNamesReplacementCount++;
                    lastName = CslText.Literal(context.NamesReplacement).Decorate(formatting, _locale, context.Sorting, cancellationToken: _token);
                }
            }
        }
        if (etAl) {
            if (includeLast) {
                return Join(new[] { Join(values, delimiter), lastName }, context.Sorting ? delimiter : delimiter + "… ");
            }
            if (context.Sorting) return Join(values, delimiter);
            XElement? etAlElement = namesElement.Element(CslStyle.Namespace + "et-al");
            string term = Attr(etAlElement ?? new XElement("et-al"), "term") ?? "et-al";
            CslText label = CslText.Literal(_locale.Term(term));
            if (etAlElement != null) label = label.Decorate(etAlElement, _locale, context.Sorting, cancellationToken: _token);
            string setting = NameOption(name, context, "delimiter-precedes-et-al") ?? "contextual";
            string separator = setting == "always" || setting == "contextual" && values.Count > 1 || setting == "after-inverted-name" && count > 0 && IsPersonalInverted(name, source[count - 1], count - 1, context, form) ? delimiter : " ";
            return Join(new[] { Join(values, delimiter), label }, separator);
        }
        if (values.Count < 2 || and.Length == 0) return Join(values, delimiter);
        string before = NameOption(name, context, "delimiter-precedes-last") ?? "contextual";
        string lastSeparator = before == "always" || before == "contextual" && values.Count > 2 || before == "after-inverted-name" && IsPersonalInverted(name, source[values.Count - 2], values.Count - 2, context, form) ? delimiter : " ";
        CslText preceding = Join(values.Take(values.Count - 1), delimiter);
        CslText finalName = values[values.Count - 1];
        return new CslText(preceding.Plain + lastSeparator + and + " " + finalName.Plain,
            preceding.Html + CslText.Escape(lastSeparator + and + " ") + finalName.Html);
    }

    private NameSelection SelectNames(XElement name, int length, CslContext context) {
        string? subsequentMin = context.Position != "first" ? NameOption(name, context, "et-al-subsequent-min") : null;
        string? subsequentFirst = context.Position != "first" ? NameOption(name, context, "et-al-subsequent-use-first") : null;
        string? minimumOption = subsequentMin ?? NameOption(name, context, "et-al-min");
        string? firstOption = subsequentFirst ?? NameOption(name, context, "et-al-use-first");
        int min = ParsePositive(minimumOption, int.MaxValue);
        int first = ParsePositive(firstOption, 1);
        bool hasMinimum = minimumOption != null || context.Sorting && context.SortNamesMinimum.HasValue;
        bool hasFirst = firstOption != null || context.Sorting && context.SortNamesFirst.HasValue;
        if (context.Sorting) { min = context.SortNamesMinimum ?? min; first = context.SortNamesFirst ?? first; }
        else first = Math.Max(first, context.Record.MinimumNames);
        bool abbreviated = hasMinimum && hasFirst && length >= min && first < length;
        bool useLast = (context.Sorting ? context.SortNamesLast : null) ?? (NameOption(name, context, "et-al-use-last") == "true");
        return new NameSelection(abbreviated ? first : length, abbreviated, abbreviated && useLast && length - first >= 2);
    }

    private static void RegisterNameVariable(CslContext context, string variable) {
        if (variable == "editor-translator") {
            context.RegisterRenderedVariable("editor");
            context.RegisterRenderedVariable("translator");
        } else context.RegisterRenderedVariable(variable);
    }

    private readonly struct NameSelection {
        internal NameSelection(int count, bool abbreviated, bool includeLast) { Count = count; Abbreviated = abbreviated; IncludeLast = includeLast; }
        internal int Count { get; }
        internal bool Abbreviated { get; }
        internal bool IncludeLast { get; }
    }

    private CslText RenderName(XElement element, JsonElement value, int index, CslContext context, string form, string variable) {
        XElement? familyElement = element.Elements(CslStyle.Namespace + "name-part").FirstOrDefault(part => Attr(part, "name") == "family");
        XElement? givenElement = element.Elements(CslStyle.Namespace + "name-part").FirstOrDefault(part => Attr(part, "name") == "given");
        string literal = Scalar(value, "literal");
        if (literal.Length > 0) {
            CslText institution = FormatNamePart(CslText.Rich(literal, false, _token), familyElement, context);
            return context.Sorting ? CslNameSortKey.Personal(CslNameSortKey.StripInstitutionArticle(institution.Plain), string.Empty, string.Empty, string.Empty) : AffixNamePart(institution, familyElement, context);
        }
        string family = Scalar(value, "family"), given = Scalar(value, "given");
        string nonDropping = Scalar(value, "non-dropping-particle"), dropping = Scalar(value, "dropping-particle"), suffix = Scalar(value, "suffix");
        string? initialize = NameOption(element, context, "initialize-with");
        int expansion = !context.Sorting && context.Record.NameExpansions.TryGetValue(variable + "\u001f" + index.ToString(CultureInfo.InvariantCulture), out int expanded) ? expanded : 0;
        if (expansion > 0) form = "long";
        bool inverted = IsPersonalInverted(element, value, index, context, form);
        CslText givenValue = CslText.Rich(given, false, _token);
        if (family.Length > 0 && initialize != null && expansion < 2)
            givenValue = givenValue.InitializeName(initialize, Attr(_style.Root, "initialize-with-hyphen") != "false",
                NameOption(element, context, "initialize") != "false", _options.MaximumIntermediateCharacters, _token);
        bool demote = (Attr(_style.Root, "demote-non-dropping-particle") ?? "display-and-sort") == "display-and-sort" || context.Sorting && (Attr(_style.Root, "demote-non-dropping-particle") ?? "display-and-sort") == "sort-only";
        CslText familyText = FormatNamePart(CslText.Rich(JoinWords(inverted && demote ? string.Empty : nonDropping, family), false, _token), familyElement, context);
        CslText givenText = FormatNamePart(givenValue, givenElement, context);
        CslText droppingText = FormatNamePart(CslText.Rich(dropping, false, _token), givenElement, context);
        if (context.Sorting) {
            if (family.Length == 0 && !givenText.IsEmpty) return CslNameSortKey.Personal(givenText.Plain, string.Empty, string.Empty, string.Empty);
            bool asian = IsAsianName(family);
            CslText sortFamily = FormatNamePart(CslText.Rich(JoinWords(demote && !asian ? string.Empty : nonDropping, family), false, _token), familyElement, context);
            CslText particles = asian ? CslText.Empty : JoinNameParts(form == "short" ? CslText.Empty : droppingText,
                demote ? FormatNamePart(CslText.Rich(nonDropping, false, _token), familyElement, context) : CslText.Empty);
            return CslNameSortKey.Personal(sortFamily.Plain, particles.Plain, form == "short" ? string.Empty : givenText.Plain, form == "short" ? string.Empty : suffix);
        }
        if (form == "short") return AffixNamePart(familyText.IsEmpty ? givenText : familyText, familyText.IsEmpty ? givenElement : familyElement, context);
        if (IsAsianName(family)) return Join(new[] { familyText, givenText }, string.Empty);
        string separator = inverted ? NameOption(element, context, "sort-separator") ?? ", " : " ";
        if (inverted) {
            CslText demoted = demote ? FormatNamePart(CslText.Rich(nonDropping, false, _token), familyElement, context) : CslText.Empty;
            givenText = JoinNameParts(givenText, droppingText, demoted);
        } else {
            familyText = JoinNameParts(droppingText, familyText);
            if (suffix.Length > 0) familyText = familyText.Affix(string.Empty, (Bool(value, "comma-suffix") ? ", " : " ") + suffix);
        }
        familyText = AffixNamePart(familyText, familyElement, context);
        givenText = AffixNamePart(givenText, givenElement, context);
        CslText result = inverted ? Join(new[] { familyText, givenText }, separator) : JoinNameParts(givenText, familyText);
        if (inverted && suffix.Length > 0) result = result.Affix(string.Empty, (Bool(value, "comma-suffix") ? ", " : separator) + suffix);
        return result;
    }

    private CslText FormatNamePart(CslText value, XElement? element, CslContext context) =>
        element == null ? value : value.Decorate(element, _locale, context.Sorting, includeAffixes: false, cancellationToken: _token);

    private static CslText AffixNamePart(CslText value, XElement? element, CslContext context) => element == null || context.Sorting ? value :
        value.Affix(Attr(element, "prefix") ?? string.Empty, Attr(element, "suffix") ?? string.Empty);

    private CslText JoinNameParts(params CslText[] parts) {
        CslText result = CslText.Empty;
        foreach (CslText part in parts) {
            string separator = result.IsEmpty || char.IsWhiteSpace(result.Plain[result.Plain.Length - 1]) ||
                "-'’".IndexOf(result.Plain[result.Plain.Length - 1]) >= 0 ? string.Empty : " ";
            result = Join(new[] { result, part }, separator);
        }
        return result;
    }

    internal bool CanExpandWithInitials(CslNameOccurrence name) {
        var context = new CslContext(name.Record, XElementScope.Citation);
        return NameOption(name.Element, context, "initialize-with") != null && NameOption(name.Element, context, "initialize") != "false";
    }

    internal string NameIdentity(CslNameOccurrence name) => CslNameIdentity.Read(name.Data, false, _token);

    internal string PreviewExpandedName(CslNameOccurrence name, int expansion) {
        // A later expression can need less detail than an earlier one. Both
        // forms share this name key, so preserve the stronger expansion.
        name.Record.NameExpansions.TryGetValue(name.Key, out int existing);
        name.Record.NameExpansions[name.Key] = Math.Max(existing, expansion);
        var context = new CslContext(name.Record, XElementScope.Citation);
        return RenderName(name.Element, name.Data, name.Index, context, name.Form, name.Variable).Decorate(name.Element, _locale, cancellationToken: _token).Plain;
    }

    private bool IsInverted(XElement element, int index, CslContext context) => context.Sorting || NameOption(element, context, "name-as-sort-order") == "all" || index == 0 && NameOption(element, context, "name-as-sort-order") == "first";
    private bool IsPersonalInverted(XElement element, JsonElement value, int index, CslContext context, string form) =>
        form != "short" && Scalar(value, "literal").Length == 0 && Scalar(value, "given").Length > 0 &&
        Scalar(value, "family").Length > 0 && !IsAsianName(Scalar(value, "family")) && IsInverted(element, index, context);
    private static string JoinWords(params string[] values) {
        var result = new StringBuilder();
        foreach (string value in values.Where(value => value.Length > 0)) {
            if (result.Length > 0 && !char.IsWhiteSpace(result[result.Length - 1]) && result[result.Length - 1] != '-' && result[result.Length - 1] != '\'' && result[result.Length - 1] != '’') result.Append(' ');
            result.Append(value);
        }
        return result.ToString();
    }
    private static CslText FinishNames(CslText result, CslContext context, bool firstNames) {
        if (!firstNames || result.IsEmpty) return new CslText(result.Plain, result.Html, 1, result.IsEmpty ? 0 : 1);
        bool textFallback = !context.FirstNamesHandled;
        if (textFallback) { context.FirstNamesHandled = true; context.FirstNames = new[] { result.Plain }; context.CompleteNamesMatch = context.PreviousNames.SequenceEqual(context.FirstNames, StringComparer.Ordinal); }
        context.FirstNamesText = result.Plain;
        if (context.SuppressFirstNames) return CslText.Literal(string.Empty, true);
        if (context.NamesReplacement != null && textFallback && context.CompleteNamesMatch) {
            if (context.NamesReplacement.Length == 0) context.EmptyNamesReplacementCount++;
            return CslText.Literal(context.NamesReplacement, true);
        }
        return new CslText(result.Plain, result.Html, 1, 1);
    }
    private static string Scalar(JsonElement value, string name) => value.TryGetProperty(name, out JsonElement property) && property.ValueKind == JsonValueKind.String ? property.GetString() ?? string.Empty : string.Empty;
    private static bool Bool(JsonElement value, string name) => value.TryGetProperty(name, out JsonElement property) && property.ValueKind == JsonValueKind.True;
    private static bool IsAsianName(string value) => value.Any(character => character >= '\u3040' && character <= '\u30ff' || character >= '\u3400' && character <= '\u9fff' || character >= '\uac00' && character <= '\ud7af');
    private static int ParsePositive(string? value, int fallback) => int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int parsed) ? parsed : fallback;
}
