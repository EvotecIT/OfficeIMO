using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

/// <summary>A form field that holds searched text, and what a redaction has to change to clear it.</summary>
internal sealed class PdfRedactionAffectedField {
    internal PdfRedactionAffectedField(string name) => Name = name;

    internal string Name { get; }
    internal bool NameHoldsSearch { get; set; }
    internal List<PdfDictionary> ValueOwners { get; } = new();
    internal List<PdfDictionary> Widgets { get; } = new();
    internal HashSet<int> Pages { get; } = new();
    internal PdfDictionary? FirstWidget { get; set; }
}

internal sealed class PdfRedactionFormAnalysis {
    internal List<PdfRedactionAffectedField> Fields { get; } = new();
    internal bool XfaHoldsSearch { get; set; }
    internal bool XfaCanBeRemoved { get; set; }
}

internal static partial class PdfRedactionApplier {
    private static readonly string[] FieldValueKeys = { "V", "DV", "RV" };
    private static readonly string[] FieldDataKeys = { "TU", "TM", "Contents" };
    private static readonly string[] CaptionKeys = { "CA", "RC", "AC" };
    private static readonly Regex XmlTag = new Regex("<[^>]*>", RegexOptions.CultureInvariant, RegexTimeout);
    private const int MaximumFormFieldDepth = 128;

    /// <summary>
    /// Finds the fields holding searched text on the pages in scope. A value is shared by every widget of its
    /// field, so a field reaching a page in scope is affected as a whole. Signature fields are not values.
    /// </summary>
    internal static PdfRedactionFormAnalysis AnalyzeFormFields(Dictionary<int, PdfIndirectObject> objects, PdfDictionary? catalog,
        IReadOnlyList<int> pageObjectNumbers, PdfRedactionSearchCriteria[] criteria, ISet<int> scope, bool includeFieldsWithoutWidgets) {
        var analysis = new PdfRedactionFormAnalysis();
        if (criteria.Length == 0 || catalog is null) return analysis;
        PdfDictionary? acroForm = PdfObjectLookup.Resolve(objects, catalog.Items.TryGetValue("AcroForm", out PdfObject? formValue) ? formValue : null) as PdfDictionary;
        var limits = new PdfReadLimits();
        var pageOfWidget = new Dictionary<PdfDictionary, int>();
        for (int index = 0; index < pageObjectNumbers.Count; index++) {
            if (!objects.TryGetValue(pageObjectNumbers[index], out PdfIndirectObject? pageObject) || pageObject.Value is not PdfDictionary page ||
                PdfObjectLookup.Resolve(objects, page.Items.TryGetValue("Annots", out PdfObject? annots) ? annots : null) is not PdfArray annotations) continue;
            foreach (PdfObject item in annotations.Items) {
                if (PdfObjectLookup.Resolve(objects, item) is PdfDictionary annotation) pageOfWidget[annotation] = index + 1;
            }
        }
        PdfDictionary? anyPage = pageObjectNumbers.Count > 0 && objects.TryGetValue(pageObjectNumbers[0], out PdfIndirectObject? first) ? first.Value as PdfDictionary : null;
        var visited = new HashSet<PdfDictionary>();
        if (acroForm is not null && PdfObjectLookup.Resolve(objects, acroForm.Items.TryGetValue("Fields", out PdfObject? fieldsValue) ? fieldsValue : null) is PdfArray fields) {
            foreach (PdfObject field in fields.Items) Visit(field, string.Empty, null, 0);
        }
        // Viewers show a widget whether or not /Fields lists it, so one left out of the tree still counts.
        foreach (PdfDictionary annotation in pageOfWidget.Keys.ToArray()) {
            if (visited.Contains(annotation) || !string.Equals(TryReadName(objects, annotation, "Subtype"), "Widget", StringComparison.Ordinal)) continue;
            PdfDictionary root = annotation;
            for (int depth = 0; depth < MaximumFormFieldDepth &&
                PdfObjectLookup.Resolve(objects, root.Items.TryGetValue("Parent", out PdfObject? parent) ? parent : null) is PdfDictionary next; depth++) root = next;
            Visit(root, string.Empty, null, 0);
        }

        bool anyPageInScope = analysis.Fields.Count > 0 || includeFieldsWithoutWidgets;
        if (acroForm is not null && anyPageInScope && acroForm.Items.TryGetValue("XFA", out PdfObject? xfaValue) && XfaHoldsSearch(PdfObjectLookup.Resolve(objects, xfaValue))) {
            analysis.XfaHoldsSearch = true;
            bool hasFields = PdfObjectLookup.Resolve(objects, acroForm.Items.TryGetValue("Fields", out PdfObject? hybrid) ? hybrid : null) is PdfArray { Items.Count: > 0 };
            bool dynamicForm = PdfObjectLookup.Resolve(objects, catalog.Items.TryGetValue("NeedsRendering", out PdfObject? rendering) ? rendering : null) is PdfBoolean { Value: true };
            analysis.XfaCanBeRemoved = hasFields && !dynamicForm;
        }
        return analysis;

        (HashSet<int> Pages, List<PdfDictionary> Widgets) Visit(PdfObject fieldObject, string parentName, string? inheritedType, int depth) {
            var pages = new HashSet<int>();
            var widgets = new List<PdfDictionary>();
            if (depth > MaximumFormFieldDepth || PdfObjectLookup.Resolve(objects, fieldObject) is not PdfDictionary node || !visited.Add(node)) return (pages, widgets);
            string partial = PdfObjectLookup.Resolve(objects, node.Items.TryGetValue("T", out PdfObject? nameValue) ? nameValue : null) is PdfStringObj text ? text.Value : string.Empty;
            string name = parentName.Length == 0 ? partial : partial.Length == 0 ? parentName : parentName + "." + partial;
            string? fieldType = TryReadName(objects, node, "FT") ?? inheritedType;
            PdfArray? kids = PdfObjectLookup.Resolve(objects, node.Items.TryGetValue("Kids", out PdfObject? kidsValue) ? kidsValue : null) as PdfArray;
            bool isSignature = string.Equals(fieldType, "Sig", StringComparison.Ordinal);
            var ownWidgets = new List<PdfDictionary>();
            var owners = new List<PdfDictionary> { node };
            if (kids is null || kids.Items.Count == 0) ownWidgets.Add(node);
            else {
                foreach (PdfObject kid in kids.Items) {
                    if (PdfObjectLookup.Resolve(objects, kid) is not PdfDictionary child) continue;
                    if (child.Items.ContainsKey("T")) {
                        (HashSet<int> childPages, List<PdfDictionary> childWidgets) = Visit(kid, name, fieldType, depth + 1);
                        pages.UnionWith(childPages);
                        widgets.AddRange(childWidgets);
                    } else {
                        ownWidgets.Add(child);
                        owners.Add(child);
                    }
                }
            }
            foreach (PdfDictionary widget in ownWidgets) {
                if (pageOfWidget.TryGetValue(widget, out int page)) pages.Add(page);
                widgets.Add(widget);
            }
            if (isSignature) return (pages, widgets);

            int[] pagesInScope = pages.Where(scope.Contains).ToArray();
            if (pagesInScope.Length == 0 && !(pages.Count == 0 && includeFieldsWithoutWidgets)) return (pages, widgets);
            List<PdfDictionary> holders = owners.Where(HoldsValue).ToList();
            List<PdfDictionary> staleWidgets = ownWidgets.Where(widget => pageOfWidget.TryGetValue(widget, out int page) && scope.Contains(page) && WidgetHolds(widget)).ToList();
            bool nameHolds = criteria.Any(criterion => criterion.Matches(name));
            if (holders.Count == 0 && staleWidgets.Count == 0 && !nameHolds) return (pages, widgets);
            var affected = new PdfRedactionAffectedField(name) { NameHoldsSearch = nameHolds };
            affected.ValueOwners.AddRange(holders);
            // A cleared value leaves every widget it reaches showing old text, so all of them follow.
            affected.Widgets.AddRange(holders.Count > 0 || nameHolds ? widgets : staleWidgets);
            affected.FirstWidget = affected.Widgets.FirstOrDefault(widget => pageOfWidget.TryGetValue(widget, out int page) && scope.Contains(page)) ?? affected.Widgets.FirstOrDefault();
            foreach (PdfDictionary widget in affected.Widgets) {
                if (pageOfWidget.TryGetValue(widget, out int page) && scope.Contains(page)) affected.Pages.Add(page);
            }
            analysis.Fields.Add(affected);
            return (pages, widgets);
        }

        bool HoldsValue(PdfDictionary owner) =>
            FieldValueKeys.Concat(FieldDataKeys).Any(key => owner.Items.TryGetValue(key, out PdfObject? value) && ValueHolds(value)) ||
            PdfObjectLookup.Resolve(objects, owner.Items.TryGetValue("Opt", out PdfObject? options) ? options : null) is PdfArray array && array.Items.Any(ValueHolds);

        bool WidgetHolds(PdfDictionary widget) {
            if (FieldDataKeys.Any(key => widget.Items.TryGetValue(key, out PdfObject? data) && ValueHolds(data))) return true;
            if (ResolveDictionary(objects, widget.Items.TryGetValue("MK", out PdfObject? characteristics) ? characteristics : null) is PdfDictionary mk &&
                CaptionKeys.Any(key => mk.Items.TryGetValue(key, out PdfObject? caption) && ValueHolds(caption))) return true;
            if (ResolveDictionary(objects, widget.Items.TryGetValue("AP", out PdfObject? appearanceValue) ? appearanceValue : null) is PdfDictionary appearances &&
                appearances.Items.Values.Any(value => ResolveDictionary(objects, value) is PdfDictionary states && states.Items.Keys.Any(state => criteria.Any(criterion => criterion.Matches(state))))) return true;
            if (anyPage is null) return false;
            PdfDictionary resources = GetInheritedDictionary(objects, anyPage, "Resources") ?? new PdfDictionary();
            return GetAppearanceStreams(objects, widget).Any(appearance => StreamHoldsSearchedText(objects, appearance,
                ResolveDictionary(objects, appearance.Dictionary.Items.TryGetValue("Resources", out PdfObject? own) ? own : null) ?? resources, anyPage, criteria, limits));
        }

        bool ValueHolds(PdfObject? value) {
            switch (PdfObjectLookup.Resolve(objects, value)) {
                case PdfStringObj text: return criteria.Any(criterion => criterion.Matches(text.Value));
                case PdfName name: return criteria.Any(criterion => criterion.Matches(name.Name));
                case PdfArray array: return array.Items.Any(ValueHolds);
                case PdfStream stream:
                    if (!TryDecodeContent(stream, objects, limits, out string content)) return false;
                    return TextHolds(content) || TextHolds(PdfTextString.Decode(PdfEncoding.Latin1GetBytes(content)));
                default: return false;
            }
        }

        bool TextHolds(string text) {
            string plain = System.Net.WebUtility.HtmlDecode(text);
            string withoutTags = XmlTag.Replace(text, string.Empty);
            return criteria.Any(criterion => criterion.Matches(text) || criterion.Matches(plain) || criterion.Matches(System.Net.WebUtility.HtmlDecode(withoutTags)));
        }

        // Undecodable packets can't be proven clean, so they count as holding the text.
        bool XfaHoldsSearch(PdfObject? xfa) {
            switch (xfa) {
                case PdfStream stream: return !TryDecodeContent(stream, objects, limits, out string content) || TextHolds(content);
                case PdfStringObj text: return TextHolds(text.Value);
                case PdfArray array: return array.Items.Any(item => PdfObjectLookup.Resolve(objects, item) is PdfStream packet && XfaHoldsSearch(packet));
                default: return false;
            }
        }
    }

    /// <summary>
    /// Clears every affected value whole. Cutting the word out can't be proven safe against field formats, comb
    /// layouts and calculation scripts. Fields and widgets stay, and appearances keep their border and background.
    /// </summary>
    private static bool ClearFormFieldsHoldingSearch(Dictionary<int, PdfIndirectObject> objects, PdfDictionary catalog, PdfReadDocument document,
        PdfRedactionArea[] areas, PdfReadLimits limits, ref int nextObjectNumber) {
        PdfRedactionSearchCriteria[] criteria = GetSearchCriteria(areas);
        if (criteria.Length == 0) return false;
        PdfRedactionFormAnalysis analysis = AnalyzeFormFields(objects, catalog, document.Pages.Select(static page => page.ObjectNumber).ToArray(), criteria,
            new HashSet<int>(areas.Where(static area => area.SearchCriteria is not null).Select(static area => area.PageNumber)),
            criteria.Any(static criterion => criterion.CoversAllPages));
        bool changed = false;
        bool needsAppearances = false;
        if (analysis.Fields.Any(static field => field.NameHoldsSearch)) throw new InvalidOperationException("A form field name holds searched text and cannot be changed safely.");
        foreach (PdfRedactionAffectedField field in analysis.Fields) {
            foreach (PdfDictionary owner in field.ValueOwners) {
                changed = true;
                ClearFieldValues(owner);
            }
            foreach (PdfDictionary widget in field.Widgets) {
                changed = true;
                ClearWidget(widget, ref nextObjectNumber);
            }
        }
        if (PdfObjectLookup.Resolve(objects, catalog.Items.TryGetValue("AcroForm", out PdfObject? formValue) ? formValue : null) is PdfDictionary acroForm) {
            if (analysis.XfaHoldsSearch && acroForm.Items.Remove("XFA")) changed = true;
            // A list box redraws its remaining options; other fields already show their cleared state.
            if (needsAppearances) acroForm.Items["NeedAppearances"] = new PdfBoolean(true);
        }
        return changed;

        bool Holds(PdfObject? value) => PdfObjectLookup.Resolve(objects, value) switch {
            PdfStringObj text => criteria.Any(criterion => criterion.Matches(text.Value)),
            PdfName name => criteria.Any(criterion => criterion.Matches(name.Name)),
            PdfArray array => array.Items.Any(Holds),
            PdfStream stream => !TryDecodeContent(stream, objects, limits, out string content) ||
                criteria.Any(criterion => criterion.Matches(content) || criterion.Matches(PdfTextString.Decode(PdfEncoding.Latin1GetBytes(content)))),
            _ => false
        };

        void ClearFieldValues(PdfDictionary owner) {
            foreach (string key in FieldValueKeys.Concat(FieldDataKeys)) {
                if (!owner.Items.TryGetValue(key, out PdfObject? value) || !Holds(value)) continue;
                // A button's value names a state, so it goes to the off state rather than to nothing.
                if (PdfObjectLookup.Resolve(objects, value) is PdfName) owner.Items[key] = new PdfName("Off");
                else owner.Items.Remove(key);
            }
            if (PdfObjectLookup.Resolve(objects, owner.Items.TryGetValue("Opt", out PdfObject? optionsValue) ? optionsValue : null) is not PdfArray options || !options.Items.Any(Holds)) return;
            var kept = new PdfArray();
            var indexMap = new Dictionary<int, int>();
            bool isButton = string.Equals(TryReadName(objects, owner, "FT"), "Btn", StringComparison.Ordinal);
            for (int index = 0; index < options.Items.Count; index++) {
                // A button's options line up with its widgets by position, so an entry is blanked, not dropped.
                if (Holds(options.Items[index]) && isButton) kept.Items.Add(new PdfStringObj(string.Empty));
                else if (Holds(options.Items[index])) continue;
                else kept.Items.Add(options.Items[index]);
                indexMap[index] = kept.Items.Count - 1;
            }
            owner.Items["Opt"] = kept;
            owner.Items.Remove("TI");
            if (PdfObjectLookup.Resolve(objects, owner.Items.TryGetValue("I", out PdfObject? selection) ? selection : null) is PdfArray selected) {
                var remapped = new PdfArray();
                foreach (PdfObject item in selected.Items) {
                    if (PdfObjectLookup.Resolve(objects, item) is PdfNumber number && indexMap.TryGetValue((int)number.Value, out int mapped)) remapped.Items.Add(new PdfNumber(mapped));
                }
                if (remapped.Items.Count > 0) owner.Items["I"] = remapped; else owner.Items.Remove("I");
            }
            needsAppearances = true;
        }

        void ClearWidget(PdfDictionary widget, ref int number) {
            foreach (string key in FieldDataKeys) {
                if (widget.Items.TryGetValue(key, out PdfObject? data) && Holds(data)) widget.Items.Remove(key);
            }
            if (ResolveDictionary(objects, widget.Items.TryGetValue("MK", out PdfObject? characteristics) ? characteristics : null) is PdfDictionary mk) {
                foreach (string key in CaptionKeys) {
                    if (mk.Items.TryGetValue(key, out PdfObject? caption) && Holds(caption)) mk.Items.Remove(key);
                }
            }
            if (widget.Items.TryGetValue("AS", out PdfObject? state) && Holds(state)) widget.Items["AS"] = new PdfName("Off");
            if (ResolveDictionary(objects, widget.Items.TryGetValue("AP", out PdfObject? appearanceValue) ? appearanceValue : null) is not PdfDictionary appearances) return;
            foreach (string key in appearances.Items.Keys.ToArray()) {
                PdfObject appearance = appearances.Items[key];
                if (PdfObjectLookup.Resolve(objects, appearance) is PdfStream stream) {
                    appearances.Items[key] = StripAppearanceText(stream, ref number);
                } else if (ResolveDictionary(objects, appearance) is PdfDictionary states) {
                    var rewritten = new PdfDictionary();
                    int renamed = 0;
                    foreach (KeyValuePair<string, PdfObject> entry in states.Items) {
                        // State names are data too, so one holding the word gets a neutral name. Off stays Off.
                        bool stateHolds = criteria.Any(criterion => criterion.Matches(entry.Key));
                        string name = stateHolds ? "Choice" + (++renamed).ToString(System.Globalization.CultureInfo.InvariantCulture) : entry.Key;
                        // A tick is drawn as text, so a state keeps it unless it shows the word itself.
                        rewritten.Items[name] = PdfObjectLookup.Resolve(objects, entry.Value) is PdfStream stateStream && StateShowsSearch(stateStream)
                            ? StripAppearanceText(stateStream, ref number) : entry.Value;
                    }
                    appearances.Items[key] = rewritten;
                }
            }
        }

        bool StateShowsSearch(PdfStream stream) {
            if (!objects.TryGetValue(document.Pages[0].ObjectNumber, out PdfIndirectObject? first) || first.Value is not PdfDictionary page) return true;
            PdfDictionary resources = ResolveDictionary(objects, stream.Dictionary.Items.TryGetValue("Resources", out PdfObject? own) ? own : null) ?? new PdfDictionary();
            return StreamHoldsSearchedText(objects, stream, resources, page, criteria, limits);
        }

        PdfObject StripAppearanceText(PdfStream stream, ref int number) {
            if (!TryDecodeContent(stream, objects, limits, out string content)) {
                // Unreadable appearance: draw nothing rather than risk showing the old value.
                content = string.Empty;
            }
            var builder = new StringBuilder(content.Length);
            int cursor = 0;
            foreach (TextObjectSpan span in EnumerateTextObjectSpans(content, limits)) {
                builder.Append(content, cursor, span.Index - cursor);
                cursor = span.Index + span.Length;
            }
            builder.Append(content, cursor, content.Length - cursor);
            int objectNumber = number++;
            objects[objectNumber] = new PdfIndirectObject(objectNumber, 0, new PdfStream(CleanStreamDictionary(stream.Dictionary), PdfEncoding.Latin1GetBytes(builder.ToString())));
            return new PdfReference(objectNumber, 0);
        }
    }

    /// <summary>The full name of the field a widget belongs to, which survives the renumbering a rewrite does.</summary>
    internal static string GetFieldName(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget) {
        var parts = new List<string>();
        PdfDictionary? current = widget;
        for (int depth = 0; current is not null && depth <= MaximumFormFieldDepth; depth++) {
            if (PdfObjectLookup.Resolve(objects, current.Items.TryGetValue("T", out PdfObject? name) ? name : null) is PdfStringObj text) parts.Insert(0, text.Value);
            current = PdfObjectLookup.Resolve(objects, current.Items.TryGetValue("Parent", out PdfObject? parent) ? parent : null) as PdfDictionary;
        }
        return string.Join(".", parts);
    }

    internal static PdfRedactionFormAnalysis AnalyzeFormFields(PdfReadDocument document, IEnumerable<PdfRedactionArea> areas) {
        PdfRedactionArea[] searched = areas.Where(static area => area.SearchCriteria is not null).ToArray();
        return AnalyzeFormFields(document.Objects, document.CatalogDictionary, document.Pages.Select(static page => page.ObjectNumber).ToArray(), GetSearchCriteria(searched),
            new HashSet<int>(searched.Select(static area => area.PageNumber)), searched.Any(static area => area.SearchCriteria!.CoversAllPages));
    }
}
