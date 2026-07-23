namespace OfficeIMO.Pdf;

internal static partial class PdfMerger {
    private static readonly char[] ViewerPreferenceSeparators = { ' ' };
    private static byte[] ApplyViewerPolicy(byte[] merged, IReadOnlyList<ImportedSource> sources, int primarySourceIndex, PdfMergeStructureMode mode, List<PdfMergeDecision> decisions) {
        int incoming = sources.Where((source, index) => index != primarySourceIndex && HasViewerState(source.Document)).Count();
        if (mode == PdfMergeStructureMode.KeepPrimary) { decisions.Add(new PdfMergeDecision("ViewerPreferences", mode, "Kept primary viewer preferences and initial view.", droppedCount: incoming)); return merged; }
        if (mode == PdfMergeStructureMode.RejectIncoming) {
            if (incoming > 0) throw new InvalidOperationException("PDF merge policy rejected incoming viewer state from " + incoming + " source(s).");
            decisions.Add(new PdfMergeDecision("ViewerPreferences", mode, "No incoming viewer state was present; kept primary viewer state."));
            return merged;
        }

        Dictionary<string, string>? values = null; string? pageMode = null; string? pageLayout = null; MergedNamedDestination? openAction = null;
        if (mode == PdfMergeStructureMode.Combine) {
            values = new Dictionary<string, string>(StringComparer.Ordinal);
            AddViewerValues(values, sources[primarySourceIndex].Document.ViewerPreferences);
            foreach (ImportedSource source in sources) AddViewerValues(values, source.Document.ViewerPreferences);
            pageMode = FirstCatalogValue(sources, primarySourceIndex, static document => document.CatalogPageMode);
            pageLayout = FirstCatalogValue(sources, primarySourceIndex, static document => document.CatalogPageLayout);
            openAction = FirstOpenAction(sources, primarySourceIndex);
        }

        byte[] output = PdfDocumentObjectGraphRewriter.Rewrite(merged, null, null, (objects, security) => {
            PdfDictionary catalog = RequireCatalog(objects, security);
            catalog.Items.Remove("ViewerPreferences"); catalog.Items.Remove("PageMode"); catalog.Items.Remove("PageLayout"); catalog.Items.Remove("OpenAction");
            if (mode == PdfMergeStructureMode.Combine) {
                if (values != null && values.Count > 0) catalog.Items["ViewerPreferences"] = BuildViewerPreferences(values);
                if (pageMode != null) catalog.Items["PageMode"] = new PdfName(pageMode);
                if (pageLayout != null) catalog.Items["PageLayout"] = new PdfName(pageLayout);
                if (openAction != null) catalog.Items["OpenAction"] = BuildDestinationArray(PdfReadDocument.Open(merged), openAction);
            }
            return security.InfoObjectNumber.HasValue && objects.ContainsKey(security.InfoObjectNumber.Value) ? security.InfoObjectNumber : null;
        });
        decisions.Add(new PdfMergeDecision("ViewerPreferences", mode, mode == PdfMergeStructureMode.Drop ? "Removed viewer preferences and initial-view state." : "Combined compatible viewer preferences with primary values winning conflicts." , incoming));
        return output;
    }

    private static byte[] ApplyCatalogStatePolicy(byte[] merged, IReadOnlyList<ImportedSource> sources, int primarySourceIndex, PdfMergeStructureMode mode, List<PdfMergeDecision> decisions) {
        int incoming = sources.Where((source, index) => index != primarySourceIndex && HasCatalogState(source)).Count();
        bool incomingOptionalContent = sources.Where((source, index) => index != primarySourceIndex)
            .Any(static source => source.CatalogState.OptionalContent != null);
        bool anyOptionalContent = sources.Any(static source => source.CatalogState.OptionalContent != null);
        if (mode == PdfMergeStructureMode.KeepPrimary && incomingOptionalContent) {
            throw new NotSupportedException("Keeping only the primary PDF optional-content configuration is blocked because incoming pages can reference hidden layers whose visibility state would be discarded.");
        }
        if (mode == PdfMergeStructureMode.KeepPrimary) { decisions.Add(new PdfMergeDecision("CatalogState", mode, "Kept primary compatible catalog state.", droppedCount: incoming)); return merged; }
        if (mode == PdfMergeStructureMode.RejectIncoming) {
            if (incoming > 0) throw new InvalidOperationException("PDF merge policy rejected incoming catalog state from " + incoming + " source(s).");
            decisions.Add(new PdfMergeDecision("CatalogState", mode, "No incoming catalog state was present; kept primary catalog state."));
            return merged;
        }
        if (mode == PdfMergeStructureMode.Combine && anyOptionalContent) {
            throw new NotSupportedException("Combining PDF optional-content configurations is blocked because rebuilding them with an all-visible default can expose content that a source intentionally hid.");
        }
        if (mode == PdfMergeStructureMode.Drop && anyOptionalContent) {
            throw new NotSupportedException("Dropping PDF optional-content configuration is blocked because page content can remain associated with hidden layers and become visible without its source visibility state.");
        }

        string? version = mode == PdfMergeStructureMode.Combine ? FirstCatalogValue(sources, primarySourceIndex, static document => document.CatalogVersion) : null;
        string? language = mode == PdfMergeStructureMode.Combine ? FirstCatalogValue(sources, primarySourceIndex, static document => document.CatalogLanguage) : null;
        byte[] output = PdfDocumentObjectGraphRewriter.Rewrite(merged, null, null, (objects, security) => {
            PdfDictionary catalog = RequireCatalog(objects, security);
            catalog.Items.Remove("Version"); catalog.Items.Remove("Lang"); catalog.Items.Remove("URI"); catalog.Items.Remove("OutputIntents"); catalog.Items.Remove("OCProperties");
            if (mode == PdfMergeStructureMode.Combine) {
                if (version != null) catalog.Items["Version"] = new PdfName(version);
                if (language != null) catalog.Items["Lang"] = new PdfStringObj(language, true);
                int nextObjectNumber = objects.Count == 0 ? 1 : objects.Keys.Max() + 1;
                var externalMap = new Dictionary<ExternalObjectKey, int>();
                PdfObject? uri = FirstCatalogObject(sources, primarySourceIndex, static state => state.CatalogUri, objects, externalMap, ref nextObjectNumber);
                if (uri != null) catalog.Items["URI"] = uri;
                PdfArray intents = CombineOutputIntents(sources, objects, externalMap, ref nextObjectNumber);
                if (intents.Items.Count > 0) catalog.Items["OutputIntents"] = intents;
            }
            return security.InfoObjectNumber.HasValue && objects.ContainsKey(security.InfoObjectNumber.Value) ? security.InfoObjectNumber : null;
        });
        string action = mode == PdfMergeStructureMode.Drop
            ? "Removed version, language, URI, output-intent, and optional-content catalog state."
            : "Combined compatible scalar, URI, and output-intent state after rejecting sources with optional-content configurations.";
        decisions.Add(new PdfMergeDecision("CatalogState", mode, action, incoming));
        return output;
    }

    private static bool HasViewerState(PdfReadDocument document) => document.ViewerPreferences != null || document.OpenAction != null || document.CatalogPageMode != null || document.CatalogPageLayout != null;
    private static bool HasCatalogState(ImportedSource source) => source.Document.CatalogVersion != null || source.Document.CatalogLanguage != null || source.CatalogState.CatalogUri != null || source.CatalogState.OutputIntents != null || source.CatalogState.OptionalContent != null;

    private static void AddViewerValues(Dictionary<string, string> target, PdfViewerPreferences? preferences) {
        if (preferences == null) return;
        foreach (var entry in preferences.Values) if (!target.ContainsKey(entry.Key)) target[entry.Key] = entry.Value;
    }

    private static string? FirstCatalogValue(IReadOnlyList<ImportedSource> sources, int primarySourceIndex, Func<PdfReadDocument, string?> selector) {
        string? primary = selector(sources[primarySourceIndex].Document); if (!string.IsNullOrEmpty(primary)) return primary;
        for (int i = 0; i < sources.Count; i++) { string? value = selector(sources[i].Document); if (!string.IsNullOrEmpty(value)) return value; }
        return null;
    }

    private static MergedNamedDestination? FirstOpenAction(IReadOnlyList<ImportedSource> sources, int primarySourceIndex) {
        int[] order = Enumerable.Range(0, sources.Count).OrderBy(index => index == primarySourceIndex ? 0 : 1).ThenBy(static index => index).ToArray();
        int[] offsets = new int[sources.Count]; int offset = 0; for (int i = 0; i < sources.Count; i++) { offsets[i] = offset; offset += sources[i].PageObjectNumbers.Length; }
        foreach (int index in order) {
            PdfDocumentOpenAction? action = sources[index].Document.OpenAction;
            if (action?.PageNumber != null) return new MergedNamedDestination("OpenAction", action.PageNumber.Value + offsets[index], action.DestinationMode, action.DestinationLeft, action.DestinationBottom, action.DestinationRight, action.DestinationTop, action.DestinationZoom);
        }
        return null;
    }

    private static PdfDictionary BuildViewerPreferences(Dictionary<string, string> values) {
        var dictionary = new PdfDictionary();
        foreach (var entry in values) dictionary.Items[entry.Key] = ParseViewerPreference(entry.Key, entry.Value);
        return dictionary;
    }

    private static PdfObject ParseViewerPreference(string key, string value) {
        if (value == "true" || value == "false") return new PdfBoolean(value == "true");
        if (key == "NumCopies" && double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double number)) return new PdfNumber(number);
        if (key == "PrintPageRange" && value.Length >= 2 && value[0] == '[' && value[value.Length - 1] == ']') {
            var array = new PdfArray(); foreach (string part in value.Substring(1, value.Length - 2).Split(ViewerPreferenceSeparators, StringSplitOptions.RemoveEmptyEntries)) if (double.TryParse(part, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out number)) array.Items.Add(new PdfNumber(number)); return array;
        }
        return new PdfName(value);
    }

    private static PdfObject? FirstCatalogObject(IReadOnlyList<ImportedSource> sources, int primarySourceIndex, Func<PdfPageExtractor.CatalogRewriteState, PdfObject?> selector, Dictionary<int, PdfIndirectObject> targetObjects, Dictionary<ExternalObjectKey, int> map, ref int next) {
        int[] order = Enumerable.Range(0, sources.Count).OrderBy(index => index == primarySourceIndex ? 0 : 1).ThenBy(static index => index).ToArray();
        foreach (int index in order) { PdfObject? value = selector(sources[index].CatalogState); if (value != null) return CloneExternalObject(value, index, sources[index].Objects, targetObjects, map, ref next); }
        return null;
    }

    private static PdfArray CombineOutputIntents(IReadOnlyList<ImportedSource> sources, Dictionary<int, PdfIndirectObject> targetObjects, Dictionary<ExternalObjectKey, int> map, ref int next) {
        var result = new PdfArray();
        for (int sourceIndex = 0; sourceIndex < sources.Count; sourceIndex++) {
            PdfObject? value = ResolveObject(sources[sourceIndex].Objects, sources[sourceIndex].CatalogState.OutputIntents);
            if (value is not PdfArray array) continue;
            foreach (PdfObject item in array.Items) result.Items.Add(CloneExternalObject(item, sourceIndex, sources[sourceIndex].Objects, targetObjects, map, ref next));
        }
        return result;
    }

    private static PdfObject CloneExternalObject(PdfObject value, int sourceIndex, Dictionary<int, PdfIndirectObject> sourceObjects, Dictionary<int, PdfIndirectObject> targetObjects, Dictionary<ExternalObjectKey, int> map, ref int next) {
        if (value is PdfReference reference) {
            var key = new ExternalObjectKey(sourceIndex, reference.ObjectNumber);
            if (!map.TryGetValue(key, out int mapped)) {
                if (!sourceObjects.TryGetValue(reference.ObjectNumber, out PdfIndirectObject? source)) throw new InvalidOperationException("Referenced catalog object was not found.");
                mapped = next++; map[key] = mapped;
                PdfObject clone = CloneExternalObject(source.Value, sourceIndex, sourceObjects, targetObjects, map, ref next);
                targetObjects[mapped] = new PdfIndirectObject(mapped, 0, clone);
            }
            return new PdfReference(mapped, 0);
        }
        if (value is PdfArray array) { var clone = new PdfArray(); foreach (PdfObject item in array.Items) clone.Items.Add(CloneExternalObject(item, sourceIndex, sourceObjects, targetObjects, map, ref next)); return clone; }
        if (value is PdfDictionary dictionary) { var clone = new PdfDictionary(); foreach (var item in dictionary.Items) clone.Items[item.Key] = CloneExternalObject(item.Value, sourceIndex, sourceObjects, targetObjects, map, ref next); return clone; }
        if (value is PdfStream stream) return new PdfStream((PdfDictionary)CloneExternalObject(stream.Dictionary, sourceIndex, sourceObjects, targetObjects, map, ref next), (byte[])stream.Data.Clone(), stream.DecodingFailed, stream.DecodingError);
        if (value is PdfStringObj text) return new PdfStringObj(text.RawBytes, text.UseTextStringEncoding);
        if (value is PdfName name) return new PdfName(name.Name); if (value is PdfNumber numeric) return new PdfNumber(numeric.Value); if (value is PdfBoolean boolean) return new PdfBoolean(boolean.Value); return PdfNull.Instance;
    }

    private readonly struct ExternalObjectKey : IEquatable<ExternalObjectKey> {
        internal ExternalObjectKey(int sourceIndex, int objectNumber) { SourceIndex = sourceIndex; ObjectNumber = objectNumber; }
        private int SourceIndex { get; } private int ObjectNumber { get; }
        public bool Equals(ExternalObjectKey other) => SourceIndex == other.SourceIndex && ObjectNumber == other.ObjectNumber;
        public override bool Equals(object? obj) => obj is ExternalObjectKey other && Equals(other);
        public override int GetHashCode() => unchecked((SourceIndex * 397) ^ ObjectNumber);
    }
}
