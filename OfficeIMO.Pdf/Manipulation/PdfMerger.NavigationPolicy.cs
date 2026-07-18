namespace OfficeIMO.Pdf;

internal static partial class PdfMerger {
    private static byte[] ApplyNamedDestinationPolicy(
        byte[] merged,
        IReadOnlyList<ImportedSource> sources,
        int primarySourceIndex,
        PdfMergeStructureMode mode,
        PdfMergeCollisionMode collisionMode,
        List<PdfMergeDecision> decisions) {
        int incomingCount = sources.Where((source, index) => index != primarySourceIndex).Sum(static source => source.Document.NamedDestinations.Count);
        switch (mode) {
            case PdfMergeStructureMode.KeepPrimary:
                merged = RewriteNamedDestinationLinksOnly(merged, GetIncomingPageIndexes(sources, primarySourceIndex));
                decisions.Add(new PdfMergeDecision("NamedDestinations", mode, "Kept primary named destinations.", droppedCount: incomingCount));
                return merged;
            case PdfMergeStructureMode.RejectIncoming:
                if (incomingCount > 0) throw new InvalidOperationException("PDF merge policy rejected " + incomingCount + " incoming named destination(s).");
                decisions.Add(new PdfMergeDecision("NamedDestinations", mode, "No incoming named destinations were present."));
                return merged;
            case PdfMergeStructureMode.Drop:
                merged = RewriteNamedDestinationNavigation(merged, Array.Empty<MergedNamedDestination>(), null, null, removeNamedDestinationLinks: true);
                decisions.Add(new PdfMergeDecision("NamedDestinations", mode, "Removed named destinations and links that depended on them."));
                return merged;
            case PdfMergeStructureMode.Combine:
                var renamed = new List<string>();
                int dropped = 0;
                Dictionary<int, Dictionary<string, string>> renamesBySource;
                Dictionary<int, HashSet<string>> droppedBySource;
                IReadOnlyList<MergedNamedDestination> destinations = CombineNamedDestinations(sources, collisionMode, renamed, ref dropped, out renamesBySource, out droppedBySource);
                Dictionary<int, Dictionary<string, string>> renamesByPage = ExpandDestinationRenamesByPage(sources, renamesBySource);
                Dictionary<int, HashSet<string>> droppedByPage = ExpandDestinationDropsByPage(sources, droppedBySource);
                merged = RewriteNamedDestinationNavigation(merged, destinations, renamesByPage, droppedByPage, removeNamedDestinationLinks: false);
                ValidateNamedDestinations(merged, destinations);
                decisions.Add(new PdfMergeDecision("NamedDestinations", mode, "Combined named destinations, retargeted pages, and updated renamed incoming links.", incomingCount - dropped, dropped, renamed.AsReadOnly()));
                return merged;
            default:
                throw new ArgumentOutOfRangeException(nameof(mode));
        }
    }

    private static byte[] ApplyPageLabelPolicy(
        byte[] merged,
        IReadOnlyList<ImportedSource> sources,
        int primarySourceIndex,
        PdfMergeStructureMode mode,
        List<PdfMergeDecision> decisions) {
        int incomingCount = sources.Where((source, index) => index != primarySourceIndex).Sum(static source => source.Document.PageLabels.Count);
        switch (mode) {
            case PdfMergeStructureMode.KeepPrimary:
                decisions.Add(new PdfMergeDecision("PageLabels", mode, "Kept primary page-label rules.", droppedCount: incomingCount));
                return merged;
            case PdfMergeStructureMode.RejectIncoming:
                if (incomingCount > 0) throw new InvalidOperationException("PDF merge policy rejected " + incomingCount + " incoming page-label rule(s).");
                decisions.Add(new PdfMergeDecision("PageLabels", mode, "No incoming page-label rules were present."));
                return merged;
            case PdfMergeStructureMode.Drop:
                merged = RewriteCatalogNavigation(merged, null, Array.Empty<MergedPageLabel>());
                decisions.Add(new PdfMergeDecision("PageLabels", mode, "Removed page-label rules."));
                return merged;
            case PdfMergeStructureMode.Combine:
                IReadOnlyList<MergedPageLabel> labels = BuildMergedPageLabels(sources);
                merged = RewriteCatalogNavigation(merged, null, labels);
                ValidatePageLabels(merged, labels);
                decisions.Add(new PdfMergeDecision("PageLabels", mode, "Combined page-label rules at their merged page offsets.", incomingCount));
                return merged;
            default:
                throw new ArgumentOutOfRangeException(nameof(mode));
        }
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<MergedNamedDestination> CombineNamedDestinations(
        IReadOnlyList<ImportedSource> sources,
        PdfMergeCollisionMode collisionMode,
        List<string> renamed,
        ref int dropped,
        out Dictionary<int, Dictionary<string, string>> renamesBySource,
        out Dictionary<int, HashSet<string>> droppedBySource) {
        var result = new List<MergedNamedDestination>();
        var names = new HashSet<string>(StringComparer.Ordinal);
        renamesBySource = new Dictionary<int, Dictionary<string, string>>();
        droppedBySource = new Dictionary<int, HashSet<string>>();
        int pageOffset = 0;
        for (int sourceIndex = 0; sourceIndex < sources.Count; sourceIndex++) {
            foreach (PdfNamedDestination destination in sources[sourceIndex].Document.NamedDestinations) {
                if (!destination.PageNumber.HasValue) continue;
                string name = destination.Name;
                if (!names.Add(name)) {
                    if (collisionMode == PdfMergeCollisionMode.Reject) throw new InvalidOperationException("PDF named destination collision: " + name);
                    if (collisionMode == PdfMergeCollisionMode.KeepFirst) {
                        if (!droppedBySource.TryGetValue(sourceIndex, out HashSet<string>? sourceDropped)) {
                            sourceDropped = new HashSet<string>(StringComparer.Ordinal);
                            droppedBySource[sourceIndex] = sourceDropped;
                        }

                        sourceDropped.Add(name);
                        dropped++;
                        continue;
                    }
                    string renamedName = GetUniqueDestinationName(name, sourceIndex, names);
                    if (!renamesBySource.TryGetValue(sourceIndex, out Dictionary<string, string>? sourceRenames)) {
                        sourceRenames = new Dictionary<string, string>(StringComparer.Ordinal);
                        renamesBySource[sourceIndex] = sourceRenames;
                    }
                    sourceRenames[name] = renamedName;
                    renamed.Add("source " + sourceIndex + ": " + name + " -> " + renamedName);
                    name = renamedName;
                    names.Add(name);
                }
                result.Add(new MergedNamedDestination(name, destination.PageNumber.Value + pageOffset, destination.DestinationMode, destination.DestinationLeft, destination.DestinationBottom, destination.DestinationRight, destination.DestinationTop, destination.DestinationZoom));
            }
            pageOffset += sources[sourceIndex].PageObjectNumbers.Length;
        }
        return result.AsReadOnly();
    }

    private static string GetUniqueDestinationName(string name, int sourceIndex, HashSet<string> names) {
        int sequence = 1;
        while (true) {
            string candidate = name + ".source" + (sourceIndex + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) +
                (sequence == 1 ? string.Empty : "." + sequence.ToString(System.Globalization.CultureInfo.InvariantCulture));
            if (!names.Contains(candidate)) return candidate;
            sequence++;
        }
    }

    private static Dictionary<int, Dictionary<string, string>> ExpandDestinationRenamesByPage(
        IReadOnlyList<ImportedSource> sources,
        Dictionary<int, Dictionary<string, string>> renamesBySource) {
        var result = new Dictionary<int, Dictionary<string, string>>();
        int pageOffset = 0;
        for (int sourceIndex = 0; sourceIndex < sources.Count; sourceIndex++) {
            if (renamesBySource.TryGetValue(sourceIndex, out Dictionary<string, string>? renames)) {
                for (int pageIndex = 0; pageIndex < sources[sourceIndex].PageObjectNumbers.Length; pageIndex++) {
                    result[pageOffset + pageIndex] = renames;
                }
            }
            pageOffset += sources[sourceIndex].PageObjectNumbers.Length;
        }
        return result;
    }

    private static Dictionary<int, HashSet<string>> ExpandDestinationDropsByPage(
        IReadOnlyList<ImportedSource> sources,
        Dictionary<int, HashSet<string>> dropsBySource) {
        var result = new Dictionary<int, HashSet<string>>();
        int pageOffset = 0;
        for (int sourceIndex = 0; sourceIndex < sources.Count; sourceIndex++) {
            if (dropsBySource.TryGetValue(sourceIndex, out HashSet<string>? drops)) {
                for (int pageIndex = 0; pageIndex < sources[sourceIndex].PageObjectNumbers.Length; pageIndex++) {
                    result[pageOffset + pageIndex] = drops;
                }
            }

            pageOffset += sources[sourceIndex].PageObjectNumbers.Length;
        }

        return result;
    }

    private static HashSet<int> GetIncomingPageIndexes(IReadOnlyList<ImportedSource> sources, int primarySourceIndex) {
        var result = new HashSet<int>();
        int pageOffset = 0;
        for (int sourceIndex = 0; sourceIndex < sources.Count; sourceIndex++) {
            if (sourceIndex != primarySourceIndex) {
                for (int pageIndex = 0; pageIndex < sources[sourceIndex].PageObjectNumbers.Length; pageIndex++) {
                    result.Add(pageOffset + pageIndex);
                }
            }

            pageOffset += sources[sourceIndex].PageObjectNumbers.Length;
        }

        return result;
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<MergedPageLabel> BuildMergedPageLabels(IReadOnlyList<ImportedSource> sources) {
        var result = new List<MergedPageLabel>();
        int pageOffset = 0;
        foreach (ImportedSource source in sources) {
            foreach (PdfPageLabel label in source.Document.PageLabels) {
                result.Add(new MergedPageLabel(label.StartPageIndex + pageOffset, label.Style, label.Prefix, label.StartNumber));
            }
            pageOffset += source.PageObjectNumbers.Length;
        }
        return result.AsReadOnly();
    }

    private static byte[] RewriteCatalogNavigation(
        byte[] merged,
        IReadOnlyList<MergedNamedDestination>? destinations,
        IReadOnlyList<MergedPageLabel>? labels) {
        PdfReadDocument document = PdfReadDocument.Open(merged);
        return PdfDocumentObjectGraphRewriter.Rewrite(merged, null, null, (objects, security) => {
            PdfDictionary catalog = RequireCatalog(objects, security);
            if (destinations is not null) RewriteNamedDestinationCatalog(objects, catalog, document, destinations);
            if (labels is not null) RewritePageLabelCatalog(catalog, labels);
            return security.InfoObjectNumber.HasValue && objects.ContainsKey(security.InfoObjectNumber.Value) ? security.InfoObjectNumber : null;
        });
    }

    private static byte[] RewriteNamedDestinationNavigation(
        byte[] merged,
        IReadOnlyList<MergedNamedDestination> destinations,
        Dictionary<int, Dictionary<string, string>>? renamesByPage,
        Dictionary<int, HashSet<string>>? droppedByPage,
        bool removeNamedDestinationLinks) {
        PdfReadDocument document = PdfReadDocument.Open(merged);
        return PdfDocumentObjectGraphRewriter.Rewrite(merged, null, null, (objects, security) => {
            PdfDictionary catalog = RequireCatalog(objects, security);
            RewriteNamedDestinationCatalog(objects, catalog, document, destinations);
            RewriteNamedDestinationLinks(objects, document, renamesByPage, droppedByPage, removeNamedDestinationLinks, null);
            return security.InfoObjectNumber.HasValue && objects.ContainsKey(security.InfoObjectNumber.Value) ? security.InfoObjectNumber : null;
        });
    }

    private static byte[] RewriteNamedDestinationLinksOnly(byte[] merged, HashSet<int> removeAllOnPages) {
        if (removeAllOnPages.Count == 0) return merged;
        PdfReadDocument document = PdfReadDocument.Open(merged);
        return PdfDocumentObjectGraphRewriter.Rewrite(merged, null, null, (objects, security) => {
            RewriteNamedDestinationLinks(objects, document, null, null, removeAll: false, removeAllOnPages: removeAllOnPages);
            return security.InfoObjectNumber.HasValue && objects.ContainsKey(security.InfoObjectNumber.Value) ? security.InfoObjectNumber : null;
        });
    }

    private static PdfDictionary RequireCatalog(Dictionary<int, PdfIndirectObject> objects, PdfDocumentSecurityInfo security) {
        if (!security.RootObjectNumber.HasValue || !objects.TryGetValue(security.RootObjectNumber.Value, out PdfIndirectObject? root) || root.Value is not PdfDictionary catalog) {
            throw new InvalidOperationException("PDF catalog is not readable.");
        }
        return catalog;
    }

    private static void RewriteNamedDestinationCatalog(Dictionary<int, PdfIndirectObject> objects, PdfDictionary catalog, PdfReadDocument document, IReadOnlyList<MergedNamedDestination> destinations) {
        catalog.Items.Remove("Dests");
        PdfDictionary? names = catalog.Items.TryGetValue("Names", out PdfObject? namesObject) ? ResolveDictionary(objects, namesObject) : null;
        if (names != null) names.Items.Remove("Dests");
        if (destinations.Count == 0) return;
        if (names == null) { names = new PdfDictionary(); catalog.Items["Names"] = names; }
        var values = new PdfArray();
        foreach (MergedNamedDestination destination in destinations.OrderBy(static item => item.Name, StringComparer.Ordinal)) {
            values.Items.Add(new PdfStringObj(destination.Name, true));
            values.Items.Add(BuildDestinationArray(document, destination));
        }
        var tree = new PdfDictionary(); tree.Items["Names"] = values; names.Items["Dests"] = tree;
    }

    private static PdfArray BuildDestinationArray(PdfReadDocument document, MergedNamedDestination destination) {
        var array = new PdfArray();
        array.Items.Add(new PdfReference(document.Pages[destination.PageNumber - 1].ObjectNumber, 0));
        PdfOpenActionDestinationMode mode = destination.Mode ?? PdfOpenActionDestinationMode.Xyz;
        switch (mode) {
            case PdfOpenActionDestinationMode.Fit: array.Items.Add(new PdfName("Fit")); break;
            case PdfOpenActionDestinationMode.FitHorizontal: array.Items.Add(new PdfName("FitH")); AddNumberOrNull(array, destination.Top); break;
            case PdfOpenActionDestinationMode.FitVertical: array.Items.Add(new PdfName("FitV")); AddNumberOrNull(array, destination.Left); break;
            case PdfOpenActionDestinationMode.FitRectangle:
                array.Items.Add(new PdfName("FitR")); AddNumberOrNull(array, destination.Left); AddNumberOrNull(array, destination.Bottom); AddNumberOrNull(array, destination.Right); AddNumberOrNull(array, destination.Top); break;
            case PdfOpenActionDestinationMode.FitBoundingBox: array.Items.Add(new PdfName("FitB")); break;
            case PdfOpenActionDestinationMode.FitBoundingBoxHorizontal: array.Items.Add(new PdfName("FitBH")); AddNumberOrNull(array, destination.Top); break;
            case PdfOpenActionDestinationMode.FitBoundingBoxVertical: array.Items.Add(new PdfName("FitBV")); AddNumberOrNull(array, destination.Left); break;
            default:
                array.Items.Add(new PdfName("XYZ")); AddNumberOrNull(array, destination.Left); AddNumberOrNull(array, destination.Top); AddNumberOrNull(array, destination.Zoom); break;
        }
        return array;
    }

    private static void AddNumberOrNull(PdfArray array, double? value) => array.Items.Add(value.HasValue ? new PdfNumber(value.Value) : PdfNull.Instance);

    private static void RewritePageLabelCatalog(PdfDictionary catalog, IReadOnlyList<MergedPageLabel> labels) {
        catalog.Items.Remove("PageLabels");
        if (labels.Count == 0) return;
        var nums = new PdfArray();
        foreach (MergedPageLabel label in labels.OrderBy(static item => item.StartPageIndex)) {
            nums.Items.Add(new PdfNumber(label.StartPageIndex));
            var dictionary = new PdfDictionary();
            if (!string.IsNullOrEmpty(label.Style)) dictionary.Items["S"] = new PdfName(label.Style!);
            if (label.Prefix != null) dictionary.Items["P"] = new PdfStringObj(label.Prefix, true);
            if (label.StartNumber.HasValue) dictionary.Items["St"] = new PdfNumber(label.StartNumber.Value);
            nums.Items.Add(dictionary);
        }
        var tree = new PdfDictionary(); tree.Items["Nums"] = nums; catalog.Items["PageLabels"] = tree;
    }

    private static void RewriteNamedDestinationLinks(
        Dictionary<int, PdfIndirectObject> objects,
        PdfReadDocument document,
        Dictionary<int, Dictionary<string, string>>? renamesByPage,
        Dictionary<int, HashSet<string>>? droppedByPage,
        bool removeAll,
        HashSet<int>? removeAllOnPages) {
        for (int pageIndex = 0; pageIndex < document.Pages.Count; pageIndex++) {
            Dictionary<string, string>? renames = null;
            HashSet<string>? dropped = null;
            bool removePageLinks = removeAll || (removeAllOnPages is not null && removeAllOnPages.Contains(pageIndex));
            renamesByPage?.TryGetValue(pageIndex, out renames);
            droppedByPage?.TryGetValue(pageIndex, out dropped);
            if (!removePageLinks && renames is null && dropped is null) continue;
            PdfReadPage page = document.Pages[pageIndex];
            if (!objects.TryGetValue(page.ObjectNumber, out PdfIndirectObject? pageObject) || pageObject.Value is not PdfDictionary pageDictionary ||
                !pageDictionary.Items.TryGetValue("Annots", out PdfObject? annotationsObject) || ResolveObject(objects, annotationsObject) is not PdfArray annotations) continue;
            var rewritten = new PdfArray();
            foreach (PdfObject annotationObject in annotations.Items) {
                PdfDictionary? annotation = ResolveDictionary(objects, annotationObject);
                if (annotation != null && TryRewriteNamedDestinationLink(objects, annotation, renames, dropped, removePageLinks, out bool removeAnnotation)) {
                    if (!removeAnnotation) rewritten.Items.Add(annotationObject);
                } else {
                    rewritten.Items.Add(annotationObject);
                }
            }
            pageDictionary.Items["Annots"] = rewritten;
        }
    }

    private static bool TryRewriteNamedDestinationLink(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary annotation,
        Dictionary<string, string>? renames,
        HashSet<string>? dropped,
        bool remove,
        out bool removeAnnotation) {
        removeAnnotation = false;
        if (annotation.Get<PdfName>("Subtype")?.Name != "Link") return false;
        if (annotation.Items.TryGetValue("Dest", out PdfObject? direct) && TryRewriteDestinationToken(objects, direct, renames, dropped, remove, out PdfObject? replacement, out removeAnnotation)) {
            if (!removeAnnotation) annotation.Items["Dest"] = replacement!;
            return true;
        }
        if (annotation.Items.TryGetValue("A", out PdfObject? actionObject) && ResolveDictionary(objects, actionObject) is PdfDictionary action &&
            action.Get<PdfName>("S")?.Name == "GoTo" && action.Items.TryGetValue("D", out PdfObject? actionDestination) &&
            TryRewriteDestinationToken(objects, actionDestination, renames, dropped, remove, out replacement, out removeAnnotation)) {
            if (!removeAnnotation) action.Items["D"] = replacement!;
            return true;
        }
        return false;
    }

    private static bool TryRewriteDestinationToken(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject token,
        Dictionary<string, string>? renames,
        HashSet<string>? dropped,
        bool remove,
        out PdfObject? replacement,
        out bool removeToken) {
        PdfObject? resolved = ResolveObject(objects, token); string? name = resolved is PdfStringObj text ? text.Value : resolved is PdfName pdfName ? pdfName.Name : null;
        replacement = null;
        removeToken = false;
        if (name == null) return false;
        if (remove || (dropped is not null && dropped.Contains(name))) {
            removeToken = true;
            return true;
        }
        if (renames == null || !renames.TryGetValue(name, out string? renamed)) return false;
        replacement = resolved is PdfName ? new PdfName(renamed) : new PdfStringObj(renamed, true);
        return true;
    }

    private static PdfObject? ResolveObject(Dictionary<int, PdfIndirectObject> objects, PdfObject? value) {
        var visited = new HashSet<int>();
        while (value is PdfReference reference && visited.Add(reference.ObjectNumber) && objects.TryGetValue(reference.ObjectNumber, out PdfIndirectObject? indirect)) value = indirect.Value;
        return value;
    }

    private static PdfDictionary? ResolveDictionary(Dictionary<int, PdfIndirectObject> objects, PdfObject? value) => ResolveObject(objects, value) as PdfDictionary;

    private static void ValidateNamedDestinations(byte[] merged, IReadOnlyList<MergedNamedDestination> expected) {
        IReadOnlyList<PdfNamedDestination> actual = PdfReadDocument.Open(merged).NamedDestinations;
        if (actual.Count != expected.Count) throw new InvalidOperationException("PDF named-destination merge validation failed; the artifact was not returned.");
        foreach (MergedNamedDestination destination in expected) {
            PdfNamedDestination? found = actual.SingleOrDefault(item => string.Equals(item.Name, destination.Name, StringComparison.Ordinal));
            if (found == null || found.PageNumber != destination.PageNumber) throw new InvalidOperationException("PDF named-destination merge validation failed for '" + destination.Name + "'.");
        }
    }

    private static void ValidatePageLabels(byte[] merged, IReadOnlyList<MergedPageLabel> expected) {
        IReadOnlyList<PdfPageLabel> actual = PdfReadDocument.Open(merged).PageLabels;
        if (actual.Count != expected.Count) throw new InvalidOperationException("PDF page-label merge validation failed; the artifact was not returned.");
        for (int i = 0; i < expected.Count; i++) {
            if (actual[i].StartPageIndex != expected[i].StartPageIndex || actual[i].Style != expected[i].Style || actual[i].Prefix != expected[i].Prefix || actual[i].StartNumber != expected[i].StartNumber) {
                throw new InvalidOperationException("PDF page-label merge validation failed at rule " + i.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".");
            }
        }
    }

    private sealed class MergedNamedDestination {
        internal MergedNamedDestination(string name, int pageNumber, PdfOpenActionDestinationMode? mode, double? left, double? bottom, double? right, double? top, double? zoom = null) { Name = name; PageNumber = pageNumber; Mode = mode; Left = left; Bottom = bottom; Right = right; Top = top; Zoom = zoom; }
        internal string Name { get; } internal int PageNumber { get; } internal PdfOpenActionDestinationMode? Mode { get; }
        internal double? Left { get; } internal double? Bottom { get; } internal double? Right { get; } internal double? Top { get; } internal double? Zoom { get; }
    }

    private sealed class MergedPageLabel {
        internal MergedPageLabel(int startPageIndex, string? style, string? prefix, int? startNumber) { StartPageIndex = startPageIndex; Style = style; Prefix = prefix; StartNumber = startNumber; }
        internal int StartPageIndex { get; } internal string? Style { get; } internal string? Prefix { get; } internal int? StartNumber { get; }
    }
}
