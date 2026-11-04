namespace OfficeIMO.Pdf;

// One alternate-text value can describe several sections through a shared property list, an inline repeat, a form
// painted twice or a structure element linked to several sections. When one of them loses content the value is text
// of redacted content, and giving the retained sections their own copy would keep it in the file, so it leaves every
// holder. Only a page-scoped search limits that: the pages it left out keep what they have. A retained section keeps
// its marks and loses only its description.
internal static partial class PdfRedactionApplier {
    private delegate IReadOnlyList<string> StructureAlternateText(int markedContentId);

    private sealed class StructureLinks {
        internal Dictionary<(int Owner, int Id), List<string>> Text { get; } = new();
        internal Dictionary<PdfDictionary, HashSet<int>> Owners { get; } = new();
    }

    private sealed class StructureNode {
        internal StructureNode(PdfDictionary element, StructureNode? parent) {
            Element = element;
            Parent = parent;
        }

        internal PdfDictionary Element { get; }
        internal StructureNode? Parent { get; }
    }

    /// <summary>Counts the reviewed sections, and adds them to the fingerprints verification compares against.</summary>
    internal static Dictionary<int, Dictionary<string, int>> CaptureAlternateTextPlacements(PdfReadDocument document, HashSet<string> sections, System.Threading.CancellationToken cancellationToken) =>
        CountAlternateTextPlacements(document.Objects, document.Pages.Select(static page => page.ObjectNumber).ToArray(), document.ReadOptions.Limits, fingerprint => {
            sections.Add(fingerprint);
            return true;
        }, cancellationToken);

    private static bool RemoveAlternateTextOfRemovedSections(Dictionary<int, PdfIndirectObject> objects, PdfRedactionPlan plan, int[] pages, PdfReadLimits limits, HashSet<PdfStream> sourceStreamIdentities, System.Threading.CancellationToken cancellationToken) {
        HashSet<string> removed = FindRemovedAlternateText(objects, pages, plan, limits, cancellationToken);
        return removed.Count > 0 && RemoveAlternateText(objects, text => removed.Contains(FingerprintAlternateText(text)), GetPageScope(plan), pages, limits, sourceStreamIdentities);
    }

    private static IEnumerable<PdfRedactionVerificationIssue> FindAlternateTextOfRemovedSections(PdfReadDocument document, PdfRedactionPlan plan, System.Threading.CancellationToken cancellationToken) {
        Dictionary<int, PdfIndirectObject> objects = document.Objects;
        PdfReadLimits limits = document.ReadOptions.Limits;
        int[] pages = document.Pages.Select(static page => page.ObjectNumber).ToArray();
        HashSet<string> removed = FindRemovedAlternateText(objects, pages, plan, limits, cancellationToken);
        if (removed.Count == 0) yield break;
        AlternateTextMatch matches = text => removed.Contains(FingerprintAlternateText(text));
        HashSet<PdfObject> excluded = FindExcludedSearchObjects(objects, pages, GetPageScope(plan), cancellationToken);
        HashSet<int> pageContentStreams = GetPageContentStreamNumbers(objects);
        foreach (PdfIndirectObject item in objects.Values) {
            cancellationToken.ThrowIfCancellationRequested();
            if (excluded.Contains(item.Value)) continue;
            bool holds = HoldsMatchingAlternateText(objects, item.Value, matches);
            if (!holds && IsContentOwner(pageContentStreams, item, out PdfStream? stream) && TryDecodeContent(stream, objects, limits, out string content))
                RewriteMarkedContentHoldingMatches(content, matches, limits, out holds);
            if (holds) yield return new PdfRedactionVerificationIssue("RedactedContentOutsidePageContent", "AlternateText",
                "Object " + item.ObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " keeps alternate text shared with removed content.");
        }
    }

    // A search limited to pages is the only request that says other pages stay as they are.
    private static PdfRedactionSearchCriteria[] GetPageScope(PdfRedactionPlan plan) =>
        plan.Areas.All(static area => area.SearchCriteria is { CoversAllPages: false }) ? GetSearchCriteria(plan.Areas) : Array.Empty<PdfRedactionSearchCriteria>();

    private static HashSet<string> FindRemovedAlternateText(Dictionary<int, PdfIndirectObject> objects, int[] pages, PdfRedactionPlan plan, PdfReadLimits limits, System.Threading.CancellationToken cancellationToken) {
        var removed = new HashSet<string>(StringComparer.Ordinal);
        if (plan.AlternateTextPlacements.Count == 0) return removed;
        Dictionary<int, Dictionary<string, int>> current = CountAlternateTextPlacements(objects, pages, limits, plan.AlternateTextSections.Contains, cancellationToken);
        foreach (KeyValuePair<int, Dictionary<string, int>> page in plan.AlternateTextPlacements) {
            current.TryGetValue(page.Key, out Dictionary<string, int>? kept);
            foreach (KeyValuePair<string, int> value in page.Value) {
                int count = 0;
                if (kept?.TryGetValue(value.Key, out count) != true || count < value.Value) removed.Add(value.Key);
            }
        }
        return removed;
    }

    private static string FingerprintAlternateText(string value) => PdfRedactionPlan.ComputeSourceSha256(Encoding.UTF8.GetBytes(value));

    // Counted per painted placement: a scrubbed copy of a form replaces one placement, and its section is the one that is gone.
    private static Dictionary<int, Dictionary<string, int>> CountAlternateTextPlacements(Dictionary<int, PdfIndirectObject> objects, int[] pages, PdfReadLimits limits,
        Func<string, bool> counts, System.Threading.CancellationToken cancellationToken) {
        StructureLinks structure = ReadStructureLinks(objects, cancellationToken);
        var forms = new Dictionary<PdfStream, Dictionary<string, int>>();
        var open = new HashSet<PdfStream>();
        var result = new Dictionary<int, Dictionary<string, int>>();
        for (int index = 0; index < pages.Length; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!objects.TryGetValue(pages[index], out PdfIndirectObject? item) || item.Value is not PdfDictionary page || !page.Items.TryGetValue("Contents", out PdfObject? contents)) continue;
            result[index + 1] = Count(ReadPageContent(objects, contents, limits), GetInheritedDictionary(objects, page, "Resources"), pages[index]);
        }
        return result;

        Dictionary<string, int> Count(string content, PdfDictionary? resources, int owner) {
            var total = new Dictionary<string, int>(StringComparer.Ordinal);
            StructureAlternateText structureText = id => structure.Text.TryGetValue((owner, id), out List<string>? values) ? values : Array.Empty<string>();
            foreach (AlternateTextScope scope in FindAlternateTextScopes(content, ReadPropertyAlternateText(objects, resources), limits, structureText)) {
                foreach (string value in scope.Values) {
                    if (counts(scope.GetFingerprint(value))) Add(total, FingerprintAlternateText(value), 1);
                }
            }
            PdfDictionary? xObjects = resources is null ? null : ResolveDictionary(objects, resources.Items.TryGetValue("XObject", out PdfObject? listed) ? listed : null);
            if (xObjects is null) return total;
            PdfContentStreamInterpreter.Interpret(content, limits.MaxContentOperations, operation => {
                if (operation.Name != "Do" || operation.Operands.Count == 0 || operation.Operands[operation.Operands.Count - 1] is not string name ||
                    !xObjects.Items.TryGetValue(name, out PdfObject? entry) || PdfObjectLookup.Resolve(objects, entry) is not PdfStream form ||
                    !string.Equals(form.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal)) return;
                foreach (KeyValuePair<string, int> inner in CountForm(form, (entry as PdfReference)?.ObjectNumber ?? 0, resources!)) Add(total, inner.Key, inner.Value);
            }, maxNestingDepth: limits.MaxContentNestingDepth, maxOperands: limits.MaxContentOperands);
            return total;
        }

        Dictionary<string, int> CountForm(PdfStream form, int owner, PdfDictionary inherited) {
            if (forms.TryGetValue(form, out Dictionary<string, int>? known)) return known;
            if (!open.Add(form) || !TryDecodeContent(form, objects, limits, out string content)) return new Dictionary<string, int>(StringComparer.Ordinal);
            PdfDictionary resources = ResolveDictionary(objects, form.Dictionary.Items.TryGetValue("Resources", out PdfObject? own) ? own : null) ?? inherited;
            Dictionary<string, int> total = Count(content, resources, owner);
            open.Remove(form);
            return forms[form] = total;
        }

        static void Add(Dictionary<string, int> total, string key, int count) =>
            total[key] = total.TryGetValue(key, out int existing) ? existing + count : count;
    }

    // A structure element's alternate text covers every marked section under it, on whichever page or form each one is.
    private static StructureLinks ReadStructureLinks(Dictionary<int, PdfIndirectObject> objects, System.Threading.CancellationToken cancellationToken) {
        var links = new StructureLinks();
        PdfDictionary? root = ResolveDictionary(objects, PdfSyntax.FindCatalog(objects, null, cancellationToken: cancellationToken)?.Items.TryGetValue("StructTreeRoot", out PdfObject? tree) == true ? tree : null);
        if (root is null || !root.Items.TryGetValue("K", out PdfObject? kids)) return links;
        var visited = new HashSet<PdfDictionary>();
        var pending = new Stack<(PdfObject Value, StructureNode? Node, int? Owner)>();
        pending.Push((kids, null, null));
        while (pending.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            (PdfObject value, StructureNode? node, int? owner) = pending.Pop();
            PdfObject? resolved = PdfObjectLookup.Resolve(objects, value);
            if (resolved is PdfNumber number) Link(node, owner, (int)number.Value);
            else if (resolved is PdfArray array) foreach (PdfObject item in array.Items) pending.Push((item, node, owner));
            else if (resolved is PdfDictionary dictionary) {
                string? type = dictionary.Get<PdfName>("Type")?.Name;
                if (type == "OBJR") continue;
                if (type == "MCR") {
                    if (dictionary.Get<PdfNumber>("MCID") is PdfNumber id) Link(node, OwnerOf(dictionary, "Stm") ?? OwnerOf(dictionary, "Pg") ?? owner, (int)id.Value);
                } else if (visited.Add(dictionary) && dictionary.Items.TryGetValue("K", out PdfObject? children)) {
                    pending.Push((children, new StructureNode(dictionary, node), OwnerOf(dictionary, "Pg") ?? owner));
                }
            }
        }
        return links;

        void Link(StructureNode? node, int? owner, int id) {
            if (owner is not int linked) return;
            if (!links.Text.TryGetValue((linked, id), out List<string>? values)) links.Text[(linked, id)] = values = new List<string>();
            for (StructureNode? current = node; current is not null; current = current.Parent) {
                if (!links.Owners.TryGetValue(current.Element, out HashSet<int>? owners)) links.Owners[current.Element] = owners = new HashSet<int>();
                owners.Add(linked);
                foreach (string key in AlternateTextKeys) {
                    if (current.Element.Items.TryGetValue(key, out PdfObject? entry) && PdfObjectLookup.Resolve(objects, entry) is PdfStringObj text) values.Add(text.Value);
                }
            }
        }

        static int? OwnerOf(PdfDictionary dictionary, string key) => dictionary.Items.TryGetValue(key, out PdfObject? value) && value is PdfReference reference ? reference.ObjectNumber : null;
    }
}
