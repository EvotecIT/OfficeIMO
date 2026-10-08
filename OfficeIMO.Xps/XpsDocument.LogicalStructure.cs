namespace OfficeIMO.Xps;

public sealed partial class XpsDocument {
    /// <summary>Reconstructs native stories in DocumentStructure reference order, merging continued paragraphs, lists and tables.</summary>
    /// <remarks>Story page numbers use the adopted payload-global interpretation. Missing metadata does not produce inferred semantics.</remarks>
    public XpsLogicalStructure ReadLogicalStructure(CancellationToken cancellationToken = default) {
        var budget = new XpsStoryFragmentsReader.Budget(cancellationToken);
        var pages = _pages.Select((page, index) => ReadPageContent(page, index, budget)).ToList();
        var diagnostics = new List<string>();
        void Loss(string value) { if (diagnostics.Count < 100 && !diagnostics.Contains(value)) diagnostics.Add(value); }
        void Attributes(XElement element, params string[] allowed) {
            foreach (var attribute in element.Attributes()) {
                budget.Charge();
                if (attribute.IsNamespaceDeclaration || attribute.Name == XNamespace.Xml + "lang") continue;
                if (attribute.Name.NamespaceName.Length != 0 || !allowed.Contains(attribute.Name.LocalName))
                    Loss("Unsupported structure attribute: " + element.Name.LocalName + "." + attribute.Name);
            }
            ValidateStructureLiteralText(element);
        }
        foreach (var page in pages) foreach (string diagnostic in page.Diagnostics) Loss(diagnostic);
        var stories = new List<XpsLogicalStory>(); var used = new HashSet<XpsStoryFragment>();
        foreach (var native in ReadDocumentStructures(cancellationToken, activeOnly: true)) {
            Attributes(native.Markup);
            var names = new HashSet<string>(StringComparer.Ordinal);
            foreach (var story in native.Markup.Elements()) {
                budget.Charge();
                if (story.Name == StructureNamespace + "DocumentStructure.Outline") continue;
                if (story.Name != StructureNamespace + "Story") { Loss("Unsupported document-structure element: " + story.Name); continue; }
                Attributes(story, "StoryName");
                string name = (string?)story.Attribute("StoryName") ?? throw new InvalidDataException("Missing story name.");
                if (!names.Add(name)) throw new InvalidDataException("Duplicate story name in DocumentStructure.");
                var fragments = new List<XpsStoryFragment>();
                if (!story.HasElements) throw new InvalidDataException("A declared story requires fragment references.");
                foreach (var reference in story.Elements()) {
                    budget.Charge();
                    if (reference.Name != StructureNamespace + "StoryFragmentReference") { Loss("Unsupported story reference: " + reference.Name); continue; }
                    Attributes(reference, "Page", "FragmentName");
                    if (reference.HasElements) throw new InvalidDataException("StoryFragmentReference cannot contain elements.");
                    if (!int.TryParse((string?)reference.Attribute("Page"), NumberStyles.None, CultureInfo.InvariantCulture, out int number) || number < 1 || number > pages.Count)
                        throw new InvalidDataException("Unresolved story-fragment page number.");
                    string? fragmentName = (string?)reference.Attribute("FragmentName");
                    var matches = pages[number - 1].Fragments.Where(f => f.StoryName == name && (fragmentName == null || f.FragmentName == fragmentName)).ToArray();
                    if (matches.Length == 0) Loss("Unresolved story fragment: " + name + " on page " + number.ToString(CultureInfo.InvariantCulture));
                    foreach (var fragment in matches) {
                        if (fragment.Type != XpsStoryFragmentType.Content) Loss("A declared body story references non-content fragment: " + name);
                        if (!used.Add(fragment)) {
                            Loss("Story fragment referenced more than once: " + name);
                            budget.Text(fragment.Blocks.Sum(b => b.TextLength));
                        }
                        fragments.Add(fragment);
                    }
                }
                if (fragments.Count > 0) stories.Add(new XpsLogicalStory(name, XpsStoryFragmentType.Content,
                    fragments.AsReadOnly(), XpsStructureMerger.Merge(fragments, budget, Loss)));
            }
        }
        foreach (var page in pages) foreach (var fragment in page.Fragments) {
            budget.Charge();
            if (used.Contains(fragment)) continue;
            if (fragment.StoryName != null) Loss("Story fragment has no declared story address: " + fragment.StoryName);
            stories.Add(new XpsLogicalStory(fragment.StoryName, fragment.Type, new[] { fragment }, fragment.Blocks));
        }
        return new XpsLogicalStructure(pages.AsReadOnly(), stories.AsReadOnly(), diagnostics.AsReadOnly());
    }
    private static void ValidateStructureLiteralText(XElement element) {
        if (element.Nodes().OfType<XText>().Any(t => !string.IsNullOrWhiteSpace(t.Value)))
            throw new InvalidDataException("Native document structure cannot contain literal text.");
    }
}

internal static class XpsStructureMerger {
    internal static IReadOnlyList<XpsStructureNode> Merge(IReadOnlyList<XpsStoryFragment> fragments,
        XpsStoryFragmentsReader.Budget budget, Action<string> loss) {
        var blocks = new List<XpsStructureNode>(); XpsStoryFragment? previous = null;
        foreach (var fragment in fragments) {
            budget.Charge();
            if (previous == null || previous.BreakAfter || fragment.BreakBefore) blocks.AddRange(fragment.Blocks);
            else blocks = MergeChildren(blocks, fragment.Blocks, budget, loss);
            previous = fragment;
        }
        return blocks.AsReadOnly();
    }
    private static List<XpsStructureNode> MergeChildren(IReadOnlyList<XpsStructureNode> leading,
        IReadOnlyList<XpsStructureNode> trailing, XpsStoryFragmentsReader.Budget budget, Action<string> loss) {
        budget.Charge(leading.Count + trailing.Count);
        var output = new List<XpsStructureNode>(leading);
        var merged = leading.Count > 0 && trailing.Count > 0 ? MergeNode(leading[leading.Count - 1], trailing[0], budget, loss) : null;
        if (merged == null) output.AddRange(trailing);
        else { output[output.Count - 1] = merged; output.AddRange(trailing.Skip(1)); }
        return output;
    }
    private static XpsStructureNode? MergeNode(XpsStructureNode leading, XpsStructureNode trailing,
        XpsStoryFragmentsReader.Budget budget, Action<string> loss) {
        budget.Charge();
        if (leading.Kind != trailing.Kind || leading.Kind == XpsStructureKind.NamedElement || leading.Kind == XpsStructureKind.Unknown) return null;
        bool emptyPlaceholder = trailing.Kind == XpsStructureKind.TableCell && trailing.Children.Count == 0 && trailing.RowSpan == 1 && trailing.ColumnSpan == 1;
        if (!emptyPlaceholder && (leading.RowSpan != trailing.RowSpan || leading.ColumnSpan != trailing.ColumnSpan)) {
            loss("Incompatible continued table-cell spans."); return null;
        }
        List<XpsStructureNode> children;
        if (leading.Kind == XpsStructureKind.TableRow) {
            // All cells of the continued boundary row are aligned, including empty
            // placeholders. Other row-group rows remain in their native order.
            if (leading.Children.Count != trailing.Children.Count || leading.Children.Any(c => c.Kind != XpsStructureKind.TableCell) || trailing.Children.Any(c => c.Kind != XpsStructureKind.TableCell)) {
                loss("Continued table rows require matching cell placeholders."); return null;
            }
            children = new List<XpsStructureNode>();
            for (int i = 0; i < leading.Children.Count; i++) {
                var cell = MergeNode(leading.Children[i], trailing.Children[i], budget, loss);
                if (cell == null) return null;
                children.Add(cell);
            }
        } else children = MergeChildren(leading.Children, trailing.Children, budget, loss);
        return new XpsStructureNode(leading.Kind, children.AsReadOnly(), marker: leading.Marker ?? trailing.Marker,
            rowSpan: leading.RowSpan, columnSpan: leading.ColumnSpan);
    }
}
