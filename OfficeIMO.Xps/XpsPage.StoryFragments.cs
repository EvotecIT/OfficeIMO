namespace OfficeIMO.Xps;

public sealed partial class XpsPage {
    /// <summary>Returns detached relationship-owned StoryFragments markup, or null when the page has none.</summary>
    public XElement? GetStoryFragmentsMarkup(CancellationToken cancellationToken = default) =>
        Document.ReadStoryFragmentsPart(this, cancellationToken)?.Markup;

    /// <summary>Reads native paragraph, list, table and figure structure without inferring semantics from placement.</summary>
    /// <remarks>For a repeated page, this page-level snapshot uses its first sequence occurrence. ReadLogicalStructure returns every occurrence.</remarks>
    public XpsPageStructure ReadContentStructure(CancellationToken cancellationToken = default) =>
        Document.ReadPageContent(this, Document.Pages.ToList().IndexOf(this), new XpsStoryFragmentsReader.Budget(cancellationToken));

    /// <summary>Creates or replaces the page's native StoryFragments part, preserving unrelated relationships and package parts.</summary>
    public void ReplaceStoryFragmentsMarkup(XElement markup) {
        if (markup == null) throw new ArgumentNullException(nameof(markup));
        Document.CommitPageContent(this, GetMarkup(), markup);
    }

    /// <summary>Atomically replaces page and StoryFragments markup, allowing named content and references to be edited together.</summary>
    public void ReplaceMarkup(XElement markup, XElement storyFragments) {
        if (markup == null) throw new ArgumentNullException(nameof(markup));
        if (storyFragments == null) throw new ArgumentNullException(nameof(storyFragments));
        if (markup.Name != _markup.Name) throw new ArgumentException("Expected FixedPage in the document dialect.", nameof(markup));
        ValidatePageDimension(XpsPackage.Number((string?)markup.Attribute("Width")));
        ValidatePageDimension(XpsPackage.Number((string?)markup.Attribute("Height")));
        Document.CommitPageContent(this, Document.ValidatePageMarkup(markup), storyFragments);
    }
}

public sealed partial class XpsDocument {
    internal (string Part, XElement Markup)? ReadStoryFragmentsPart(XpsPage page, CancellationToken token) {
        string? part = ResolveStoryFragmentsPart(page, token);
        if (part == null) return null;
        var markup = ReadXml(part, token);
        if (markup.Name != StructureNamespace + "StoryFragments") throw new InvalidDataException("Invalid StoryFragments root or dialect.");
        return (part, markup);
    }
    private string? ResolveStoryFragmentsPart(XpsPage page, CancellationToken token) {
        string relName = RelationshipPartName(page.PartName);
        if (!_parts.ContainsKey(relName)) return null;
        var relationships = ReadXml(relName, token);
        if (relationships.Name != XpsPackage.Relationships + "Relationships") throw new InvalidDataException("Invalid fixed-page relationships.");
        var candidates = relationships.Elements(XpsPackage.Relationships + "Relationship")
            .Where(r => (string?)r.Attribute("Type") == XpsPackage.Namespace(Format) + "/storyfragments").ToArray();
        if (candidates.Length > 1) throw new InvalidDataException("A fixed page can reference only one StoryFragments part.");
        if (candidates.Length == 0) return null;
        var relationship = candidates[0];
        if (((string?)relationship.Attribute("TargetMode") ?? "Internal") != "Internal") throw new InvalidDataException("StoryFragments must be package-local.");
        string part = XpsPackage.Resolve(page.PartName, (string?)relationship.Attribute("Target") ?? "");
        if (ContentType(part) != XpsPackage.Type("storyfragments")) throw new InvalidDataException("Invalid StoryFragments content type.");
        return part;
    }
    internal XpsPageStructure ReadPageContent(XpsPage page, int pageIndex, XpsStoryFragmentsReader.Budget budget) {
        budget.Charge();
        string? part = ResolveStoryFragmentsPart(page, budget.Token);
        if (part == null) return new XpsPageStructure(false, Array.Empty<XpsStoryFragment>(), Array.Empty<string>());
        if (!budget.StoryParts.TryGetValue(part, out var markup)) {
            markup = ReadXml(part, budget.Token); budget.StoryParts.Add(part, markup);
        }
        return new XpsStoryFragmentsReader(page, pageIndex, page.GetMarkup(), budget).Read(markup);
    }
    private static string RelationshipPartName(string source) {
        int slash = source.LastIndexOf('/');
        return source.Substring(0, slash + 1) + "_rels/" + source.Substring(slash + 1) + ".rels";
    }

    internal void CommitPageContent(XpsPage page, XElement pageMarkup, XElement? storyMarkup = null) {
        var old = ReadStoryFragmentsPart(page, default);
        var replacements = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase) { [page.PartName] = XpsPackage.Serialize(pageMarkup) };
        var types = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        XElement? prospectiveStory = storyMarkup != null ? ValidatePageMarkup(storyMarkup) : old?.Markup;
        if (prospectiveStory != null) {
            var budget = new XpsStoryFragmentsReader.Budget(default);
            var result = new XpsStoryFragmentsReader(page, _pages.IndexOf(page), pageMarkup, budget).Read(prospectiveStory);
            if (result.HasUnresolvedNames)
                throw new InvalidDataException("Page edits cannot leave unresolved native structure references.");
            ValidateStoryReplacement(page, result.Fragments);
            if (storyMarkup != null) {
                // A shared part has several owners: a replacement must remain valid for each.
                string part = old?.Part ?? NewPartName(page.PartName + ".StoryFragments", ".frag");
                foreach (var other in _pageCache.Values.Where(p => p != page)) {
                    budget.Charge();
                    string? related = ResolveStoryFragmentsPart(other, default);
                    if (!string.Equals(related, part, StringComparison.OrdinalIgnoreCase)) continue;
                    var otherResult = new XpsStoryFragmentsReader(other, _pages.IndexOf(other), other.GetMarkup(), budget).Read(prospectiveStory);
                    if (otherResult.HasUnresolvedNames)
                        throw new InvalidDataException("Shared StoryFragments replacement is invalid for another page.");
                    ValidateStoryReplacement(other, otherResult.Fragments);
                }
                replacements[part] = XpsPackage.Serialize(prospectiveStory); types[part] = XpsPackage.Type("storyfragments");
                if (!old.HasValue) {
                    string relName = RelationshipPartName(page.PartName);
                    XElement relationships = _parts.ContainsKey(relName) ? ReadXml(relName, default) : new XElement(XpsPackage.Relationships + "Relationships");
                    relationships.Add(new XElement(XpsPackage.Relationships + "Relationship", new XAttribute("Id", NextRelationshipId(relationships)),
                        new XAttribute("Type", XpsPackage.Namespace(Format) + "/storyfragments"), new XAttribute("Target", "/" + part)));
                    replacements[relName] = XpsPackage.Serialize(relationships);
                }
            }
        }
        _ = PrepareOutput(default, replacements, types);
        foreach (var replacement in replacements) _parts[replacement.Key] = replacement.Value;
        foreach (var type in types) _types[type.Key] = type.Value;
        page.ApplyMarkup(pageMarkup);
    }
    // Reject removal/renaming of a declared fragment instead of leaving a dangling
    // story address. Reordering or editing its content does not change the address.
    private void ValidateStoryReplacement(XpsPage page, IReadOnlyList<XpsStoryFragment> fragments) {
        var occurrences = new HashSet<int>(Enumerable.Range(0, _pages.Count).Where(i => _pages[i] == page));
        foreach (var structure in ReadDocumentStructures(activeOnly: true)) foreach (var story in structure.Markup.Elements(StructureNamespace + "Story")) {
            string? name = (string?)story.Attribute("StoryName");
            foreach (var reference in story.Elements(StructureNamespace + "StoryFragmentReference")) {
                if (!int.TryParse((string?)reference.Attribute("Page"), NumberStyles.None, CultureInfo.InvariantCulture, out int number) || !occurrences.Contains(number - 1)) continue;
                string? fragment = (string?)reference.Attribute("FragmentName");
                if (!fragments.Any(f => f.Type == XpsStoryFragmentType.Content && f.StoryName == name && (fragment == null || f.FragmentName == fragment)))
                    throw new InvalidDataException("StoryFragments edits cannot remove a declared story address.");
            }
        }
    }
}
