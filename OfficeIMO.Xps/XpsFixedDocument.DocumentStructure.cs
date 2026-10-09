namespace OfficeIMO.Xps;

public sealed partial class XpsFixedDocument {
    /// <summary>Returns detached relationship-owned DocumentStructure markup, or null when none is attached.</summary>
    public XElement? GetDocumentStructureMarkup(CancellationToken cancellationToken = default) =>
        Owner.ReadOwnedDocumentStructure(this, cancellationToken)?.Markup;

    /// <summary>Creates or replaces native outlines and story addresses, retaining unknown markup and unrelated relationships.</summary>
    /// <remarks>Story addresses must resolve to page fragments. Numeric page addresses follow the adopted payload-global interpretation.</remarks>
    public void ReplaceDocumentStructureMarkup(XElement markup) {
        if (markup == null) throw new ArgumentNullException(nameof(markup));
        Owner.CommitDocumentStructure(this, markup);
    }
}

public sealed partial class XpsDocument {
    internal (string Part, XElement Markup)? ReadOwnedDocumentStructure(XpsFixedDocument document, CancellationToken token) {
        string relName = RelationshipPartName(document.PartName);
        if (!_parts.ContainsKey(relName)) return null;
        var relationships = ReadXml(relName, token);
        if (relationships.Name != XpsPackage.Relationships + "Relationships") throw new InvalidDataException("Invalid fixed-document relationships.");
        var candidates = relationships.Elements(XpsPackage.Relationships + "Relationship")
            .Where(r => (string?)r.Attribute("Type") == XpsPackage.Namespace(Format) + "/documentstructure").ToArray();
        if (candidates.Length > 1) throw new InvalidDataException("A fixed document can reference only one DocumentStructure part.");
        if (candidates.Length == 0) return null;
        if (((string?)candidates[0].Attribute("TargetMode") ?? "Internal") != "Internal") throw new InvalidDataException("DocumentStructure must be package-local.");
        string part = XpsPackage.Resolve(document.PartName, (string?)candidates[0].Attribute("Target") ?? "");
        if (ContentType(part) != XpsPackage.Type("documentstructure")) throw new InvalidDataException("Invalid DocumentStructure content type.");
        var markup = ReadXml(part, token);
        if (markup.Name != StructureNamespace + "DocumentStructure") throw new InvalidDataException("Invalid DocumentStructure root or dialect.");
        return (part, markup);
    }
    internal void CommitDocumentStructure(XpsFixedDocument document, XElement markup) {
        var copy = ValidatePageMarkup(markup);
        if (copy.Name != StructureNamespace + "DocumentStructure") throw new InvalidDataException("Invalid DocumentStructure root or dialect.");
        ValidateStructureLiteralText(copy);
        var budget = new XpsStoryFragmentsReader.Budget(default);
        var pageCache = new Dictionary<int, XpsPageStructure>(); var storyNames = new HashSet<string>(StringComparer.Ordinal);
        foreach (var story in copy.Elements(StructureNamespace + "Story")) {
            ValidateStructureLiteralText(story);
            string name = (string?)story.Attribute("StoryName") ?? throw new InvalidDataException("Missing story name.");
            if (!storyNames.Add(name)) throw new InvalidDataException("Duplicate story name in DocumentStructure.");
            var references = story.Elements(StructureNamespace + "StoryFragmentReference").ToArray();
            if (references.Length == 0) throw new InvalidDataException("A declared story requires fragment references.");
            foreach (var reference in references) {
                budget.Charge();
                ValidateStructureLiteralText(reference);
                if (reference.HasElements) throw new InvalidDataException("StoryFragmentReference cannot contain elements.");
                if (!int.TryParse((string?)reference.Attribute("Page"), NumberStyles.None, CultureInfo.InvariantCulture, out int number) || number < 1 || number > _pages.Count)
                    throw new InvalidDataException("Unresolved story-fragment page number.");
                int index = number - 1;
                if (!pageCache.TryGetValue(index, out var page)) { page = ReadPageContent(_pages[index], index, budget); pageCache.Add(index, page); }
                string? fragment = (string?)reference.Attribute("FragmentName");
                if (!page.Fragments.Any(f => f.Type == XpsStoryFragmentType.Content && f.StoryName == name && (fragment == null || f.FragmentName == fragment)))
                    throw new InvalidDataException("Unresolved declared story address.");
            }
        }
        foreach (var outline in copy.Elements(StructureNamespace + "DocumentStructure.Outline")
            .Elements(StructureNamespace + "DocumentOutline").Elements(StructureNamespace + "OutlineEntry")) {
            budget.Charge();
            if (outline.Attribute("Description") == null || outline.Attribute("OutlineTarget") == null ||
                !int.TryParse((string?)outline.Attribute("OutlineLevel") ?? "1", NumberStyles.Integer, CultureInfo.InvariantCulture, out int level) || level < 1)
                throw new InvalidDataException("Invalid native outline entry.");
        }
        var old = ReadOwnedDocumentStructure(document, default);
        string part = old?.Part ?? NewPartName(document.PartName + ".DocumentStructure", ".struct");
        var replacements = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase) { [part] = XpsPackage.Serialize(copy) };
        if (!old.HasValue) {
            string relName = RelationshipPartName(document.PartName);
            var relationships = _parts.ContainsKey(relName) ? ReadXml(relName, default) : new XElement(XpsPackage.Relationships + "Relationships");
            relationships.Add(new XElement(XpsPackage.Relationships + "Relationship", new XAttribute("Id", NextRelationshipId(relationships)),
                new XAttribute("Type", XpsPackage.Namespace(Format) + "/documentstructure"), new XAttribute("Target", "/" + part)));
            replacements[relName] = XpsPackage.Serialize(relationships);
        }
        var types = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { [part] = XpsPackage.Type("documentstructure") };
        _ = PrepareOutput(default, replacements, types);
        foreach (var replacement in replacements) _parts[replacement.Key] = replacement.Value;
        foreach (var type in types) _types[type.Key] = type.Value;
    }
}
