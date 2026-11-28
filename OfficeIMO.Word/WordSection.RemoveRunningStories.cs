using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word;

public partial class WordDocument {
    internal bool TryRemoveOwnedRunningStories(bool header, WordHeaderFooterType[]? types) {
        if (_ownedModel is not { } model || IsOwnedReadOnly) return false;
        const int W = OfficeXmlNamespaces.WordprocessingML;
        const int R = OfficeXmlNamespaces.Relationships;
        const int Relationships = OfficeXmlNamespaces.PackageRelationships;
        string name = header ? "header" : "footer";
        string referenceName = name + "Reference";
        string relationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + name;
        bool allTypes = types == null || types.Length == 0;
        HashSet<string>? selectedTypes = allTypes ? null : new HashSet<string>(types!.Select(WordSection.RunningStoryTypeToken), StringComparer.Ordinal);

        OfficeXmlNode body = new(model.Document, model.Body);
        OfficeXmlNode linksRoot = default;
        string linksPart = OfficeOpcPartNames.GetRelationshipsPartName(model.MainPartName);
        if (model.Package.PartExists(linksPart)) {
            OfficeXmlTree links = model.Package.GetRelationships(model.MainPartName);
            linksRoot = new OfficeXmlNode(links, links.Root);
        }
        var relationships = linksRoot.Children(Relationships, "Relationship").ToList();
        string? sourceDirectory = OfficeOpcPartNames.GetRelationshipSourceDirectory(linksPart);
        var matchingRelationships = new List<(OfficeXmlNode Node, string Id, string Part)>();
        foreach (OfficeXmlNode link in relationships) {
            string? type = link.Attribute("Type");
            if (type == null || OfficeXmlNamespaces.NormalizeRelationshipType(type) != relationshipType ||
                link.Attribute("Id") is not { } id || ResolveTarget(link, sourceDirectory) is not { } part ||
                !model.Package.PartExists(part)) continue;
            matchingRelationships.Add((link, id, part));
        }
        if (allTypes && matchingRelationships.Count == 0) return true;

        var references = new HashSet<OfficeXmlNode>();
        var selectedIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (OfficeXmlNode node in WordFieldInventory.EnumerateFieldContent(body)) {
            if (!node.Is(W, referenceName) || node.AncestorOrSelf(W, "sectPr").IsNone) continue;
            string? id = node.Attribute(R, "id");
            if (!allTypes && (id == null || node.Attribute(W, "type") is not { } type || !selectedTypes!.Contains(type))) continue;
            references.Add(node);
            if (id != null) selectedIds.Add(id);
        }

        // Reserve package relationships used by preserved markup, including inactive compatibility content.
        var keptIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (OfficeXmlNode node in WordXmlNodes.Descendants(body)) {
            if (!references.Contains(node) && node.Attribute(R, "id") is { } id) keptIds.Add(id);
        }

        var linksToRemove = new HashSet<OfficeXmlNode>();
        var partsToRemove = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var match in matchingRelationships) {
            if ((allTypes || selectedIds.Contains(match.Id)) && !keptIds.Contains(match.Id)) {
                linksToRemove.Add(match.Node);
                partsToRemove.Add(match.Part);
            }
        }
        foreach (OfficeXmlNode link in relationships) {
            if (!linksToRemove.Contains(link) && ResolveTarget(link, sourceDirectory) is { } target) partsToRemove.Remove(target);
        }

        foreach (string otherLinksPart in model.Package.PartNames) {
            if (partsToRemove.Count == 0) break;
            if (string.Equals(otherLinksPart, linksPart, StringComparison.OrdinalIgnoreCase)) continue;
            string? otherSourceDirectory = OfficeOpcPartNames.GetRelationshipSourceDirectory(otherLinksPart);
            if (otherSourceDirectory == null) continue;

            OfficeXmlTree otherLinks = model.Package.GetXml(otherLinksPart);
            foreach (OfficeXmlNode link in new OfficeXmlNode(otherLinks, otherLinks.Root).Children(Relationships, "Relationship")) {
                if (ResolveTarget(link, otherSourceDirectory) is { } target) partsToRemove.Remove(target);
            }
        }

        // Stories with dependencies use the existing package-graph cleanup.
        if (partsToRemove.Any(part => model.Package.PartExists(OfficeOpcPartNames.GetRelationshipsPartName(part)))) return false;
        if (references.Count == 0 && linksToRemove.Count == 0) return true;

        var activeSections = new Dictionary<OfficeXmlNode, WordSection>();
        foreach (WordSection section in Sections) activeSections[section.SectionNode] = section;
        foreach (OfficeXmlNode reference in references) {
            if (activeSections.TryGetValue(reference.AncestorOrSelf(W, "sectPr"), out WordSection? section)) {
                section.ForgetOwnedStory(header, reference.Attribute(W, "type") ?? "default");
            }
            reference.Remove();
        }
        foreach (OfficeXmlNode link in linksToRemove) link.Remove();
        RetireOwnedRunningStories(partsToRemove);
        foreach (string part in partsToRemove) model.Package.DeletePart(part);
        model.MarkChanged();
        return true;

        static string? ResolveTarget(OfficeXmlNode link, string? sourceDirectory) =>
            string.Equals(link.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase) ||
            link.Attribute("Target") is not { } target ? null : OfficeOpcPartNames.ResolveRelationshipTargetFrom(sourceDirectory, target);
    }
}

public partial class WordSection {
    internal static string RunningStoryTypeToken(WordHeaderFooterType type) => type switch {
        WordHeaderFooterType.Default => "default",
        WordHeaderFooterType.First => "first",
        WordHeaderFooterType.Even => "even",
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, "Unsupported running story type.")
    };

    internal void ForgetOwnedStory(bool header, string type) {
        if (header) {
            if (type == "first") _headers.First = null;
            else if (type == "even") _headers.Even = null;
            else _headers.Default = null;
        } else {
            if (type == "first") _footers.First = null;
            else if (type == "even") _footers.Even = null;
            else _footers.Default = null;
        }
    }
}
