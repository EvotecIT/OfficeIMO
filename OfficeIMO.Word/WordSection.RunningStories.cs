using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word;

public partial class WordSection {
    private const int WStory = OfficeXmlNamespaces.WordprocessingML;
    private bool _ownedStoriesLoaded;

    internal bool HasRunningStoryReference(string element, WordHeaderFooterType kind) {
        string token = kind is WordHeaderFooterType.First ? "first" : kind is WordHeaderFooterType.Even ? "even" : "default";
        foreach (OfficeXmlNode reference in SectionNode.Children(WStory, element)) {
            if (reference.Attribute(WStory, "type") == token) return true;
        }
        return false;
    }

    private void ReadOwnedStories(WordPackageModel model) {
        if (_ownedStoriesLoaded) return;
        foreach (OfficeXmlNode reference in SectionNode.Children()) {
            if (reference.Is(WStory, "headerReference")) ReadStory(reference, header: true);
            else if (reference.Is(WStory, "footerReference")) ReadStory(reference, header: false);
        }
        _ownedStoriesLoaded = true;

        void ReadStory(OfficeXmlNode reference, bool header) {
            string? id = reference.Attribute(OfficeXmlNamespaces.Relationships, "id");
            string? part = id == null ? null : model.Package.GetRelationshipTarget(model.MainPartName, id);
            string type = reference.Attribute(WStory, "type") ?? "default";
            string contentType = RunningStoryContentType(header ? "header" : "footer");
            if (part != null && model.Package.GetContentType(part) != contentType) part = null;
            AttachOwnedStory(header, type, part);
        }
    }

    private void AttachOwnedStory(bool header, string type, string? part) {
        if (header) {
            var story = new WordHeader(_document, this, part);
            if (type == "first") _headers.First = story;
            else if (type == "even") _headers.Even = story;
            else _headers.Default = story;
        } else {
            var story = new WordFooter(_document, this, part);
            if (type == "first") _footers.First = story;
            else if (type == "even") _footers.Even = story;
            else _footers.Default = story;
        }
    }

    internal bool TryAddOwnedStory(bool header, string type) {
        if (_document._ownedModel is not { } model) return false;
        ReadOwnedStories(model);
        string relationshipName = header ? "header" : "footer";
        string referenceName = relationshipName + "Reference";
        OfficeXmlNode section = SectionNode;
        foreach (OfficeXmlNode existing in section.Children(WStory, referenceName)) {
            if (existing.Attribute(WStory, "type") == type) return true;
        }

        _ = model.Package.GetRelationships(model.MainPartName);
        int index = 1;
        string part;
        do {
            part = "/word/" + relationshipName + index.ToString(CultureInfo.InvariantCulture) + ".xml";
            index++;
        } while (model.Package.PartExists(part));

        model.Package.CreateXmlPart(part, RunningStoryContentType(relationshipName), WStory, header ? "hdr" : "ftr");
        string id = model.Package.AddRelationship(model.MainPartName,
            OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + relationshipName, part.Substring("/word/".Length),
            id: "R" + Guid.NewGuid().ToString("N").Substring(0, 16));

        section = EditSectionNode();
        OfficeXmlNode reference = section.CreateElement(WStory, referenceName);
        reference.SetAttribute(WStory, "type", type);
        reference.SetAttribute(OfficeXmlNamespaces.Relationships, "id", id);
        OfficeXmlNode lastReference = default;
        foreach (OfficeXmlNode child in section.Children()) {
            if (child.Is(WStory, "headerReference") || child.Is(WStory, "footerReference")) lastReference = child;
        }
        section.InsertBefore(reference, lastReference.Exists ? lastReference.NextSibling : section.FirstChild);
        AttachOwnedStory(header, type, part);
        return true;
    }

    private static string RunningStoryContentType(string name) =>
        "application/vnd.openxmlformats-officedocument.wordprocessingml." + name + "+xml";

    internal bool TryGetOwnedStorySettings(out OfficeXmlNode settings) {
        settings = default;
        if (_document._ownedModel is not { } model) return false;
        string? part = WordOwnedSave.RelatedParts(model, OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "settings")
            .FirstOrDefault();
        if (part == null) return true;
        OfficeXmlTree tree = model.Package.GetXml(part);
        settings = new OfficeXmlNode(tree, tree.Root);
        return true;
    }
}
