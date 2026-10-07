using System.Globalization;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    /// <summary>
    /// Appends a creator and its ordered MARC roles, sorting name and language in one package edit.
    /// Existing creators and refinements remain intact. Requires EPUB 3 and an unused package identifier.
    /// </summary>
    public void AddCreator(string id, EpubContributorMetadata creator) => AddContributorRecord("creator", id, creator);

    /// <summary>
    /// Appends a contributor and its ordered MARC roles, sorting name and language in one package edit.
    /// Existing contributors and refinements remain intact. Requires EPUB 3 and an unused package identifier.
    /// </summary>
    public void AddContributor(string id, EpubContributorMetadata contributor) => AddContributorRecord("contributor", id, contributor);

    /// <summary>
    /// Appends a named series or set with its kind, hierarchical position and sorting name atomically.
    /// Existing collections remain intact. Requires EPUB 3 and an unused package identifier.
    /// </summary>
    public void AddCollection(string id, EpubCollectionMetadata collection) {
        if (collection == null) throw new ArgumentNullException(nameof(collection));
        RequirePublishingMetadataId(id);
        ValidatePublishingName(collection.Name, collection.FileAs, collection.Language);
        string kind = collection.Kind switch {
            EpubCollectionKind.Series => "series",
            EpubCollectionKind.Set => "set",
            _ => throw new ArgumentOutOfRangeException(nameof(collection.Kind))
        };
        if (collection.Position == null) throw new ArgumentNullException(nameof(collection.Position));
        string position = string.Join(".", collection.Position.Select(value => value.ToString(CultureInfo.InvariantCulture)));
        var member = new XElement(Opf + "meta", new XAttribute("id", id),
            new XAttribute("property", "belongs-to-collection"), collection.Name);
        member.SetAttributeValue(XNamespace.Xml + "lang", collection.Language);
        var elements = new List<XElement> { member, PublishingRefinement(id, "collection-type", kind) };
        if (position.Length != 0) elements.Add(PublishingRefinement(id, "group-position", position));
        if (collection.FileAs != null) elements.Add(PublishingRefinement(id, "file-as", collection.FileAs));
        AppendPublishingMetadata(elements);
    }

    private void AddContributorRecord(string elementName, string id, EpubContributorMetadata contributor) {
        if (contributor == null) throw new ArgumentNullException(nameof(contributor));
        RequirePublishingMetadataId(id);
        ValidatePublishingName(contributor.Name, contributor.FileAs, contributor.Language);
        if (contributor.MarcRoles == null) throw new ArgumentNullException(nameof(contributor.MarcRoles));
        string[] roles = contributor.MarcRoles.Select(role => {
            if (role == null || role.Length != 3 || role.Any(value => value < 'a' || value > 'z'))
                throw new ArgumentException("MARC relator codes require three lowercase ASCII letters.", nameof(contributor.MarcRoles));
            return role;
        }).Distinct(StringComparer.Ordinal).ToArray();
        if (roles.Length != 0 && EpubVocabulary.Expand(Root, "marc:relators") != "http://id.loc.gov/vocabulary/relators")
            throw new InvalidOperationException("The marc prefix does not identify the MARC vocabulary.");
        var contributorElement = new XElement(Dc + elementName, new XAttribute("id", id), contributor.Name);
        contributorElement.SetAttributeValue(XNamespace.Xml + "lang", contributor.Language);
        var elements = new List<XElement> { contributorElement };
        foreach (string role in roles) {
            XElement refinement = PublishingRefinement(id, "role", role);
            refinement.SetAttributeValue("scheme", "marc:relators");
            elements.Add(refinement);
        }
        if (contributor.FileAs != null) elements.Add(PublishingRefinement(id, "file-as", contributor.FileAs));
        AppendPublishingMetadata(elements);
    }

    private void RequirePublishingMetadataId(string id) {
        if (PackageVersion != "3.0") throw new NotSupportedException("Refined publishing metadata requires EPUB 3.");
        VerifyAvailableId(id);
    }

    private static void ValidatePublishingName(string name, string? fileAs, string? language) {
        RequireText(name, nameof(name));
        XmlConvert.VerifyXmlChars(name);
        if (fileAs != null) { RequireText(fileAs, nameof(fileAs)); XmlConvert.VerifyXmlChars(fileAs); }
        if (language != null) EpubLanguageTag.Require(language, nameof(language));
    }

    private static XElement PublishingRefinement(string id, string property, string value) =>
        new XElement(Opf + "meta", new XAttribute("refines", "#" + id), new XAttribute("property", property), value);

    private void AppendPublishingMetadata(IEnumerable<XElement> elements) =>
        EditPackageElement(RequireSection("metadata"), proposed => proposed.Add(elements.Select(element => new XElement(element))));
}
