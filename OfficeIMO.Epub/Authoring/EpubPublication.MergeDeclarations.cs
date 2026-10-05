namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private void VerifyMergeDeclarations(EpubManifestItem first, EpubManifestItem second, XElement firstPosition, XElement secondPosition) {
        XElement declaration = RequireSection("manifest").Elements(Opf + "item").Single(item => (string?)item.Attribute("id") == second.Id);
        if (declaration.Elements().Any() || declaration.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration &&
            attribute.Name != "id" && attribute.Name != "href" && attribute.Name != "media-type" && attribute.Name != "properties"))
            throw new NotSupportedException("The second resource has additional declarations that require an explicit merge policy.");
        var firstCopy = new XElement(firstPosition); var secondCopy = new XElement(secondPosition);
        firstCopy.Attribute("id")?.Remove(); firstCopy.Attribute("idref")?.Remove();
        secondCopy.Attribute("id")?.Remove(); secondCopy.Attribute("idref")?.Remove();
        if (!XNode.DeepEquals(firstCopy, secondCopy)) throw new NotSupportedException("Chapter reading-position attributes conflict.");
        var removedIds = new[] { second.Id, (string?)secondPosition.Attribute("id") }.Where(id => id != null).ToArray();
        if (Root.Descendants().Attributes("refines").Any(attribute => removedIds.Any(id => ReferencesPackageId(attribute.Value, id!))) ||
            Manifest.Any(item => item.FallbackStyleId == second.Id || item.MediaOverlayId == second.Id) ||
            RequireSection("metadata").Elements(Opf + "meta").Any(meta => (string?)meta.Attribute("name") == "cover" && (string?)meta.Attribute("content") == second.Id) ||
            Root.Element(Opf + "bindings")?.Elements(Opf + "mediaType").Any(item => (string?)item.Attribute("handler") == second.Id) == true)
            throw new NotSupportedException("Package refinements or structural references require an explicit merge policy.");
        string[] inferred = { "svg", "mathml", "remote-resources", "switch" };
        if (!Tokens(first.Properties).Except(inferred, StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal).SequenceEqual(
            Tokens(second.Properties).Except(inferred, StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal), StringComparer.Ordinal))
            throw new NotSupportedException("Chapter manifest properties conflict.");
    }
}
