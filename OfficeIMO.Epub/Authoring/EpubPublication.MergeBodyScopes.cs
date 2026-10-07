namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private static XElement PrepareMergeBodyScope(XElement body) {
        var scope = new XElement(Html + "div");
        foreach (XAttribute attribute in body.Attributes().Where(IsMergeBodyScopeAttribute).ToArray()) {
            attribute.Remove();
            scope.Add(attribute);
        }
        foreach (XNode node in body.Nodes().ToArray()) {
            node.Remove();
            scope.Add(node);
        }
        body.Add(scope);
        return scope;
    }

    private static bool IsMergeBodyScopeAttribute(XAttribute attribute) =>
        attribute.Name == XNamespace.Xml + "id" || IsMergeLanguageAttribute(attribute) ||
        attribute.Name.NamespaceName.Length == 0 && (attribute.Name.LocalName == "id" || attribute.Name.LocalName == "class" ||
            attribute.Name.LocalName == "style" || attribute.Name.LocalName == "title" ||
            attribute.Name.LocalName.StartsWith("data-", StringComparison.Ordinal));
}
