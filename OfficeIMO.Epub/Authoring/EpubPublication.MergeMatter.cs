namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private static XElement? PrepareMergeMatterScope(XElement body) {
        string[] tokens = Tokens((string?)body.Attribute(Ops + "type")).ToArray();
        string[] partitions = tokens.Where(IsMergeMatterToken).Distinct(StringComparer.Ordinal).ToArray();
        if (partitions.Length > 1)
            throw new NotSupportedException("Resolve conflicting document matter partitions before merging.");
        if (partitions.Length == 0) return null;
        string remaining = string.Join(" ", tokens.Where(token => !IsMergeMatterToken(token)));
        body.SetAttributeValue(Ops + "type", remaining.Length == 0 ? null : remaining);
        var section = new XElement(Html + "section", new XAttribute(Ops + "type", partitions[0]));
        foreach (XNode node in body.Nodes().ToArray()) {
            node.Remove();
            section.Add(node);
        }
        body.Add(section);
        return section;
    }

    private static bool IsMergeMatterToken(string token) =>
        token == "frontmatter" || token == "bodymatter" || token == "backmatter";
}
