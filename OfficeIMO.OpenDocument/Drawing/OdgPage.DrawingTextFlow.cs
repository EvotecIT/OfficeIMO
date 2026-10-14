namespace OfficeIMO.OpenDocument;

public sealed partial class OdgPage {
    // Index declarations once per page projection, including sources on other pages
    // or masters. Unresolved/ambiguous targets remain unsupported; no flow is inferred.
    private static HashSet<XElement> FindTextFlowParticipants(OdgDocument document, CancellationToken cancellationToken) {
        var frames = new List<XElement>();
        var participants = new HashSet<XElement>();
        var targets = new HashSet<string>(StringComparer.Ordinal);
        foreach (string part in new[] { "content.xml", "styles.xml" }) {
            if (!document.Package.ContainsEntry(part)) continue;
            foreach (XElement element in document.GetXml(part).Descendants()) {
                cancellationToken.ThrowIfCancellationRequested();
                if (element.Name != OdfNamespaces.Draw + "frame") continue;
                frames.Add(element);
                XAttribute? next = element.Element(OdfNamespaces.Draw + "text-box")?
                    .Attribute(OdfNamespaces.Draw + "chain-next-name");
                if (next == null || next.Value.Length == 0) continue;
                participants.Add(element);
                targets.Add(next.Value);
            }
        }
        foreach (XElement frame in frames) {
            cancellationToken.ThrowIfCancellationRequested();
            if (targets.Contains((string?)frame.Attribute(OdfNamespaces.Draw + "name") ?? string.Empty))
                participants.Add(frame);
        }
        return participants;
    }
}
