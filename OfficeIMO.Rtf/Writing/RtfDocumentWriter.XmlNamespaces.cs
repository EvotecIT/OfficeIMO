namespace OfficeIMO.Rtf.Writing;

internal static partial class RtfDocumentWriter {
    private static void WriteXmlNamespaceTable(StringBuilder builder, RtfDocument document, int unicodeSkipCount) {
        if (document.XmlNamespaces.Count == 0) return;

        builder.Append(@"{\*\xmlnstbl");
        foreach (RtfXmlNamespace xmlNamespace in document.XmlNamespaces.OrderBy(ns => ns.Id)) {
            builder.Append(@"{\xmlns");
            builder.Append(xmlNamespace.Id.ToString(CultureInfo.InvariantCulture));
            builder.Append(' ');
            builder.Append(EscapeText(xmlNamespace.Uri, unicodeSkipCount));
            builder.Append(";}");
        }

        builder.Append('}');
    }
}
