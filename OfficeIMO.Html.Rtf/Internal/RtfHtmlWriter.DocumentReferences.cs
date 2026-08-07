using System.Globalization;

namespace OfficeIMO.Html;

internal static partial class RtfHtmlWriter {
    private static void AppendFileReferencesMetadata(StringBuilder builder, RtfDocument document, string newline) {
        if (document.FileReferences.Count == 0) {
            return;
        }

        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        for (int index = 0; index < document.FileReferences.Count; index++) {
            RtfFileReference file = document.FileReferences[index];
            string prefix = "file." + index.ToString(CultureInfo.InvariantCulture);
            AddInt(values, prefix + ".id", file.Id);
            values[prefix + ".path"] = file.Path;
            AddNullableInt(values, prefix + ".relativePathStart", file.RelativePathStart);
            AddNullableInt(values, prefix + ".operatingSystemNumber", file.OperatingSystemNumber);
            AddEnum(values, prefix + ".sources", (RtfFileSource?)file.Sources);
        }

        builder.Append(newline);
        builder.Append("<meta name=\"officeimo-rtf-file-references\" content=\"");
        builder.Append(EncodeAttribute(RtfHtmlMetadataCodec.Encode(values)));
        builder.Append("\">");
    }

    private static void AppendXmlNamespacesMetadata(StringBuilder builder, RtfDocument document, string newline) {
        if (document.XmlNamespaces.Count == 0) {
            return;
        }

        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        for (int index = 0; index < document.XmlNamespaces.Count; index++) {
            RtfXmlNamespace xmlNamespace = document.XmlNamespaces[index];
            string prefix = "namespace." + index.ToString(CultureInfo.InvariantCulture);
            AddInt(values, prefix + ".id", xmlNamespace.Id);
            values[prefix + ".uri"] = xmlNamespace.Uri;
        }

        builder.Append(newline);
        builder.Append("<meta name=\"officeimo-rtf-xml-namespaces\" content=\"");
        builder.Append(EncodeAttribute(RtfHtmlMetadataCodec.Encode(values)));
        builder.Append("\">");
    }
}
