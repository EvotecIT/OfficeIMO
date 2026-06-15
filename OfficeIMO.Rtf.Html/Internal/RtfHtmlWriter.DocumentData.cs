using System.Globalization;

namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlWriter {
    private static void AppendUserPropertiesMetadata(StringBuilder builder, RtfDocument document, string newline) {
        if (document.UserProperties.Count == 0) {
            return;
        }

        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        for (int index = 0; index < document.UserProperties.Count; index++) {
            RtfUserProperty property = document.UserProperties[index];
            string prefix = "property." + index.ToString(CultureInfo.InvariantCulture);
            values[prefix + ".name"] = property.Name;
            AddNullableInt(values, prefix + ".typeCode", property.TypeCode);
            AddString(values, prefix + ".staticValue", property.StaticValue);
            AddString(values, prefix + ".linkedValue", property.LinkedValue);
        }

        builder.Append(newline);
        builder.Append("<meta name=\"officeimo-rtf-user-properties\" content=\"");
        builder.Append(EncodeAttribute(RtfHtmlMetadataCodec.Encode(values)));
        builder.Append("\">");
    }

    private static void AppendDocumentVariablesMetadata(StringBuilder builder, RtfDocument document, string newline) {
        if (document.DocumentVariables.Count == 0) {
            return;
        }

        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        for (int index = 0; index < document.DocumentVariables.Count; index++) {
            RtfDocumentVariable variable = document.DocumentVariables[index];
            string prefix = "variable." + index.ToString(CultureInfo.InvariantCulture);
            values[prefix + ".name"] = variable.Name;
            values[prefix + ".value"] = variable.Value;
        }

        builder.Append(newline);
        builder.Append("<meta name=\"officeimo-rtf-document-variables\" content=\"");
        builder.Append(EncodeAttribute(RtfHtmlMetadataCodec.Encode(values)));
        builder.Append("\">");
    }
}
