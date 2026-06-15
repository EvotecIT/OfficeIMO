using System.Globalization;

namespace OfficeIMO.Html.Rtf;

internal static partial class RtfHtmlWriter {
    private static void AppendListAttributes(StringBuilder builder, RtfParagraph paragraph) {
        if (paragraph.ListKind == RtfListKind.None) {
            return;
        }

        if (paragraph.ListId.HasValue) {
            AppendListIntegerAttribute(builder, "data-officeimo-rtf-list-id", paragraph.ListId.Value);
        }

        if (paragraph.ListDefinitionId.HasValue) {
            AppendListIntegerAttribute(builder, "data-officeimo-rtf-list-definition-id", paragraph.ListDefinitionId.Value);
        }

        if (paragraph.ListLevel.HasValue) {
            AppendListIntegerAttribute(builder, "data-officeimo-rtf-list-level", paragraph.ListLevel.Value);
        }

        string? listText = paragraph.ListText?.ToPlainText();
        if (listText != null) {
            builder.Append(" data-officeimo-rtf-list-text=\"");
            builder.Append(EncodeAttribute(listText));
            builder.Append('"');
        }
    }

    private static void AppendListIntegerAttribute(StringBuilder builder, string name, int value) {
        builder.Append(' ');
        builder.Append(name);
        builder.Append("=\"");
        builder.Append(value.ToString(CultureInfo.InvariantCulture));
        builder.Append('"');
    }
}
