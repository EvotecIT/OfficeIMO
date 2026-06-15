using System.Globalization;

namespace OfficeIMO.Html.Rtf;

internal static partial class RtfHtmlWriter {
    private static void AppendRevisionTablesMetadata(StringBuilder builder, RtfDocument document, string newline) {
        if (document.RevisionAuthors.Count == 0 &&
            !document.RevisionRootSaveId.HasValue &&
            document.RevisionSaveIds.Count == 0) {
            return;
        }

        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        for (int index = 0; index < document.RevisionAuthors.Count; index++) {
            string prefix = "author." + index.ToString(CultureInfo.InvariantCulture);
            values[prefix + ".name"] = document.RevisionAuthors[index].Name ?? string.Empty;
        }

        AddNullableInt(values, "rsid.root", document.RevisionRootSaveId);
        for (int index = 0; index < document.RevisionSaveIds.Count; index++) {
            string key = "rsid." + index.ToString(CultureInfo.InvariantCulture);
            AddInt(values, key, document.RevisionSaveIds[index]);
        }

        builder.Append(newline);
        builder.Append("<meta name=\"officeimo-rtf-revision-tables\" content=\"");
        builder.Append(EncodeAttribute(RtfHtmlMetadataCodec.Encode(values)));
        builder.Append("\">");
    }
}
