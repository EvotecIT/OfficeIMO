using System.Globalization;

namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlWriter {
    private static void AppendFontTableMetadata(StringBuilder builder, RtfDocument document, string newline) {
        if (document.Fonts.Count == 0) {
            return;
        }

        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        for (int index = 0; index < document.Fonts.Count; index++) {
            RtfFont font = document.Fonts[index];
            string prefix = "font." + index.ToString(CultureInfo.InvariantCulture);
            AddInt(values, prefix + ".id", font.Id);
            values[prefix + ".name"] = font.Name;
            AddEnum(values, prefix + ".family", font.Family);
            AddNullableInt(values, prefix + ".charset", font.Charset);
            AddNullableInt(values, prefix + ".pitch", font.Pitch);
            AddNullableInt(values, prefix + ".codePage", font.CodePage);
            AddNullableInt(values, prefix + ".bias", font.Bias);
            AddString(values, prefix + ".alternateName", font.AlternateName);
            AddString(values, prefix + ".panose", font.Panose);
            AddString(values, prefix + ".nonTaggedName", font.NonTaggedName);
            AddFontEmbedding(values, prefix + ".embedding", font.Embedding);
        }

        builder.Append(newline);
        builder.Append("<meta name=\"officeimo-rtf-fonts\" content=\"");
        builder.Append(EncodeAttribute(RtfHtmlMetadataCodec.Encode(values)));
        builder.Append("\">");
    }

    private static void AddFontEmbedding(Dictionary<string, string> values, string prefix, RtfFontEmbedding? embedding) {
        if (embedding == null) {
            return;
        }

        AddEnum(values, prefix + ".type", (RtfEmbeddedFontType?)embedding.Type);
        AddString(values, prefix + ".fileName", embedding.FileName);
        AddNullableInt(values, prefix + ".fileCodePage", embedding.FileCodePage);
        AddString(values, prefix + ".data", EncodeBytes(embedding.Data));
    }

    private static void AddString(Dictionary<string, string> values, string key, string? value) {
        if (!string.IsNullOrEmpty(value)) {
            values[key] = value!;
        }
    }
}
