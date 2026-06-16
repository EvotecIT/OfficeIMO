namespace OfficeIMO.Rtf.Writing;

internal static partial class RtfDocumentWriter {
    private static void WriteShape(StringBuilder builder, RtfShape shape, int? defaultLanguageId, int unicodeSkipCount) {
        builder.Append(@"{\shp{\*\shpinst");
        foreach (RtfShapeInstruction instruction in shape.Instructions) {
            builder.Append('\\');
            builder.Append(instruction.Name);
            if (instruction.HasParameter && instruction.Parameter.HasValue) {
                builder.Append(instruction.Parameter.Value.ToString(CultureInfo.InvariantCulture));
            }
        }

        foreach (RtfShapeProperty property in shape.Properties) {
            builder.Append(@"{\sp{\sn ");
            builder.Append(EscapeText(property.Name, unicodeSkipCount));
            builder.Append(@"}{\sv ");
            builder.Append(EscapeText(property.Value, unicodeSkipCount));
            builder.Append("}}");
        }

        if (shape.TextBoxParagraphs.Count > 0) {
            builder.Append(@"{\shptxt");
            foreach (RtfParagraph paragraph in shape.TextBoxParagraphs) {
                WriteParagraph(builder, paragraph, defaultLanguageId, unicodeSkipCount);
            }

            builder.Append('}');
        }

        builder.Append("}}");
    }
}
