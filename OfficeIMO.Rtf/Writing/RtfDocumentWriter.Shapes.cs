namespace OfficeIMO.Rtf.Writing;

internal static partial class RtfDocumentWriter {
    private static void WriteShape(StringBuilder builder, RtfShape shape, RtfWriteContext context) {
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
            builder.Append(EscapeText(property.Name, context.UnicodeSkipCount));
            builder.Append(@"}{\sv ");
            builder.Append(EscapeText(property.Value, context.UnicodeSkipCount));
            builder.Append("}}");
        }

        if (shape.TextBoxParagraphs.Count > 0) {
            builder.Append(@"{\shptxt");
            foreach (RtfParagraph paragraph in shape.TextBoxParagraphs) {
                WriteParagraph(builder, paragraph, context);
            }

            builder.Append('}');
        }

        builder.Append("}}");
    }
}
