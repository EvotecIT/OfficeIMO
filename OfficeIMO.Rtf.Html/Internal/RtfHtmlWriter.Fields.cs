namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlWriter {
    private static void AppendField(StringBuilder builder, RtfField field, RtfHtmlSaveOptions options, RtfDocument document) {
        builder.Append("<span data-officeimo-rtf-field=\"true\" data-officeimo-rtf-field-instruction=\"");
        builder.Append(EncodeAttribute(field.Instruction));
        builder.Append("\">");
        AppendInlines(builder, field.Result.Inlines, options, document);
        builder.Append("</span>");
    }
}
