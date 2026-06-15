using System.Globalization;

namespace OfficeIMO.Html.Rtf;

internal static partial class RtfHtmlWriter {
    private static void AppendObject(StringBuilder builder, RtfObject rtfObject, RtfHtmlSaveOptions options, RtfDocument document, bool blockTag) {
        string tag = blockTag ? "div" : "span";
        builder.Append('<');
        builder.Append(tag);
        builder.Append(" data-officeimo-rtf-object=\"");
        builder.Append(FormatObjectKind(rtfObject.Kind));
        builder.Append('"');
        AppendMetadataAttribute(builder, "data-officeimo-rtf-object-class", rtfObject.ClassName);
        AppendMetadataAttribute(builder, "data-officeimo-rtf-object-name", rtfObject.Name);
        AppendMetadataAttribute(builder, "data-officeimo-rtf-object-data", EncodeBytes(rtfObject.Data));
        AppendMetadataAttribute(builder, "data-officeimo-rtf-object-width", rtfObject.Width);
        AppendMetadataAttribute(builder, "data-officeimo-rtf-object-height", rtfObject.Height);
        AppendMetadataAttribute(builder, "data-officeimo-rtf-object-scale-x", rtfObject.ScaleX);
        AppendMetadataAttribute(builder, "data-officeimo-rtf-object-scale-y", rtfObject.ScaleY);
        AppendMetadataAttribute(builder, "data-officeimo-rtf-object-result", EncodeParagraphContent(rtfObject.Result, options, document));
        AppendImageMetadata(builder, "data-officeimo-rtf-object-result-image", rtfObject.ResultImage);
        AppendMetadataAttribute(builder, "title", rtfObject.Result.ToPlainText());
        builder.Append("></");
        builder.Append(tag);
        builder.Append('>');
    }

    private static void AppendShape(StringBuilder builder, RtfShape shape, RtfHtmlSaveOptions options, RtfDocument document, bool blockTag) {
        string tag = blockTag ? "div" : "span";
        builder.Append('<');
        builder.Append(tag);
        builder.Append(" data-officeimo-rtf-shape=\"true\"");
        AppendMetadataAttribute(builder, "data-officeimo-rtf-shape-instructions", EncodeShapeInstructions(shape.Instructions));
        AppendMetadataAttribute(builder, "data-officeimo-rtf-shape-properties", EncodeShapeProperties(shape.Properties));
        AppendMetadataAttribute(builder, "data-officeimo-rtf-shape-text", EncodeShapeText(shape, options, document));
        AppendMetadataAttribute(builder, "title", shape.ToPlainText());
        builder.Append("></");
        builder.Append(tag);
        builder.Append('>');
    }

    private static string? EncodeParagraphContent(RtfParagraph paragraph, RtfHtmlSaveOptions options, RtfDocument document) {
        if (paragraph.Inlines.Count == 0) {
            return null;
        }

        var builder = new StringBuilder();
        AppendParagraph(builder, paragraph, options, document);
        return EncodeString(builder.ToString());
    }

    private static string? EncodeShapeText(RtfShape shape, RtfHtmlSaveOptions options, RtfDocument document) {
        if (shape.TextBoxParagraphs.Count == 0) {
            return null;
        }

        string newline = options.GetNewLine();
        var builder = new StringBuilder();
        for (int index = 0; index < shape.TextBoxParagraphs.Count; index++) {
            if (index > 0) {
                builder.Append(newline);
            }

            AppendParagraph(builder, shape.TextBoxParagraphs[index], options, document);
        }

        return EncodeString(builder.ToString());
    }

    private static void AppendImageMetadata(StringBuilder builder, string prefix, RtfImage? image) {
        if (image == null) {
            return;
        }

        AppendMetadataAttribute(builder, prefix + "-format", FormatImageFormat(image.Format));
        AppendMetadataAttribute(builder, prefix + "-data", EncodeBytes(image.Data));
        AppendMetadataAttribute(builder, prefix + "-description", image.Description);
        AppendMetadataAttribute(builder, prefix + "-source-width", image.SourceWidth);
        AppendMetadataAttribute(builder, prefix + "-source-height", image.SourceHeight);
        AppendMetadataAttribute(builder, prefix + "-desired-width", image.DesiredWidthTwips);
        AppendMetadataAttribute(builder, prefix + "-desired-height", image.DesiredHeightTwips);
    }

    private static string? EncodeShapeInstructions(IReadOnlyList<RtfShapeInstruction> instructions) {
        if (instructions.Count == 0) {
            return null;
        }

        var builder = new StringBuilder();
        foreach (RtfShapeInstruction instruction in instructions) {
            builder.Append(EncodeString(instruction.Name));
            builder.Append('|');
            builder.Append(instruction.Parameter?.ToString(CultureInfo.InvariantCulture) ?? string.Empty);
            builder.Append('|');
            builder.Append(instruction.HasParameter ? '1' : '0');
            builder.AppendLine();
        }

        return EncodeString(builder.ToString());
    }

    private static string? EncodeShapeProperties(IReadOnlyList<RtfShapeProperty> properties) {
        if (properties.Count == 0) {
            return null;
        }

        var builder = new StringBuilder();
        foreach (RtfShapeProperty property in properties) {
            builder.Append(EncodeString(property.Name));
            builder.Append('|');
            builder.Append(EncodeString(property.Value));
            builder.AppendLine();
        }

        return EncodeString(builder.ToString());
    }

    private static string FormatObjectKind(RtfObjectKind kind) {
        switch (kind) {
            case RtfObjectKind.Embedded:
                return "embedded";
            case RtfObjectKind.Linked:
                return "linked";
            case RtfObjectKind.AutoLinked:
                return "auto-linked";
            case RtfObjectKind.Subscription:
                return "subscription";
            case RtfObjectKind.Publisher:
                return "publisher";
            case RtfObjectKind.IconEmbedded:
                return "icon-embedded";
            default:
                return "unknown";
        }
    }

    private static string FormatImageFormat(RtfImageFormat format) {
        switch (format) {
            case RtfImageFormat.Png:
                return "png";
            case RtfImageFormat.Jpeg:
                return "jpeg";
            case RtfImageFormat.Dib:
                return "dib";
            case RtfImageFormat.Wmf:
                return "wmf";
            case RtfImageFormat.Emf:
                return "emf";
            default:
                return "unknown";
        }
    }

    private static string? EncodeBytes(byte[] data) {
        return data.Length == 0 ? null : Convert.ToBase64String(data);
    }

    private static string EncodeString(string value) {
        return Convert.ToBase64String(Encoding.UTF8.GetBytes(value));
    }

    private static void AppendMetadataAttribute(StringBuilder builder, string name, int? value) {
        if (!value.HasValue) {
            return;
        }

        AppendMetadataAttribute(builder, name, value.Value.ToString(CultureInfo.InvariantCulture));
    }

    private static void AppendMetadataAttribute(StringBuilder builder, string name, string? value) {
        if (string.IsNullOrEmpty(value)) {
            return;
        }

        builder.Append(' ');
        builder.Append(name);
        builder.Append("=\"");
        builder.Append(EncodeAttribute(value!));
        builder.Append('"');
    }
}
