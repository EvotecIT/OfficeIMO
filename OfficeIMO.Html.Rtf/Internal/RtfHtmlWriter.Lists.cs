using System.Globalization;

namespace OfficeIMO.Html;

internal static partial class RtfHtmlWriter {
    private sealed class HtmlListFrame {
        internal HtmlListFrame(RtfListFormatting formatting) => Formatting = formatting;
        internal RtfListFormatting Formatting { get; }
        internal bool ItemOpen { get; set; }
    }

    private static void AppendListParagraph(StringBuilder builder, RtfParagraph paragraph, RtfListMarker marker, List<HtmlListFrame> frames, RtfToHtmlOptions options, RtfDocument document) {
        if (!marker.IsNumberFormatSupported) options.AddDiagnostic("RtfHtmlListNumberFormatFlattened", "The requested numbering format was represented by decimal marker text.", "Paragraph/List", action: RtfConversionAction.Flattened);
        if (marker.Formatting.Level.PictureIndex.HasValue) options.AddDiagnostic("RtfHtmlListPictureFlattened", "The picture bullet was represented by its text marker.", "Paragraph/List", action: RtfConversionAction.Flattened);
        int level = marker.Formatting.LevelIndex;
        while (frames.Count > 0 && (frames[frames.Count - 1].Formatting.LevelIndex > level ||
            (frames[frames.Count - 1].Formatting.LevelIndex == level && !SameList(frames[frames.Count - 1].Formatting, marker.Formatting)))) {
            CloseLastList(builder, frames);
        }
        if (frames.Count == 0 || frames[frames.Count - 1].Formatting.LevelIndex < level) {
            OpenEffectiveList(builder, marker);
            frames.Add(new HtmlListFrame(marker.Formatting));
        } else if (frames[frames.Count - 1].ItemOpen) builder.Append("</li>");
        AppendParagraph(builder, paragraph, options, document, marker, closeListItem: false);
        frames[frames.Count - 1].ItemOpen = true;
    }

    private static bool SameList(RtfListFormatting first, RtfListFormatting second) => first.Identity == second.Identity &&
        first.Level.Kind == second.Level.Kind && (first.Level.NumberFormatN ?? first.Level.NumberFormat) == (second.Level.NumberFormatN ?? second.Level.NumberFormat);

    private static void OpenEffectiveList(StringBuilder builder, RtfListMarker marker) {
        RtfListLevel level = marker.Formatting.Level;
        if (level.Kind == RtfListKind.Bullet) { builder.Append("<ul>"); return; }
        builder.Append("<ol");
        if (marker.Value != 1) builder.Append(" start=\"").Append(marker.Value.ToString(CultureInfo.InvariantCulture)).Append('"');
        string? type = (level.NumberFormatN ?? level.NumberFormat) switch { 1 => "I", 2 => "i", 3 => "A", 4 => "a", _ => null };
        if (type != null) builder.Append(" type=\"").Append(type).Append('"');
        builder.Append('>');
    }

    private static bool RequiresExplicitMarker(RtfParagraph paragraph, RtfListMarker marker) {
        RtfListLevel level = marker.Formatting.Level;
        if (!marker.IsNumberFormatSupported || (level.NumberFormatN ?? level.NumberFormat) is 5 or 22) return true;
        if (!marker.Formatting.IsDefined && paragraph.ListText != null) return true;
        string defaultText = level.Kind == RtfListKind.Bullet ? "\u2022" : "%" + (marker.Formatting.LevelIndex + 1).ToString(CultureInfo.InvariantCulture) + ".";
        return level.Text != null && level.Text != defaultText;
    }

    private static void CloseLists(StringBuilder builder, List<HtmlListFrame> frames) {
        while (frames.Count > 0) CloseLastList(builder, frames);
    }

    private static void CloseLastList(StringBuilder builder, List<HtmlListFrame> frames) {
        HtmlListFrame frame = frames[frames.Count - 1];
        if (frame.ItemOpen) builder.Append("</li>");
        CloseList(builder, frame.Formatting.Level.Kind);
        frames.RemoveAt(frames.Count - 1);
    }

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
