using System.Text;

namespace OfficeIMO.Drawing;

public static partial class OfficeTextBlockRenderer {
    private static bool AppendSvgRichTextLinkStart(StringBuilder builder, string? target) {
        if (!OfficeDrawingLinkPolicy.TryNormalize(target, out string uri)) return false;
        builder.Append("<a").AppendAttribute("href", uri).Append('>');
        return true;
    }
}
