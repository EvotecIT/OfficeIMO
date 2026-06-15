using System.Text;

namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlWriter {
    private static void AppendCharacterEffectsStyle(StringBuilder builder, RtfRun run) {
        if (run.Hidden) {
            builder.Append("visibility:hidden;--officeimo-rtf-hidden:true;");
        }

        if (run.Outline) {
            builder.Append("--officeimo-rtf-outline:true;");
        }

        if (run.Shadow) {
            builder.Append("text-shadow:1pt 1pt 0 currentColor;--officeimo-rtf-shadow:true;");
        }

        if (run.Emboss) {
            builder.Append("--officeimo-rtf-emboss:true;");
        }

        if (run.Imprint) {
            builder.Append("--officeimo-rtf-imprint:true;");
        }
    }
}
