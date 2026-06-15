using System.Text;

namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlWriter {
    private static void AppendCapsStyle(StringBuilder builder, RtfCapsStyle capsStyle) {
        switch (capsStyle) {
            case RtfCapsStyle.Caps:
                builder.Append("text-transform:uppercase;--officeimo-rtf-caps-style:caps;");
                break;
            case RtfCapsStyle.SmallCaps:
                builder.Append("font-variant-caps:small-caps;--officeimo-rtf-caps-style:small-caps;");
                break;
        }
    }
}
