using System.Globalization;
using System.Text;

namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlWriter {
    private static void AppendLanguageDirectionStyle(StringBuilder builder, RtfRun run) {
        if (run.LanguageId.HasValue) {
            builder.Append("--officeimo-rtf-lang:");
            builder.Append(run.LanguageId.Value.ToString(CultureInfo.InvariantCulture));
            builder.Append(';');
        }

        string? direction = FormatTextDirection(run.Direction);
        if (direction != null) {
            builder.Append("direction:");
            builder.Append(direction);
            builder.Append(";unicode-bidi:isolate;--officeimo-rtf-direction:");
            builder.Append(direction);
            builder.Append(';');
        }
    }

    private static string? FormatLanguageTag(int? languageId) {
        if (!languageId.HasValue) {
            return null;
        }

        try {
            return CultureInfo.GetCultureInfo(languageId.Value).Name;
        } catch (CultureNotFoundException) {
            return null;
        }
    }

    private static string? FormatTextDirection(RtfTextDirection? direction) {
        if (!direction.HasValue) {
            return null;
        }

        return direction.Value == RtfTextDirection.RightToLeft ? "rtl" : "ltr";
    }

    private static bool HasRunSpan(RtfRun run, RtfDocument document) =>
        TryGetRunStyle(run, document, out _) ||
        FormatLanguageTag(run.LanguageId) != null ||
        FormatTextDirection(run.Direction) != null;
}
