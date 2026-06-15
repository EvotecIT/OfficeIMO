using System.Globalization;
using System.Text;

namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlWriter {
    private static void AppendLanguageDirectionStyle(StringBuilder builder, RtfRun run) {
        AppendLanguageDirectionStyle(builder, run.LanguageId, run.Direction);
    }

    private static void AppendLanguageDirectionStyle(StringBuilder builder, int? languageId, RtfTextDirection? direction) {
        if (languageId.HasValue) {
            builder.Append("--officeimo-rtf-lang:");
            builder.Append(languageId.Value.ToString(CultureInfo.InvariantCulture));
            builder.Append(';');
        }

        string? formattedDirection = FormatTextDirection(direction);
        if (formattedDirection != null) {
            builder.Append("direction:");
            builder.Append(formattedDirection);
            builder.Append(";unicode-bidi:isolate;--officeimo-rtf-direction:");
            builder.Append(formattedDirection);
            builder.Append(';');
        }
    }

    private static void AppendLanguageDirectionAttributes(StringBuilder builder, int? languageId, RtfTextDirection? direction) {
        string? language = FormatLanguageTag(languageId);
        if (language != null) {
            builder.Append(" lang=\"");
            builder.Append(EncodeAttribute(language));
            builder.Append('"');
        }

        string? formattedDirection = FormatTextDirection(direction);
        if (formattedDirection != null) {
            builder.Append(" dir=\"");
            builder.Append(formattedDirection);
            builder.Append('"');
        }
    }

    private static void AppendLanguageDirectionStyleAttribute(StringBuilder builder, int? languageId, RtfTextDirection? direction) {
        var styleBuilder = new StringBuilder();
        AppendLanguageDirectionStyle(styleBuilder, languageId, direction);
        if (styleBuilder.Length == 0) {
            return;
        }

        builder.Append(" style=\"");
        builder.Append(EncodeAttribute(styleBuilder.ToString()));
        builder.Append('"');
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
        run.StyleId.HasValue ||
        TryGetRunStyle(run, document, out _) ||
        FormatLanguageTag(run.LanguageId) != null ||
        FormatTextDirection(run.Direction) != null;
}
