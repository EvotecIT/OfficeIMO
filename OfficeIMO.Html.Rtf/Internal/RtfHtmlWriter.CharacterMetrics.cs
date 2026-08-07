using System.Globalization;
using System.Text;

namespace OfficeIMO.Html;

internal static partial class RtfHtmlWriter {
    private static void AppendCharacterMetricsStyle(StringBuilder builder, RtfRun run) {
        if (run.CharacterSpacingTwips.HasValue) {
            builder.Append("letter-spacing:");
            builder.Append(FormatPoints(run.CharacterSpacingTwips.Value / 20d));
            builder.Append("pt;");
        }

        if (run.CharacterScalePercent.HasValue) {
            builder.Append("font-stretch:");
            builder.Append(run.CharacterScalePercent.Value.ToString(CultureInfo.InvariantCulture));
            builder.Append("%;");
            builder.Append("--officeimo-rtf-character-scale:");
            builder.Append(run.CharacterScalePercent.Value.ToString(CultureInfo.InvariantCulture));
            builder.Append(';');
        }

        if (run.CharacterOffsetHalfPoints.HasValue) {
            builder.Append("vertical-align:");
            builder.Append(FormatPoints(run.CharacterOffsetHalfPoints.Value / 2d));
            builder.Append("pt;");
            builder.Append("--officeimo-rtf-character-offset:");
            builder.Append(run.CharacterOffsetHalfPoints.Value.ToString(CultureInfo.InvariantCulture));
            builder.Append(';');
        }
    }
}
