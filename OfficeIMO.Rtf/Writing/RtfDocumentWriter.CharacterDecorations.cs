using System.Globalization;
using System.Text;

namespace OfficeIMO.Rtf.Writing;

internal static partial class RtfDocumentWriter {
    private static void WriteCharacterDecoration(StringBuilder builder, RtfRun run, RunWriteState state) {
        WriteOptionalInt(builder, @"\chcbpat", run.CharacterBackgroundColorIndex, ref state.CharacterBackgroundColorIndex, resetValue: 0);
        WriteOptionalInt(builder, @"\chcfpat", run.CharacterShadingForegroundColorIndex, ref state.CharacterShadingForegroundColorIndex, resetValue: 0);
        WriteOptionalInt(builder, @"\chshdng", run.CharacterShadingPatternPercent, ref state.CharacterShadingPatternPercent, resetValue: 0);
        if (state.CharacterShadingPattern != run.CharacterShadingPattern) {
            string? control = GetCharacterShadingPatternControl(run.CharacterShadingPattern);
            if (control != null) {
                builder.Append(control);
                builder.Append(' ');
            }

            state.CharacterShadingPattern = run.CharacterShadingPattern;
        }

        if (run.CharacterBorder.HasAnyValue && !CharacterBorderEquals(run.CharacterBorder, state.CharacterBorder)) {
            WriteCharacterBorder(builder, run.CharacterBorder);
            state.CharacterBorder.CopyFrom(run.CharacterBorder);
        }
    }

    private static string? GetCharacterShadingPatternControl(RtfShadingPattern pattern) {
        switch (pattern) {
            case RtfShadingPattern.Horizontal:
                return @"\chbghoriz";
            case RtfShadingPattern.Vertical:
                return @"\chbgvert";
            case RtfShadingPattern.ForwardDiagonal:
                return @"\chbgfdiag";
            case RtfShadingPattern.BackwardDiagonal:
                return @"\chbgbdiag";
            case RtfShadingPattern.Cross:
                return @"\chbgcross";
            case RtfShadingPattern.DiagonalCross:
                return @"\chbgdcross";
            case RtfShadingPattern.DarkHorizontal:
                return @"\chbgdkhoriz";
            case RtfShadingPattern.DarkVertical:
                return @"\chbgdkvert";
            case RtfShadingPattern.DarkForwardDiagonal:
                return @"\chbgdkfdiag";
            case RtfShadingPattern.DarkBackwardDiagonal:
                return @"\chbgdkbdiag";
            case RtfShadingPattern.DarkCross:
                return @"\chbgdkcross";
            case RtfShadingPattern.DarkDiagonalCross:
                return @"\chbgdkdcross";
            default:
                return null;
        }
    }

    private static void WriteCharacterBorder(StringBuilder builder, RtfCharacterBorder border) {
        builder.Append(@"\chbrdr");
        builder.Append(border.Style switch {
            RtfParagraphBorderStyle.Double => @"\brdrdb",
            RtfParagraphBorderStyle.Dotted => @"\brdrdot",
            RtfParagraphBorderStyle.Dashed => @"\brdrdash",
            RtfParagraphBorderStyle.None => @"\brdrnil",
            _ => @"\brdrs"
        });
        if (border.Width.HasValue) {
            builder.Append(@"\brdrw");
            builder.Append(border.Width.Value.ToString(CultureInfo.InvariantCulture));
        }

        if (border.ColorIndex.HasValue) {
            builder.Append(@"\brdrcf");
            builder.Append(border.ColorIndex.Value.ToString(CultureInfo.InvariantCulture));
        }

        builder.Append(' ');
    }

    private static bool CharacterBorderEquals(RtfCharacterBorder left, RtfCharacterBorder right) =>
        left.Style == right.Style &&
        left.Width == right.Width &&
        left.ColorIndex == right.ColorIndex;
}
