using System.Globalization;
using System.Text;

namespace OfficeIMO.Rtf.Writing;

internal static partial class RtfDocumentWriter {
    private static void WriteRunPrefix(StringBuilder builder, RtfRun run, RunWriteState state) {
        if (NeedsPlainReset(run, state)) {
            builder.Append(@"\plain ");
            state.ResetCharacter();
        }

        WriteRevisionPrefix(builder, run, state);
        Toggle(builder, @"\b", run.Bold, ref state.Bold);
        Toggle(builder, @"\i", run.Italic, ref state.Italic);
        if (state.UnderlineStyle != run.UnderlineStyle) {
            builder.Append(GetUnderlineControl(run.UnderlineStyle));
            builder.Append(' ');
            state.UnderlineStyle = run.UnderlineStyle;
        }
        Toggle(builder, @"\strike", run.Strike, ref state.Strike);
        Toggle(builder, @"\striked", run.DoubleStrike, ref state.DoubleStrike);
        Toggle(builder, @"\v", run.Hidden, ref state.Hidden);
        Toggle(builder, @"\outl", run.Outline, ref state.Outline);
        Toggle(builder, @"\shad", run.Shadow, ref state.Shadow);
        Toggle(builder, @"\embo", run.Emboss, ref state.Emboss);
        Toggle(builder, @"\impr", run.Imprint, ref state.Imprint);
        if (state.Direction != run.Direction) {
            if (run.Direction.HasValue) {
                builder.Append(run.Direction.Value == RtfTextDirection.RightToLeft ? @"\rtlch" : @"\ltrch");
                builder.Append(' ');
            }

            state.Direction = run.Direction;
        }

        if (state.CapsStyle != run.CapsStyle) {
            if (state.CapsStyle == RtfCapsStyle.Caps) {
                builder.Append(@"\caps0 ");
            } else if (state.CapsStyle == RtfCapsStyle.SmallCaps) {
                builder.Append(@"\scaps0 ");
            }

            if (run.CapsStyle == RtfCapsStyle.Caps) {
                builder.Append(@"\caps ");
            } else if (run.CapsStyle == RtfCapsStyle.SmallCaps) {
                builder.Append(@"\scaps ");
            }

            state.CapsStyle = run.CapsStyle;
        }

        if (state.VerticalPosition != run.VerticalPosition) {
            builder.Append(run.VerticalPosition switch {
                RtfVerticalPosition.Superscript => @"\super",
                RtfVerticalPosition.Subscript => @"\sub",
                _ => @"\nosupersub"
            });
            builder.Append(' ');
            state.VerticalPosition = run.VerticalPosition;
        }

        if (run.StyleId.HasValue && state.StyleId != run.StyleId.Value) {
            builder.Append(@"\cs");
            builder.Append(run.StyleId.Value.ToString(CultureInfo.InvariantCulture));
            builder.Append(' ');
            state.StyleId = run.StyleId.Value;
        }

        if (run.FontSize.HasValue && !Nullable.Equals(state.FontSize, run.FontSize)) {
            int halfPoints = (int)Math.Round(run.FontSize.Value * 2d, MidpointRounding.AwayFromZero);
            builder.Append(@"\fs");
            builder.Append(halfPoints.ToString(CultureInfo.InvariantCulture));
            builder.Append(' ');
            state.FontSize = run.FontSize;
        }

        if (run.FontId.HasValue && state.FontId != run.FontId.Value) {
            builder.Append(@"\f");
            builder.Append(run.FontId.Value.ToString(CultureInfo.InvariantCulture));
            builder.Append(' ');
            state.FontId = run.FontId.Value;
        }

        if (run.ForegroundColorIndex.HasValue && state.ForegroundColorIndex != run.ForegroundColorIndex.Value) {
            builder.Append(@"\cf");
            builder.Append(run.ForegroundColorIndex.Value.ToString(CultureInfo.InvariantCulture));
            builder.Append(' ');
            state.ForegroundColorIndex = run.ForegroundColorIndex.Value;
        } else if (!run.ForegroundColorIndex.HasValue && state.ForegroundColorIndex.HasValue) {
            builder.Append(@"\cf0 ");
            state.ForegroundColorIndex = null;
        }

        if (run.HighlightColorIndex.HasValue && state.HighlightColorIndex != run.HighlightColorIndex.Value) {
            builder.Append(@"\highlight");
            builder.Append(run.HighlightColorIndex.Value.ToString(CultureInfo.InvariantCulture));
            builder.Append(' ');
            state.HighlightColorIndex = run.HighlightColorIndex.Value;
        } else if (!run.HighlightColorIndex.HasValue && state.HighlightColorIndex.HasValue) {
            builder.Append(@"\highlight0 ");
            state.HighlightColorIndex = null;
        }

        WriteCharacterDecoration(builder, run, state);

        if (run.UnderlineColorIndex.HasValue && state.UnderlineColorIndex != run.UnderlineColorIndex.Value) {
            builder.Append(@"\ulc");
            builder.Append(run.UnderlineColorIndex.Value.ToString(CultureInfo.InvariantCulture));
            builder.Append(' ');
            state.UnderlineColorIndex = run.UnderlineColorIndex.Value;
        } else if (!run.UnderlineColorIndex.HasValue && state.UnderlineColorIndex.HasValue) {
            builder.Append(@"\ulc0 ");
            state.UnderlineColorIndex = null;
        }

        WriteOptionalInt(builder, @"\expndtw", run.CharacterSpacingTwips, ref state.CharacterSpacingTwips, resetValue: 0);
        WriteOptionalInt(builder, @"\charscalex", run.CharacterScalePercent, ref state.CharacterScalePercent, resetValue: 100);
        WriteOptionalInt(builder, @"\kerning", run.KerningHalfPoints, ref state.KerningHalfPoints, resetValue: 0);
        WriteCharacterOffset(builder, run.CharacterOffsetHalfPoints, state);
        WriteLanguage(builder, run, state);
    }

    private static bool NeedsPlainReset(RtfRun run, RunWriteState state) =>
        (!run.StyleId.HasValue && state.StyleId.HasValue) ||
        (!run.FontSize.HasValue && state.FontSize.HasValue) ||
        (!run.FontId.HasValue && state.FontId.HasValue) ||
        (!run.Direction.HasValue && state.Direction.HasValue) ||
        (state.CharacterShadingPattern != RtfShadingPattern.None && run.CharacterShadingPattern == RtfShadingPattern.None) ||
        (state.CharacterBorder.HasAnyValue && !run.CharacterBorder.HasAnyValue);

    private static void Toggle(StringBuilder builder, string onControl, bool desired, ref bool current, string? offControl = null) {
        if (desired == current) return;
        builder.Append(desired ? onControl : offControl ?? onControl + "0");
        builder.Append(' ');
        current = desired;
    }

    private static void WriteOptionalInt(StringBuilder builder, string control, int? desired, ref int? current, int resetValue) {
        if (desired.HasValue && current != desired.Value) {
            builder.Append(control);
            builder.Append(desired.Value.ToString(CultureInfo.InvariantCulture));
            builder.Append(' ');
            current = desired.Value;
        } else if (!desired.HasValue && current.HasValue) {
            builder.Append(control);
            builder.Append(resetValue.ToString(CultureInfo.InvariantCulture));
            builder.Append(' ');
            current = null;
        }
    }

    private static void WriteCharacterOffset(StringBuilder builder, int? desired, RunWriteState state) {
        if (desired == state.CharacterOffsetHalfPoints) {
            return;
        }

        if (desired.HasValue) {
            builder.Append(desired.Value < 0 ? @"\dn" : @"\up");
            builder.Append(Math.Abs(desired.Value).ToString(CultureInfo.InvariantCulture));
            builder.Append(' ');
            state.CharacterOffsetHalfPoints = desired.Value;
        } else {
            builder.Append(@"\up0 ");
            state.CharacterOffsetHalfPoints = null;
        }
    }

    private static void WriteLanguage(StringBuilder builder, RtfRun run, RunWriteState state) {
        int? desired = run.LanguageId ?? state.DefaultLanguageId;
        if (desired == state.LanguageId) {
            return;
        }

        builder.Append(@"\lang");
        builder.Append(desired.GetValueOrDefault(0).ToString(CultureInfo.InvariantCulture));
        builder.Append(' ');
        state.LanguageId = desired;
    }

    private static string GetUnderlineControl(RtfUnderlineStyle style) {
        switch (style) {
            case RtfUnderlineStyle.Single:
                return @"\ul";
            case RtfUnderlineStyle.Words:
                return @"\ulw";
            case RtfUnderlineStyle.Double:
                return @"\uldb";
            case RtfUnderlineStyle.Dotted:
                return @"\uld";
            case RtfUnderlineStyle.Dash:
                return @"\uldash";
            case RtfUnderlineStyle.DashDot:
                return @"\uldashd";
            case RtfUnderlineStyle.DashDotDot:
                return @"\uldashdd";
            case RtfUnderlineStyle.Thick:
                return @"\ulth";
            case RtfUnderlineStyle.ThickDotted:
                return @"\ulthd";
            case RtfUnderlineStyle.ThickDash:
                return @"\ulthdash";
            case RtfUnderlineStyle.ThickDashDot:
                return @"\ulthdashd";
            case RtfUnderlineStyle.ThickDashDotDot:
                return @"\ulthdashdd";
            case RtfUnderlineStyle.Wave:
                return @"\ulwave";
            case RtfUnderlineStyle.HeavyWave:
                return @"\ulhwave";
            case RtfUnderlineStyle.DoubleWave:
                return @"\uldbwave";
            case RtfUnderlineStyle.LongDash:
                return @"\ulldash";
            case RtfUnderlineStyle.ThickLongDash:
                return @"\ulthldash";
            default:
                return @"\ulnone";
        }
    }

    private static void ResetRunState(StringBuilder builder, RunWriteState state) {
        if (state.StyleId.HasValue || state.FontSize.HasValue || state.FontId.HasValue ||
            state.CharacterShadingPattern != RtfShadingPattern.None || state.CharacterBorder.HasAnyValue) {
            builder.Append(@"\plain ");
            state.ResetCharacter();
            return;
        }

        if (state.Bold) builder.Append(@"\b0 ");
        if (state.Italic) builder.Append(@"\i0 ");
        if (state.UnderlineStyle != RtfUnderlineStyle.None) builder.Append(@"\ulnone ");
        if (state.Strike) builder.Append(@"\strike0 ");
        if (state.DoubleStrike) builder.Append(@"\striked0 ");
        if (state.Hidden) builder.Append(@"\v0 ");
        if (state.Outline) builder.Append(@"\outl0 ");
        if (state.Shadow) builder.Append(@"\shad0 ");
        if (state.Emboss) builder.Append(@"\embo0 ");
        if (state.Imprint) builder.Append(@"\impr0 ");
        if (state.Direction.HasValue) builder.Append(@"\ltrch ");
        if (state.CapsStyle == RtfCapsStyle.Caps) builder.Append(@"\caps0 ");
        if (state.CapsStyle == RtfCapsStyle.SmallCaps) builder.Append(@"\scaps0 ");
        if (state.VerticalPosition != RtfVerticalPosition.Baseline) builder.Append(@"\nosupersub ");
        if (state.ForegroundColorIndex.HasValue) builder.Append(@"\cf0 ");
        if (state.HighlightColorIndex.HasValue) builder.Append(@"\highlight0 ");
        if (state.CharacterBackgroundColorIndex.HasValue) builder.Append(@"\chcbpat0 ");
        if (state.CharacterShadingForegroundColorIndex.HasValue) builder.Append(@"\chcfpat0 ");
        if (state.CharacterShadingPatternPercent.HasValue) builder.Append(@"\chshdng0 ");
        if (state.UnderlineColorIndex.HasValue) builder.Append(@"\ulc0 ");
        if (state.CharacterSpacingTwips.HasValue) builder.Append(@"\expndtw0 ");
        if (state.CharacterScalePercent.HasValue) builder.Append(@"\charscalex100 ");
        if (state.KerningHalfPoints.HasValue) builder.Append(@"\kerning0 ");
        if (state.CharacterOffsetHalfPoints.HasValue) builder.Append(@"\up0 ");
        if (state.LanguageId != state.DefaultLanguageId) {
            builder.Append(@"\lang");
            builder.Append(state.DefaultLanguageId.GetValueOrDefault(0).ToString(CultureInfo.InvariantCulture));
            builder.Append(' ');
        }
        ResetRevisionState(builder, state);
        state.ResetCharacter();
    }

    private sealed class RunWriteState {
        public RunWriteState(int? defaultLanguageId = null) {
            DefaultLanguageId = defaultLanguageId;
            LanguageId = defaultLanguageId;
        }

        public bool Bold;
        public bool Italic;
        public RtfUnderlineStyle UnderlineStyle;
        public bool Strike;
        public bool DoubleStrike;
        public bool Hidden;
        public bool Outline;
        public bool Shadow;
        public bool Emboss;
        public bool Imprint;
        public RtfTextDirection? Direction;
        public RtfCapsStyle CapsStyle;
        public RtfVerticalPosition VerticalPosition;
        public double? FontSize;
        public int? FontId;
        public int? ForegroundColorIndex;
        public int? HighlightColorIndex;
        public int? CharacterBackgroundColorIndex;
        public int? CharacterShadingForegroundColorIndex;
        public int? CharacterShadingPatternPercent;
        public RtfShadingPattern CharacterShadingPattern;
        public RtfCharacterBorder CharacterBorder { get; } = new RtfCharacterBorder();
        public int? UnderlineColorIndex;
        public int? CharacterSpacingTwips;
        public int? CharacterScalePercent;
        public int? KerningHalfPoints;
        public int? CharacterOffsetHalfPoints;
        public int? StyleId;
        public int? DefaultLanguageId;
        public int? LanguageId;
        public RtfRevisionKind RevisionKind;
        public int? RevisionAuthorIndex;
        public int? RevisionTimestampValue;
        public int? CharacterRevisionSaveId;
        public int? InsertionRevisionSaveId;
        public int? DeletionRevisionSaveId;

        public void ResetCharacter() {
            Bold = false;
            Italic = false;
            UnderlineStyle = RtfUnderlineStyle.None;
            Strike = false;
            DoubleStrike = false;
            Hidden = false;
            Outline = false;
            Shadow = false;
            Emboss = false;
            Imprint = false;
            Direction = null;
            CapsStyle = RtfCapsStyle.None;
            VerticalPosition = RtfVerticalPosition.Baseline;
            FontSize = null;
            FontId = null;
            ForegroundColorIndex = null;
            HighlightColorIndex = null;
            CharacterBackgroundColorIndex = null;
            CharacterShadingForegroundColorIndex = null;
            CharacterShadingPatternPercent = null;
            CharacterShadingPattern = RtfShadingPattern.None;
            CharacterBorder.Clear();
            UnderlineColorIndex = null;
            CharacterSpacingTwips = null;
            CharacterScalePercent = null;
            KerningHalfPoints = null;
            CharacterOffsetHalfPoints = null;
            StyleId = null;
            LanguageId = DefaultLanguageId;
            RevisionKind = RtfRevisionKind.None;
            RevisionAuthorIndex = null;
            RevisionTimestampValue = null;
            CharacterRevisionSaveId = null;
            InsertionRevisionSaveId = null;
            DeletionRevisionSaveId = null;
        }
    }
}
