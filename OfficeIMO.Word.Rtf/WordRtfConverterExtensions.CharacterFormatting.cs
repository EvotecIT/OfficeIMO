using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Rtf;
using OfficeIMO.Word;

namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    private static void ApplyWordRunFormatting(WordParagraph wordRun, RtfRun run, RtfDocument rtfDocument) {
        if (!string.IsNullOrWhiteSpace(wordRun.CharacterStyleId)) {
            run.StyleId = FindRtfStyleId(rtfDocument, wordRun.CharacterStyleId!, RtfStyleKind.Character);
        }

        run.Bold = wordRun.Bold;
        run.Italic = wordRun.Italic;
        run.UnderlineStyle = ToRtfUnderlineStyle(wordRun.Underline);
        if (TryGetUnderlineColor(wordRun, out byte underlineRed, out byte underlineGreen, out byte underlineBlue)) {
            run.UnderlineColorIndex = GetOrAddColor(rtfDocument, underlineRed, underlineGreen, underlineBlue);
        }

        run.Strike = wordRun.Strike;
        run.DoubleStrike = wordRun.DoubleStrike;
        run.Hidden = IsHiddenWordRun(wordRun);
        run.Outline = wordRun.Outline;
        run.Shadow = wordRun.Shadow;
        run.Emboss = wordRun.Emboss;
        run.Imprint = IsImprintWordRun(wordRun);
        run.CapsStyle = ToRtfCapsStyle(wordRun.CapsStyle);
        if (wordRun.FontFamily is string fontFamily && !string.IsNullOrWhiteSpace(fontFamily)) {
            run.FontId = rtfDocument.AddFont(fontFamily);
        }

        if (!string.IsNullOrWhiteSpace(wordRun.ColorHex) &&
            TryParseHexColor(wordRun.ColorHex, out byte foregroundRed, out byte foregroundGreen, out byte foregroundBlue)) {
            run.ForegroundColorIndex = GetOrAddColor(rtfDocument, foregroundRed, foregroundGreen, foregroundBlue);
        }

        if (wordRun.Highlight.HasValue &&
            TryGetRtfHighlightColor(wordRun.Highlight.Value, out byte highlightRed, out byte highlightGreen, out byte highlightBlue)) {
            run.HighlightColorIndex = GetOrAddColor(rtfDocument, highlightRed, highlightGreen, highlightBlue);
        }

        run.CharacterSpacingTwips = wordRun.Spacing == 0 ? null : wordRun.Spacing;
        run.CharacterScalePercent = GetCharacterScale(wordRun);
        run.KerningHalfPoints = GetKerningHalfPoints(wordRun);
        run.CharacterOffsetHalfPoints = GetCharacterOffsetHalfPoints(wordRun);
        run.LanguageId = ToRtfLanguageId(wordRun._run?.RunProperties?.Languages?.Val?.Value);
        run.Direction = IsRightToLeftWordRun(wordRun) ? RtfTextDirection.RightToLeft : null;

        run.VerticalPosition = wordRun.VerticalTextAlignment == VerticalPositionValues.Superscript
            ? RtfVerticalPosition.Superscript
            : wordRun.VerticalTextAlignment == VerticalPositionValues.Subscript
                ? RtfVerticalPosition.Subscript
                : RtfVerticalPosition.Baseline;
    }

    private static void ApplyRtfRunFormatting(RtfRun run, WordParagraph wordRun, RtfDocument? rtfDocument) {
        if (run.StyleId.HasValue && rtfDocument != null && HasRtfStyle(rtfDocument, run.StyleId.Value, RtfStyleKind.Character)) {
            wordRun.CharacterStyleId = GetWordStyleId(run.StyleId.Value, RtfStyleKind.Character);
        }

        wordRun.Bold = run.Bold;
        wordRun.Italic = run.Italic;
        if (run.UnderlineStyle != RtfUnderlineStyle.None) {
            wordRun.SetUnderline(ToWordUnderlineStyle(run.UnderlineStyle));
            if (run.UnderlineColorIndex.HasValue && rtfDocument != null) {
                string? underlineColor = GetColorHex(rtfDocument, run.UnderlineColorIndex.Value);
                if (underlineColor != null) {
                    SetUnderlineColor(wordRun, underlineColor);
                }
            }
        }

        wordRun.Strike = run.Strike;
        wordRun.DoubleStrike = run.DoubleStrike;
        wordRun.CapsStyle = ToWordCapsStyle(run.CapsStyle);
        wordRun.Outline = run.Outline;
        wordRun.Shadow = run.Shadow;
        wordRun.Emboss = run.Emboss;
        if (run.Hidden) {
            SetHiddenWordRun(wordRun);
        }

        if (run.Imprint) {
            SetImprintWordRun(wordRun);
        }

        if (run.Direction.HasValue) {
            SetRunDirection(wordRun, run.Direction.Value);
        }

        if (run.FontId.HasValue &&
            TryGetFontName(rtfDocument, run.FontId.Value, out string? fontName)) {
            wordRun.FontFamily = fontName;
        }

        if (rtfDocument != null && run.ForegroundColorIndex.HasValue) {
            string? colorHex = GetColorHex(rtfDocument, run.ForegroundColorIndex.Value);
            if (colorHex != null) {
                wordRun.ColorHex = colorHex;
            }
        }

        if (run.HighlightColorIndex.HasValue &&
            TryGetWordHighlightColor(rtfDocument, run.HighlightColorIndex.Value, out HighlightColorValues highlight)) {
            wordRun.SetHighlight(highlight);
        }

        wordRun.Spacing = run.CharacterSpacingTwips;
        if (run.CharacterScalePercent.HasValue) {
            SetCharacterScale(wordRun, run.CharacterScalePercent.Value);
        }

        if (run.KerningHalfPoints.HasValue) {
            SetKerning(wordRun, run.KerningHalfPoints.Value);
        }

        if (run.CharacterOffsetHalfPoints.HasValue) {
            SetCharacterOffset(wordRun, run.CharacterOffsetHalfPoints.Value);
        }

        if (run.LanguageId.HasValue) {
            SetRunLanguage(wordRun, run.LanguageId.Value);
        }

        if (run.VerticalPosition == RtfVerticalPosition.Superscript) {
            wordRun.SetVerticalTextAlignment(VerticalPositionValues.Superscript);
        } else if (run.VerticalPosition == RtfVerticalPosition.Subscript) {
            wordRun.SetVerticalTextAlignment(VerticalPositionValues.Subscript);
        }

        if (run.FontSize.HasValue) {
            wordRun.FontSize = (int)Math.Round(run.FontSize.Value, MidpointRounding.AwayFromZero);
        }
    }

    private static RtfCapsStyle ToRtfCapsStyle(CapsStyle capsStyle) {
        switch (capsStyle) {
            case CapsStyle.Caps:
                return RtfCapsStyle.Caps;
            case CapsStyle.SmallCaps:
                return RtfCapsStyle.SmallCaps;
            default:
                return RtfCapsStyle.None;
        }
    }

    private static CapsStyle ToWordCapsStyle(RtfCapsStyle capsStyle) {
        switch (capsStyle) {
            case RtfCapsStyle.Caps:
                return CapsStyle.Caps;
            case RtfCapsStyle.SmallCaps:
                return CapsStyle.SmallCaps;
            default:
                return CapsStyle.None;
        }
    }

    private static RtfUnderlineStyle ToRtfUnderlineStyle(UnderlineValues? underline) {
        if (!underline.HasValue) {
            return RtfUnderlineStyle.None;
        }

        UnderlineValues value = underline.Value;
        if (value == UnderlineValues.Single) return RtfUnderlineStyle.Single;
        if (value == UnderlineValues.Words) return RtfUnderlineStyle.Words;
        if (value == UnderlineValues.Double) return RtfUnderlineStyle.Double;
        if (value == UnderlineValues.Thick) return RtfUnderlineStyle.Thick;
        if (value == UnderlineValues.Dotted) return RtfUnderlineStyle.Dotted;
        if (value == UnderlineValues.DottedHeavy) return RtfUnderlineStyle.ThickDotted;
        if (value == UnderlineValues.Dash) return RtfUnderlineStyle.Dash;
        if (value == UnderlineValues.DashedHeavy) return RtfUnderlineStyle.ThickDash;
        if (value == UnderlineValues.DashLong) return RtfUnderlineStyle.LongDash;
        if (value == UnderlineValues.DashLongHeavy) return RtfUnderlineStyle.ThickLongDash;
        if (value == UnderlineValues.DotDash) return RtfUnderlineStyle.DashDot;
        if (value == UnderlineValues.DashDotHeavy) return RtfUnderlineStyle.ThickDashDot;
        if (value == UnderlineValues.DotDotDash) return RtfUnderlineStyle.DashDotDot;
        if (value == UnderlineValues.DashDotDotHeavy) return RtfUnderlineStyle.ThickDashDotDot;
        if (value == UnderlineValues.Wave) return RtfUnderlineStyle.Wave;
        if (value == UnderlineValues.WavyHeavy) return RtfUnderlineStyle.HeavyWave;
        if (value == UnderlineValues.WavyDouble) return RtfUnderlineStyle.DoubleWave;
        return RtfUnderlineStyle.None;
    }

    private static UnderlineValues ToWordUnderlineStyle(RtfUnderlineStyle style) {
        switch (style) {
            case RtfUnderlineStyle.Words:
                return UnderlineValues.Words;
            case RtfUnderlineStyle.Double:
                return UnderlineValues.Double;
            case RtfUnderlineStyle.Dotted:
                return UnderlineValues.Dotted;
            case RtfUnderlineStyle.Dash:
                return UnderlineValues.Dash;
            case RtfUnderlineStyle.DashDot:
                return UnderlineValues.DotDash;
            case RtfUnderlineStyle.DashDotDot:
                return UnderlineValues.DotDotDash;
            case RtfUnderlineStyle.Thick:
                return UnderlineValues.Thick;
            case RtfUnderlineStyle.ThickDotted:
                return UnderlineValues.DottedHeavy;
            case RtfUnderlineStyle.ThickDash:
                return UnderlineValues.DashedHeavy;
            case RtfUnderlineStyle.ThickDashDot:
                return UnderlineValues.DashDotHeavy;
            case RtfUnderlineStyle.ThickDashDotDot:
                return UnderlineValues.DashDotDotHeavy;
            case RtfUnderlineStyle.Wave:
                return UnderlineValues.Wave;
            case RtfUnderlineStyle.HeavyWave:
                return UnderlineValues.WavyHeavy;
            case RtfUnderlineStyle.DoubleWave:
                return UnderlineValues.WavyDouble;
            case RtfUnderlineStyle.LongDash:
                return UnderlineValues.DashLong;
            case RtfUnderlineStyle.ThickLongDash:
                return UnderlineValues.DashLongHeavy;
            default:
                return UnderlineValues.Single;
        }
    }

    private static bool IsHiddenWordRun(WordParagraph wordRun) {
        return wordRun._run?.RunProperties?.Vanish != null;
    }

    private static bool IsImprintWordRun(WordParagraph wordRun) {
        return wordRun._run?.RunProperties?.Imprint != null;
    }

    private static bool IsRightToLeftWordRun(WordParagraph wordRun) {
        return wordRun._run?.RunProperties?.RightToLeftText != null;
    }

    private static void SetRunDirection(WordParagraph wordRun, RtfTextDirection direction) {
        if (wordRun._run == null) {
            return;
        }

        wordRun._run.RunProperties ??= new RunProperties();
        if (direction == RtfTextDirection.RightToLeft) {
            wordRun._run.RunProperties.RightToLeftText = new RightToLeftText();
        } else {
            wordRun._run.RunProperties.RightToLeftText?.Remove();
        }
    }

    private static void SetImprintWordRun(WordParagraph wordRun) {
        if (wordRun._run == null) {
            return;
        }

        wordRun._run.RunProperties ??= new RunProperties();
        wordRun._run.RunProperties.Imprint = new Imprint();
    }

    private static bool TryGetUnderlineColor(WordParagraph wordRun, out byte red, out byte green, out byte blue) {
        red = 0;
        green = 0;
        blue = 0;
        string? color = wordRun._run?.RunProperties?.Underline?.Color?.Value;
        if (string.IsNullOrWhiteSpace(color) ||
            string.Equals(color, "auto", StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        return TryParseHexColor(color!, out red, out green, out blue);
    }

    private static void SetUnderlineColor(WordParagraph wordRun, string colorHex) {
        if (wordRun._run == null) {
            return;
        }

        wordRun._run.RunProperties ??= new RunProperties();
        wordRun._run.RunProperties.Underline ??= new Underline();
        wordRun._run.RunProperties.Underline.Color = colorHex;
    }

    private static int? GetCharacterScale(WordParagraph wordRun) {
        long? value = wordRun._run?.RunProperties?.CharacterScale?.Val?.Value;
        if (!value.HasValue || value.Value == 100 || value.Value > int.MaxValue || value.Value < int.MinValue) {
            return null;
        }

        return (int)value.Value;
    }

    private static int? GetKerningHalfPoints(WordParagraph wordRun) {
        uint? value = wordRun._run?.RunProperties?.Kern?.Val?.Value;
        if (!value.HasValue || value.Value == 0 || value.Value > int.MaxValue) {
            return null;
        }

        return (int)value.Value;
    }

    private static int? GetCharacterOffsetHalfPoints(WordParagraph wordRun) {
        string? value = wordRun._run?.RunProperties?.Position?.Val?.Value;
        if (!int.TryParse(value, out int offset) || offset == 0) {
            return null;
        }

        return offset;
    }

    private static void SetCharacterScale(WordParagraph wordRun, int percent) {
        if (wordRun._run == null) {
            return;
        }

        wordRun._run.RunProperties ??= new RunProperties();
        wordRun._run.RunProperties.CharacterScale = new CharacterScale { Val = percent };
    }

    private static void SetKerning(WordParagraph wordRun, int halfPoints) {
        if (wordRun._run == null || halfPoints < 0) {
            return;
        }

        wordRun._run.RunProperties ??= new RunProperties();
        wordRun._run.RunProperties.Kern = new Kern { Val = (uint)halfPoints };
    }

    private static void SetCharacterOffset(WordParagraph wordRun, int halfPoints) {
        if (wordRun._run == null) {
            return;
        }

        wordRun._run.RunProperties ??= new RunProperties();
        wordRun._run.RunProperties.Position = new Position { Val = halfPoints.ToString(System.Globalization.CultureInfo.InvariantCulture) };
    }

    private static bool TryGetFontName(RtfDocument? document, int fontId, out string? fontName) {
        fontName = null;
        if (document == null) {
            return false;
        }

        RtfFont? font = document.Fonts.FirstOrDefault(item => item.Id == fontId);
        if (font == null) {
            return false;
        }

        fontName = font.Name;
        return true;
    }
}
