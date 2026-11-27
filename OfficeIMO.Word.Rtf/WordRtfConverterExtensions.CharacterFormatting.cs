using OfficeIMO.Rtf;
using OfficeIMO.Word;

namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    private static void ApplyWordRunFormatting(WordParagraph wordRun, RtfRun run, RtfDocument rtfDocument) {
        if (!string.IsNullOrWhiteSpace(wordRun.CharacterStyleId)) {
            run.StyleId = FindRtfStyleId(rtfDocument, wordRun.CharacterStyleId!, RtfStyleKind.Character);
        }

        OfficeXmlNode scopedProperties = wordRun.ScopedRunPropertiesNode();
        run.DirectBold = ReadToggle(scopedProperties.WChild("b"));
        run.DirectItalic = ReadToggle(scopedProperties.WChild("i"));
        run.DirectUnderlineStyle = scopedProperties.WChild("u").Exists
            ? ToRtfUnderlineStyle(wordRun.Underline) : null;
        if (TryGetUnderlineColor(wordRun, out byte underlineRed, out byte underlineGreen, out byte underlineBlue)) {
            run.UnderlineColorIndex = GetOrAddColor(rtfDocument, underlineRed, underlineGreen, underlineBlue);
        }

        run.Strikethrough = wordRun.Strikethrough;
        run.DoubleStrikethrough = wordRun.DoubleStrikethrough;
        run.DirectHidden = ReadToggle(scopedProperties.WChild("vanish"));
        run.Outline = wordRun.Outline;
        run.Shadow = wordRun.Shadow;
        run.Emboss = wordRun.Emboss;
        run.Imprint = IsImprintWordRun(wordRun);
        run.CapsStyle = ToRtfCapsStyle(wordRun.CapsStyle);
        run.FontSize = wordRun.FontSize?.ToPoints();
        if (wordRun.FontFamily is string fontFamily && !string.IsNullOrWhiteSpace(fontFamily)) {
            run.FontId = rtfDocument.AddFont(fontFamily);
        }

        if (!string.IsNullOrWhiteSpace(wordRun.RunColorValue) &&
            TryParseHexColor(wordRun.RunColorValue, out byte foregroundRed, out byte foregroundGreen, out byte foregroundBlue)) {
            run.ForegroundColorIndex = GetOrAddColor(rtfDocument, foregroundRed, foregroundGreen, foregroundBlue);
        }
        if (string.Equals(wordRun.RunColorValue, "auto", StringComparison.OrdinalIgnoreCase)) run.ForegroundColorIndex = 0;

        if (wordRun.Highlight.HasValue &&
            TryGetRtfHighlightColor(wordRun.Highlight.Value, out byte highlightRed, out byte highlightGreen, out byte highlightBlue)) {
            run.HighlightColorIndex = GetOrAddColor(rtfDocument, highlightRed, highlightGreen, highlightBlue);
        }

        run.CharacterSpacingTwips = wordRun.Spacing == OfficeLength.Zero ? null : checked((int?)wordRun.Spacing?.ToTwips());
        run.CharacterScalePercent = GetCharacterScale(wordRun);
        run.KerningHalfPoints = GetKerningHalfPoints(wordRun);
        run.CharacterOffsetHalfPoints = GetCharacterOffsetHalfPoints(wordRun);
        run.LanguageId = ToRtfLanguageId(DirectRunProperties(wordRun).WChild("lang").WVal());
        run.Direction = IsRightToLeftWordRun(wordRun) ? RtfTextDirection.RightToLeft : null;

        run.VerticalPosition = wordRun.VerticalTextAlignment == WordVerticalTextPosition.Superscript
            ? RtfVerticalPosition.Superscript
            : wordRun.VerticalTextAlignment == WordVerticalTextPosition.Subscript
                ? RtfVerticalPosition.Subscript
                : RtfVerticalPosition.Baseline;
    }

    private static void ApplyRtfRunFormatting(RtfRun run, WordParagraph wordRun, RtfDocument? rtfDocument) {
        if (run.StyleId.HasValue && rtfDocument != null && HasRtfStyle(rtfDocument, run.StyleId.Value, RtfStyleKind.Character)) {
            wordRun.CharacterStyleId = GetWordStyleId(run.StyleId.Value, RtfStyleKind.Character);
        }

        wordRun.ApplyDirectEmphasis(run.DirectBold, run.DirectItalic,
            run.DirectUnderlineStyle.HasValue ? ToWordUnderlineStyle(run.UnderlineStyle) : null);
        if (run.UnderlineStyle != RtfUnderlineStyle.None) {
            wordRun.SetUnderline(ToWordUnderlineStyle(run.UnderlineStyle));
            if (run.UnderlineColorIndex.HasValue && rtfDocument != null) {
                string? underlineColor = GetColorHex(rtfDocument, run.UnderlineColorIndex.Value);
                if (underlineColor != null) {
                    SetUnderlineColor(wordRun, underlineColor);
                }
            }
        }

        wordRun.Strikethrough = run.Strikethrough;
        wordRun.DoubleStrikethrough = run.DoubleStrikethrough;
        wordRun.CapsStyle = ToWordCapsStyle(run.CapsStyle);
        wordRun.Outline = run.Outline;
        wordRun.Shadow = run.Shadow;
        wordRun.Emboss = run.Emboss;
        if (run.DirectHidden.HasValue) {
            SetHiddenWordRun(wordRun, run.Hidden);
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
                wordRun.RunColorValue = colorHex;
            } else if (run.ForegroundColorIndex == 0) {
                wordRun.RunColorValue = "auto";
            }
        }

        if (run.HighlightColorIndex.HasValue &&
            TryGetWordHighlightColor(rtfDocument, run.HighlightColorIndex.Value, out WordHighlightColor highlight)) {
            wordRun.SetHighlight(highlight);
        }

        wordRun.Spacing = run.CharacterSpacingTwips is int nativeSpacing ? OfficeLength.Emus((long)nativeSpacing * 635) : null;
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
            wordRun.SetVerticalTextAlignment(WordVerticalTextPosition.Superscript);
        } else if (run.VerticalPosition == RtfVerticalPosition.Subscript) {
            wordRun.SetVerticalTextAlignment(WordVerticalTextPosition.Subscript);
        }

        if (run.FontSize.HasValue) {
            wordRun.FontSize = global::OfficeIMO.OfficeLength.Points(run.FontSize.Value);
        }
    }

    private static RtfCapsStyle ToRtfCapsStyle(WordCapsStyle capsStyle) {
        switch (capsStyle) {
            case WordCapsStyle.Caps:
                return RtfCapsStyle.Caps;
            case WordCapsStyle.SmallCaps:
                return RtfCapsStyle.SmallCaps;
            default:
                return RtfCapsStyle.None;
        }
    }

    private static WordCapsStyle ToWordCapsStyle(RtfCapsStyle capsStyle) {
        switch (capsStyle) {
            case RtfCapsStyle.Caps:
                return WordCapsStyle.Caps;
            case RtfCapsStyle.SmallCaps:
                return WordCapsStyle.SmallCaps;
            default:
                return WordCapsStyle.None;
        }
    }

    private static RtfUnderlineStyle ToRtfUnderlineStyle(WordUnderlineStyle? underline) {
        if (!underline.HasValue) {
            return RtfUnderlineStyle.None;
        }

        switch (underline.Value) {
            case WordUnderlineStyle.Single: return RtfUnderlineStyle.Single;
            case WordUnderlineStyle.Words: return RtfUnderlineStyle.Words;
            case WordUnderlineStyle.Double: return RtfUnderlineStyle.Double;
            case WordUnderlineStyle.Thick: return RtfUnderlineStyle.Thick;
            case WordUnderlineStyle.Dotted: return RtfUnderlineStyle.Dotted;
            case WordUnderlineStyle.DottedHeavy: return RtfUnderlineStyle.ThickDotted;
            case WordUnderlineStyle.Dash: return RtfUnderlineStyle.Dash;
            case WordUnderlineStyle.DashedHeavy: return RtfUnderlineStyle.ThickDash;
            case WordUnderlineStyle.DashLong: return RtfUnderlineStyle.LongDash;
            case WordUnderlineStyle.DashLongHeavy: return RtfUnderlineStyle.ThickLongDash;
            case WordUnderlineStyle.DotDash: return RtfUnderlineStyle.DashDot;
            case WordUnderlineStyle.DashDotHeavy: return RtfUnderlineStyle.ThickDashDot;
            case WordUnderlineStyle.DotDotDash: return RtfUnderlineStyle.DashDotDot;
            case WordUnderlineStyle.DashDotDotHeavy: return RtfUnderlineStyle.ThickDashDotDot;
            case WordUnderlineStyle.Wave: return RtfUnderlineStyle.Wave;
            case WordUnderlineStyle.WavyHeavy: return RtfUnderlineStyle.HeavyWave;
            case WordUnderlineStyle.WavyDouble: return RtfUnderlineStyle.DoubleWave;
            default: return RtfUnderlineStyle.None;
        }
    }

    private static WordUnderlineStyle ToWordUnderlineStyle(RtfUnderlineStyle style) {
        switch (style) {
            case RtfUnderlineStyle.None:
                return WordUnderlineStyle.None;
            case RtfUnderlineStyle.Words:
                return WordUnderlineStyle.Words;
            case RtfUnderlineStyle.Double:
                return WordUnderlineStyle.Double;
            case RtfUnderlineStyle.Dotted:
                return WordUnderlineStyle.Dotted;
            case RtfUnderlineStyle.Dash:
                return WordUnderlineStyle.Dash;
            case RtfUnderlineStyle.DashDot:
                return WordUnderlineStyle.DotDash;
            case RtfUnderlineStyle.DashDotDot:
                return WordUnderlineStyle.DotDotDash;
            case RtfUnderlineStyle.Thick:
                return WordUnderlineStyle.Thick;
            case RtfUnderlineStyle.ThickDotted:
                return WordUnderlineStyle.DottedHeavy;
            case RtfUnderlineStyle.ThickDash:
                return WordUnderlineStyle.DashedHeavy;
            case RtfUnderlineStyle.ThickDashDot:
                return WordUnderlineStyle.DashDotHeavy;
            case RtfUnderlineStyle.ThickDashDotDot:
                return WordUnderlineStyle.DashDotDotHeavy;
            case RtfUnderlineStyle.Wave:
                return WordUnderlineStyle.Wave;
            case RtfUnderlineStyle.HeavyWave:
                return WordUnderlineStyle.WavyHeavy;
            case RtfUnderlineStyle.DoubleWave:
                return WordUnderlineStyle.WavyDouble;
            case RtfUnderlineStyle.LongDash:
                return WordUnderlineStyle.DashLong;
            case RtfUnderlineStyle.ThickLongDash:
                return WordUnderlineStyle.DashLongHeavy;
            default:
                return WordUnderlineStyle.Single;
        }
    }

    // The w:u w:val token of each underline style, paired by position with UnderlineStyles.
    private static readonly string[] UnderlineTokens = {
        "single", "words", "double", "thick", "dotted", "dottedHeavy", "dash", "dashedHeavy", "dashLong",
        "dashLongHeavy", "dotDash", "dashDotHeavy", "dotDotDash", "dashDotDotHeavy", "wave", "wavyHeavy", "wavyDouble", "none"
    };

    private static readonly WordUnderlineStyle[] UnderlineStyles = {
        WordUnderlineStyle.Single, WordUnderlineStyle.Words, WordUnderlineStyle.Double, WordUnderlineStyle.Thick,
        WordUnderlineStyle.Dotted, WordUnderlineStyle.DottedHeavy, WordUnderlineStyle.Dash, WordUnderlineStyle.DashedHeavy,
        WordUnderlineStyle.DashLong, WordUnderlineStyle.DashLongHeavy, WordUnderlineStyle.DotDash, WordUnderlineStyle.DashDotHeavy,
        WordUnderlineStyle.DotDotDash, WordUnderlineStyle.DashDotDotHeavy, WordUnderlineStyle.Wave, WordUnderlineStyle.WavyHeavy,
        WordUnderlineStyle.WavyDouble, WordUnderlineStyle.None
    };

    private static string ToUnderlineToken(WordUnderlineStyle style) => UnderlineTokens[Array.IndexOf(UnderlineStyles, style)];

    private static WordUnderlineStyle? ParseUnderlineToken(string? token) {
        int index = token == null ? -1 : Array.IndexOf(UnderlineTokens, token);
        return index < 0 ? null : UnderlineStyles[index];
    }

    private static bool IsImprintWordRun(WordParagraph wordRun) {
        return DirectRunProperties(wordRun).WChild("imprint").Exists;
    }

    private static bool IsRightToLeftWordRun(WordParagraph wordRun) {
        return DirectRunProperties(wordRun).WChild("rtl").Exists;
    }

    private static void SetRunDirection(WordParagraph wordRun, RtfTextDirection direction) {
        if (wordRun.RunNode().IsNone) {
            return;
        }

        OfficeXmlNode properties = GetOrAddRunProperties(wordRun);
        if (direction == RtfTextDirection.RightToLeft) {
            WordXmlEdits.ReplaceChild(properties, "rtl", WRPrView.Order);
        } else {
            WordXmlEdits.RemoveChild(properties, "rtl");
        }
    }

    private static void SetImprintWordRun(WordParagraph wordRun) {
        if (wordRun.RunNode().IsNone) {
            return;
        }

        WordXmlEdits.ReplaceChild(GetOrAddRunProperties(wordRun), "imprint", WRPrView.Order);
    }

    private static bool TryGetUnderlineColor(WordParagraph wordRun, out byte red, out byte green, out byte blue) {
        red = 0;
        green = 0;
        blue = 0;
        string? color = DirectRunProperties(wordRun).WChild("u").WAttr("color");
        if (string.IsNullOrWhiteSpace(color) ||
            string.Equals(color, "auto", StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        return TryParseHexColor(color!, out red, out green, out blue);
    }

    private static void SetUnderlineColor(WordParagraph wordRun, string colorHex) {
        if (wordRun.RunNode().IsNone) {
            return;
        }

        GetOrAddRunProperties(wordRun).GetOrAddChild(WordXml.W, "u", WRPrView.Order).SetAttribute(WordXml.W, "color", colorHex);
    }

    private static int? GetCharacterScale(WordParagraph wordRun) {
        long? value = WordXml.ParseInt64(DirectRunProperties(wordRun).WChild("w").WVal());
        if (!value.HasValue || value.Value == 100 || value.Value > int.MaxValue || value.Value < int.MinValue) {
            return null;
        }

        return (int)value.Value;
    }

    private static int? GetKerningHalfPoints(WordParagraph wordRun) {
        uint? value = WordXml.ParseUInt32(DirectRunProperties(wordRun).WChild("kern").WVal());
        if (!value.HasValue || value.Value == 0 || value.Value > int.MaxValue) {
            return null;
        }

        return (int)value.Value;
    }

    private static int? GetCharacterOffsetHalfPoints(WordParagraph wordRun) {
        string? value = DirectRunProperties(wordRun).WChild("position").WVal();
        if (!int.TryParse(value, out int offset) || offset == 0) {
            return null;
        }

        return offset;
    }

    private static void SetCharacterScale(WordParagraph wordRun, int percent) {
        if (wordRun.RunNode().IsNone) {
            return;
        }

        WordXmlEdits.ReplaceChild(GetOrAddRunProperties(wordRun), "w", WRPrView.Order, FormatInt(percent)!);
    }

    private static void SetKerning(WordParagraph wordRun, int halfPoints) {
        if (wordRun.RunNode().IsNone || halfPoints < 0) {
            return;
        }

        WordXmlEdits.ReplaceChild(GetOrAddRunProperties(wordRun), "kern", WRPrView.Order, FormatInt(halfPoints)!);
    }

    private static void SetCharacterOffset(WordParagraph wordRun, int halfPoints) {
        if (wordRun.RunNode().IsNone) {
            return;
        }

        WordXmlEdits.ReplaceChild(GetOrAddRunProperties(wordRun), "position", WRPrView.Order, FormatInt(halfPoints)!);
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
