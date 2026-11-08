using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

internal sealed partial class WordMl2003Projection {
    private static readonly HashSet<string> OnOffProperties = Names(
        "b b-cs i i-cs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden " +
        "rtl cs specVanish oMath keepNext keepLines pageBreakBefore widowControl suppressLineNumbers suppressAutoHyphens " +
        "kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd contextualSpacing mirrorIndents " +
        "suppressOverlap bidiVisual cantSplit tblHeader hidden noWrap tcFitText hideMark formProt noEndnote titlePg rtlGutter " +
        "evenAndOddHeaders mirrorMargins displayBackgroundShape doNotHyphenateCaps autoHyphenation doNotEmbedSystemFonts punctuationKerning");

    // Word 2003 wordnet.xsd uses hyphen-separated tokens. The target SDK owns the
    // accepted value vocabulary; a token without a qualified target is a reported loss.
    private static bool TryMapPropertyValue(string context, string attribute, string? parent, string source, out string value) {
        if (attribute == "val" && context is "nfc" or "plt" or "suff")
            return TryMapListValue(context, source, out value);
        Func<string, bool>? valid = (context, attribute) switch {
            ("spacing", "line-rule") => IsEnum<LineSpacingRuleValues>,
            ("trHeight", "h-rule") => IsEnum<HeightRuleValues>,
            ("u", "val") => IsEnum<UnderlineValues>,
            ("highlight", "val") => IsEnum<HighlightColorValues>,
            ("effect", "val") => IsEnum<TextEffectValues>,
            ("bdr", "val") => IsEnum<BorderValues>,
            ("shd", "val") => IsEnum<ShadingPatternValues>,
            ("vertAlign", "val") => IsEnum<VerticalPositionValues>,
            ("em", "val") => IsEnum<EmphasisMarkValues>,
            ("tab", "val") => IsEnum<TabStopValues>,
            ("tab", "leader") => IsEnum<TabStopLeaderCharValues>,
            ("jc", "val") when parent is "tblPr" or "tblPrEx" or "trPr" => IsEnum<TableRowAlignmentValues>,
            ("jc", "val") => IsEnum<JustificationValues>,
            ("textDirection", "val") => IsEnum<TextDirectionValues>,
            ("textAlignment", "val") => IsEnum<VerticalTextAlignmentValues>,
            ("textboxTightWrap", "val") => IsEnum<TextBoxTightWrapValues>,
            ("hmerge" or "vmerge", "val") => IsEnum<MergedCellValues>,
            ("tblLayout", "type") => IsEnum<TableLayoutValues>,
            ("tblOverlap", "val") => IsEnum<TableOverlapValues>,
            ("tblStylePr", "type") => IsEnum<TableStyleOverrideValues>,
            ("tblW" or "tblInd" or "tblCellSpacing" or "wBefore" or "wAfter" or "tcW", "type") => IsEnum<TableWidthUnitValues>,
            ("tblpPr", "vertAnchor") => IsEnum<VerticalAnchorValues>,
            ("tblpPr", "horzAnchor") => IsEnum<HorizontalAnchorValues>,
            ("tblpPr", "tblpXSpec") => IsEnum<HorizontalAlignmentValues>,
            ("tblpPr", "tblpYSpec") => IsEnum<VerticalAlignmentValues>,
            ("type", "val") when parent == "sectPr" => IsEnum<SectionMarkValues>,
            ("pgSz", "orient") => IsEnum<PageOrientationValues>,
            ("pgBorders", "zOrder") => IsEnum<PageBorderZOrderValues>,
            ("pgBorders", "display") => IsEnum<PageBorderDisplayValues>,
            ("pgBorders", "offsetFrom") => IsEnum<PageBorderOffsetValues>,
            ("lnNumType", "restart") => IsEnum<LineNumberRestartValues>,
            ("pgNumType", "fmt") => IsEnum<NumberFormatValues>,
            ("pgNumType", "chap-sep") => IsEnum<ChapterSeparatorValues>,
            ("vAlign", "val") when parent == "tcPr" => IsEnum<TableVerticalAlignmentValues>,
            ("vAlign", "val") => IsEnum<VerticalJustificationValues>,
            ("docGrid", "type") => IsEnum<DocGridValues>,
            ("br", "type") => IsEnum<BreakValues>,
            ("br", "clear") => IsEnum<BreakTextRestartLocationValues>,
            ("fldChar", "fldCharType") => IsEnum<FieldCharValues>,
            ("rFonts", "hint") => IsEnum<FontTypeHintValues>,
            ("family", "val") => IsEnum<FontFamilyValues>,
            ("pitch", "val") => IsEnum<FontPitchValues>,
            ("view", "val") => IsEnum<ViewValues>,
            ("zoom", "val") => IsEnum<PresetZoomValues>,
            ("characterSpacingControl", "val") => IsEnum<CharacterSpacingValues>,
            ("lvlJc", "val") => IsEnum<LevelJustificationValues>,
            _ when attribute == "val" && parent is "pBdr" or "tblBorders" or "tcBorders" or "pgBorders" => IsEnum<BorderValues>,
            _ when attribute == "type" && parent is "tblCellMar" or "tcMar" => IsEnum<TableWidthUnitValues>,
            _ => null
        };
        value = source;
        if (valid != null) {
            value = NormalizeLegacyEnum(source);
            if (context is "tblLayout" or "tblOverlap" or "family" or "pitch") value = value.ToLowerInvariant();
            if (context == "rFonts" && source == "fareast") value = "eastAsia";
            if (context == "tab" && attribute == "val" && source == "list") value = "num";
            if (context == "characterSpacingControl") value = source switch {
                "DontCompress" => "doNotCompress", "CompressPunctuation" => "compressPunctuation",
                "CompressPunctuationAndJapaneseKana" => "compressPunctuationAndJapaneseKana", _ => value
            };
            return valid(value);
        }
        bool onOff = attribute == "val" && OnOffProperties.Contains(context) ||
            attribute is "before-autospacing" or "after-autospacing" && context == "spacing" ||
            context == "cols" && attribute is "equalWidth" or "sep" ||
            context == "tblLook" && attribute is "firstRow" or "lastRow" or "firstColumn" or "lastColumn" or "noHBand" or "noVBand" ||
            context == "fldChar" && attribute is "fldLock" or "dirty" ||
            context == "lvl" && attribute == "tentative" ||
            context == "isLgl" && attribute == "val" ||
            attribute is "shadow" or "frame" && (context == "bdr" || parent is "pBdr" or "tblBorders" or "tcBorders" or "pgBorders");
        return !onOff || source is "on" or "off" or "true" or "false" or "1" or "0";
    }

    private static bool IsEnum<T>(string value) where T : struct, IEnumValue, IEnumValueFactory<T> =>
        default(T).Create(value).IsValid;

    private static string NormalizeLegacyEnum(string value) {
        int hyphen = value.IndexOf('-');
        if (hyphen < 0) return value;
        if (hyphen == 0 || hyphen == value.Length - 1 || value.IndexOf("--", StringComparison.Ordinal) >= 0) return value;
        var result = new StringBuilder(value.Length);
        result.Append(value, 0, hyphen);
        for (int index = hyphen; index < value.Length; index++) {
            if (value[index] == '-' && index + 1 < value.Length) result.Append(char.ToUpperInvariant(value[++index]));
            else result.Append(value[index]);
        }
        return result.ToString();
    }
}
