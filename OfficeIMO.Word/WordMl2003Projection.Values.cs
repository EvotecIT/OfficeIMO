using System.Text;

namespace OfficeIMO.Word;

internal sealed partial class WordMl2003Projection {
    private static readonly HashSet<string> OnOffProperties = Names(
        "b b-cs i i-cs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden " +
        "rtl cs specVanish oMath keepNext keepLines pageBreakBefore widowControl suppressLineNumbers suppressAutoHyphens " +
        "kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd contextualSpacing mirrorIndents " +
        "suppressOverlap bidiVisual cantSplit tblHeader hidden noWrap tcFitText hideMark formProt noEndnote titlePg rtlGutter " +
        "evenAndOddHeaders mirrorMargins displayBackgroundShape doNotHyphenateCaps autoHyphenation doNotEmbedSystemFonts punctuationKerning");

    // Word 2003 wordnet.xsd uses hyphen-separated tokens. Check the mapped token
    // against the OOXML value vocabulary for its destination attribute.
    private static bool TryMapPropertyValue(string context, string attribute, string? parent, string source, out string value) {
        if (attribute == "val" && context is "nfc" or "plt" or "suff")
            return TryMapListValue(context, source, out value);
        Func<string, bool>? valid = (context, attribute) switch {
            ("spacing", "line-rule") => static token => WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordLineSpacingRule _),
            ("trHeight", "h-rule") => static token => ContainsName("auto exact atLeast", token),
            ("u", "val") => static token => WordUnderlineStyleTokens.TryParse(token, out _),
            ("highlight", "val") => static token => WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordHighlightColor _),
            ("effect", "val") => static token => ContainsName("blinkBackground lights antsBlack antsRed shimmer sparkle none", token),
            ("bdr", "val") => static token => WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordBorderStyle _),
            ("shd", "val") => static token => WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordShadingPattern _),
            ("vertAlign", "val") => static token => WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordVerticalTextPosition _),
            ("em", "val") => static token => ContainsName("none dot comma circle underDot", token),
            ("tab", "val") => static token => WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordTabAlignment _),
            ("tab", "leader") => static token => WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordTabLeader _),
            ("jc", "val") when parent is "tblPr" or "tblPrEx" or "trPr" => static token => WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordTableAlignment _),
            ("jc", "val") => static token => WordParagraphAlignmentTokens.TryParse(token, out _),
            ("textDirection", "val") => static token => WordTextDirectionTokens.TryParse(token, out _),
            ("textAlignment", "val") => static token => WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordVerticalCharacterAlignment _),
            ("textboxTightWrap", "val") => static token => ContainsName("none allLines firstAndLastLine firstLineOnly lastLineOnly", token),
            ("hmerge" or "vmerge", "val") => static token => WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordCellMerge _),
            ("tblLayout", "type") => static token => ContainsName("fixed autofit", token),
            ("tblOverlap", "val") => static token => WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordTableOverlap _),
            ("tblStylePr", "type") => static token => ContainsName("wholeTable firstRow lastRow firstCol lastCol band1Vert band2Vert band1Horz band2Horz neCell nwCell seCell swCell", token),
            ("tblW" or "tblInd" or "tblCellSpacing" or "wBefore" or "wAfter" or "tcW", "type") => static token => WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordTableWidthUnit _),
            ("tblpPr", "vertAnchor") => static token => WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordTableVerticalAnchor _),
            ("tblpPr", "horzAnchor") => static token => WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordTableHorizontalAnchor _),
            ("tblpPr", "tblpXSpec") => static token => WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordTableHorizontalAlignment _),
            ("tblpPr", "tblpYSpec") => static token => WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordTableVerticalPositionAlignment _),
            ("type", "val") when parent == "sectPr" => static token => ContainsName("nextPage nextColumn continuous evenPage oddPage", token),
            ("pgSz", "orient") => static token => ContainsName("portrait landscape", token),
            ("pgBorders", "zOrder") => static token => ContainsName("front back", token),
            ("pgBorders", "display") => static token => ContainsName("allPages firstPage notFirstPage", token),
            ("pgBorders", "offsetFrom") => static token => ContainsName("page text", token),
            ("lnNumType", "restart") => static token => ContainsName("newPage newSection continuous", token),
            ("pgNumType", "fmt") => static token => WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordNumberFormat _),
            ("pgNumType", "chap-sep") => static token => ContainsName("hyphen period colon emDash enDash", token),
            ("vAlign", "val") when parent == "tcPr" => static token => ContainsName("top center bottom", token),
            ("vAlign", "val") => static token => ContainsName("top center both bottom", token),
            ("docGrid", "type") => static token => ContainsName("default lines linesAndChars snapToChars", token),
            ("br", "type") => static token => ContainsName("page column textWrapping", token),
            ("br", "clear") => static token => ContainsName("none left right all", token),
            ("fldChar", "fldCharType") => static token => ContainsName("begin separate end", token),
            ("rFonts", "hint") => static token => ContainsName("default eastAsia cs", token),
            ("family", "val") => static token => ContainsName("decorative modern roman script swiss auto", token),
            ("pitch", "val") => static token => ContainsName("default fixed variable", token),
            ("view", "val") => static token => ContainsName("none print outline masterPages normal web", token),
            ("zoom", "val") => static token => WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordZoomPreset _),
            ("characterSpacingControl", "val") => static token => WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordCharacterSpacing _),
            ("lvlJc", "val") => static token => WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordListLevelAlignment _),
            _ when attribute == "val" && parent is "pBdr" or "tblBorders" or "tcBorders" or "pgBorders" => static token => WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordBorderStyle _),
            _ when attribute == "type" && parent is "tblCellMar" or "tcMar" => static token => WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordTableWidthUnit _),
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
