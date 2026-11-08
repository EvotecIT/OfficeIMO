using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

internal sealed partial class WordMl2003Projection {
    private static readonly Dictionary<string, HashSet<string>> PropertyChildren = new(StringComparer.Ordinal) {
        ["rPr"] = Names("rStyle rFonts b b-cs i i-cs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing w kern position sz sz-cs highlight u effect bdr shd fitText vertAlign rtl cs em lang specVanish oMath"),
        ["pPr"] = Names("pStyle keepNext keepLines pageBreakBefore widowControl listPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr"),
        ["tblPr"] = Names("tblStyle tblpPr tblOverlap bidiVisual tblStyleRowBandSize tblStyleColBandSize tblW jc tblCellSpacing tblInd tblBorders shd tblLayout tblCellMar tblLook tblCaption tblDescription"),
        ["tblPrEx"] = Names("tblW jc tblCellSpacing tblInd tblBorders shd tblLayout tblCellMar tblLook"),
        ["trPr"] = Names("cnfStyle divId gridBefore gridAfter wBefore wAfter cantSplit trHeight tblHeader tblCellSpacing jc hidden"),
        ["tcPr"] = Names("cnfStyle tcW gridSpan hmerge vmerge tcBorders shd noWrap tcMar textDirection tcFitText vAlign hideMark"),
        ["sectPr"] = Names("type pgSz pgMar paperSrc pgBorders lnNumType pgNumType cols formProt vAlign noEndnote titlePg textDirection bidi rtlGutter docGrid"),
        ["tblGrid"] = Names("gridCol"),
        ["tabs"] = Names("tab"),
        ["pBdr"] = Names("top left bottom right between bar"),
        ["tblBorders"] = Names("top left bottom right insideH insideV"),
        ["tcBorders"] = Names("top left bottom right insideH insideV tl2br tr2bl"),
        ["pgBorders"] = Names("top left bottom right"),
        ["tblCellMar"] = Names("top left bottom right"),
        ["tcMar"] = Names("top left bottom right"),
        ["cols"] = Names("col"),
        ["tblStylePr"] = Names("pPr rPr tblPr trPr tcPr"),
        ["listDef"] = Names("lsid plt tmpl name styleLink listStyleLink lvl"),
        ["list"] = Names("ilst lvlOverride"),
        ["lvlOverride"] = Names("startOverride lvl"),
        ["lvl"] = Names("start nfc lvlRestart pStyle isLgl suff lvlText lvlJc pPr rPr")
    };
    private static readonly Dictionary<string, HashSet<string>> PropertyAttributes = new(StringComparer.Ordinal) {
        ["rFonts"] = Names("ascii h-ansi fareast cs hint"),
        ["spacing"] = Names("before before-lines before-autospacing after after-lines after-autospacing line line-rule"),
        ["ind"] = Names("left left-chars right right-chars hanging hanging-chars first-line first-line-chars"),
        ["lang"] = Names("val fareast bidi"),
        ["color"] = Names("val"),
        ["u"] = Names("val color"),
        ["bdr"] = Names("val sz space color shadow frame"),
        ["shd"] = Names("val color fill"),
        ["fitText"] = Names("val id"),
        ["tab"] = Names("val leader pos"),
        ["cnfStyle"] = Names("val"),
        ["tblW"] = Names("w type"),
        ["tblInd"] = Names("w type"),
        ["tblCellSpacing"] = Names("w type"),
        ["wBefore"] = Names("w type"),
        ["wAfter"] = Names("w type"),
        ["tcW"] = Names("w type"),
        ["gridCol"] = Names("w"),
        ["trHeight"] = Names("val h-rule"),
        ["tblLayout"] = Names("type"),
        ["tblLook"] = Names("val firstRow lastRow firstColumn lastColumn noHBand noVBand"),
        ["tblpPr"] = Names("leftFromText rightFromText topFromText bottomFromText vertAnchor horzAnchor tblpXSpec tblpX tblpYSpec tblpY"),
        ["pgSz"] = Names("w h orient code"),
        ["pgMar"] = Names("top right bottom left header footer gutter"),
        ["paperSrc"] = Names("first other"),
        ["pgBorders"] = Names("zOrder display offsetFrom"),
        ["lnNumType"] = Names("count-by start distance restart"),
        ["pgNumType"] = Names("fmt start chap-style chap-sep"),
        ["cols"] = Names("equalWidth space num sep"),
        ["col"] = Names("w space"),
        ["docGrid"] = Names("type line-pitch char-space"),
        ["zoom"] = Names("val percent"),
        ["br"] = Names("type clear"),
        ["fldChar"] = Names("fldCharType fldLock dirty"),
        ["tblStylePr"] = Names("type"),
        ["listDef"] = Names("listDefId"),
        ["list"] = Names("ilfo"),
        ["lvlOverride"] = Names("ilvl"),
        ["lvl"] = Names("ilvl tplc tentative")
    };
    private static readonly Dictionary<string, string> ElementNames = new(StringComparer.Ordinal) {
        ["b-cs"] = "bCs", ["i-cs"] = "iCs", ["sz-cs"] = "szCs",
        ["hmerge"] = "hMerge", ["vmerge"] = "vMerge", ["panose-1"] = "panose1",
        ["listDef"] = "abstractNum", ["list"] = "num", ["ilst"] = "abstractNumId",
        ["lsid"] = "nsid", ["plt"] = "multiLevelType", ["listStyleLink"] = "numStyleLink", ["nfc"] = "numFmt",
        ["listPr"] = "numPr", ["ilfo"] = "numId"
    };
    private static readonly Dictionary<string, string> AttributeNames = new(StringComparer.Ordinal) {
        ["h-ansi"] = "hAnsi", ["fareast"] = "eastAsia", ["line-rule"] = "lineRule",
        ["before-lines"] = "beforeLines", ["before-autospacing"] = "beforeAutospacing",
        ["after-lines"] = "afterLines", ["after-autospacing"] = "afterAutospacing",
        ["left-chars"] = "leftChars", ["right-chars"] = "rightChars", ["hanging-chars"] = "hangingChars",
        ["first-line"] = "firstLine", ["first-line-chars"] = "firstLineChars",
        ["h-rule"] = "hRule", ["count-by"] = "countBy", ["chap-style"] = "chapStyle",
        ["chap-sep"] = "chapSep", ["line-pitch"] = "linePitch", ["char-space"] = "charSpace",
        ["listDefId"] = "abstractNumId", ["ilfo"] = "numId"
    };

    private static HashSet<string> Names(string words) => new(words.Split(' '), StringComparer.Ordinal);

    private XElement? MapProperties(XElement source, string context, string? parent = null) {
        _token.ThrowIfCancellationRequested();
        if (context == "listPr") return MapListProperties(source);
        string name = ElementNames.TryGetValue(context, out string? mappedName) ? mappedName : context;
        var output = new XElement(ModernWord + name);
        output.SetAttributeValue(XNamespace.Xmlns + "w", ModernWord.NamespaceName);
        PropertyAttributes.TryGetValue(context, out HashSet<string>? allowed);
        if (context == "spacing" && parent == "rPr") allowed = RunSpacingAttributes;
        else if (parent is "pBdr" or "tblBorders" or "tcBorders" or "pgBorders") allowed = BorderAttributes;
        else if (parent is "tblCellMar" or "tcMar") allowed = WidthAttributes;
        bool container = PropertyChildren.TryGetValue(context, out HashSet<string>? children);
        foreach (XAttribute attribute in source.Attributes()) {
            if (IsMetadataAttribute(attribute)) continue;
            string local = attribute.Name.LocalName;
            if (attribute.Name.Namespace != Word || !(allowed?.Contains(local) ?? (!container && local == "val"))) {
                Omit(attribute.Name, context + " attribute " + local);
                continue;
            }
            string target = AttributeNames.TryGetValue(local, out string? mapped) ? mapped : local;
            string sourceValue = IsHexPropertyAttribute(context, local)
                ? attribute.Value.Trim(' ', '\t', '\r', '\n') : attribute.Value;
            if (!TryMapPropertyValue(context, local, parent, sourceValue, out string value)) {
                string token = attribute.Value.Length <= 80 ? attribute.Value : attribute.Value.Substring(0, 80) + "...";
                Omit(attribute.Name, context + " property with unsupported " + local + " enum value '" + token + "'");
                return null;
            }
            output.SetAttributeValue(ModernWord + target, value);
        }
        foreach (XElement child in source.Elements()) {
            if (context == "pPr" && parent == "p" && child.Name == Word + "sectPr" ||
                context == "sectPr" && (child.Name == Word + "hdr" || child.Name == Word + "ftr")) continue;
            if (context == "rPr" && child.Name == Auxiliary + "font") continue;
            if (child.Name.Namespace != Word || children == null || !children.Contains(child.Name.LocalName)) {
                Omit(child, context + " formatting");
                continue;
            }
            output.Add(MapProperties(child, child.Name.LocalName, context));
        }
        if (context == "rPr" && source.Element(Auxiliary + "font") is { } font) {
            CheckAttributes(font, "val", Auxiliary);
            string? family = (string?)font.Attribute(Auxiliary + "val");
            if (!string.IsNullOrEmpty(family)) {
                XElement? fonts = output.Element(ModernWord + "rFonts");
                if (fonts == null) { fonts = new XElement(ModernWord + "rFonts"); output.AddFirst(fonts); }
                if (fonts.Attribute(ModernWord + "ascii") == null) fonts.SetAttributeValue(ModernWord + "ascii", family);
                if (fonts.Attribute(ModernWord + "hAnsi") == null) fonts.SetAttributeValue(ModernWord + "hAnsi", family);
            }
        }
        if (source.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value)))
            Omit(source, context + " text content");
        return output;
    }

    private T NormalizeMappedProperties<T>(T properties, string? mappedXml = null) where T : OpenXmlCompositeElement {
        OrderMappedProperties(properties);
        AdmitMappedProperties(properties, mappedXml);
        return properties;
    }

    private void OrderMappedProperties(OpenXmlCompositeElement properties) {
        // These particles contain choices or repeated values whose source order is meaningful.
        bool preserveOrder = properties is TableRowProperties or TableStyleConditionalFormattingTableRowProperties
            or TableGrid or Tabs or Columns;
        OpenXmlElement[] children = properties.ChildElements.ToArray();
        if (!preserveOrder) properties.RemoveAllChildren();
        foreach (OpenXmlElement child in children) {
            bool repeatedListChild = properties is AbstractNum && child is Level ||
                properties is NumberingInstance && child is LevelOverride;
            bool duplicate = !preserveOrder && !repeatedListChild && properties.ChildElements.Any(existing =>
                existing.LocalName == child.LocalName && existing.NamespaceUri == child.NamespaceUri);
            if (child is OpenXmlUnknownElement || duplicate || !preserveOrder && !repeatedListChild && !properties.AddChild(child, false)) {
                if (preserveOrder) child.Remove();
                string name = LegacyPropertyName(child.LocalName);
                Omit(Word + name, duplicate ? "duplicate " + properties.LocalName + " property"
                    : properties.LocalName + " property not supported in this formatting context");
                continue;
            }
            if (repeatedListChild) properties.AppendChild(child);
            if (child is OpenXmlCompositeElement nested) OrderMappedProperties(nested);
        }
    }

    private static bool IsHexPropertyAttribute(string context, string attribute) =>
        attribute is "color" or "fill" ||
        attribute == "val" && context is "color" or "tblLook" or "panose-1" or "charset" or "lsid" or "tmpl" ||
        context == "lvl" && attribute == "tplc";

    private static string LegacyPropertyName(string name) {
        foreach (KeyValuePair<string, string> mapping in ElementNames) {
            if (mapping.Value == name) return mapping.Key;
        }
        return name;
    }

    private static bool IsMetadataAttribute(XAttribute attribute) =>
        attribute.IsNamespaceDeclaration || attribute.Name == XNamespace.Xml + "space" ||
        attribute.Name == Vml + "ext" ||
        attribute.Name.NamespaceName == "http://schemas.microsoft.com/office/word/2003/wordml/sp2" &&
            attribute.Name.LocalName.StartsWith("rsid", StringComparison.Ordinal) ||
        attribute.Name.Namespace == Auxiliary && attribute.Name.LocalName == "bdrwidth";

    private void CheckAttributes(XElement source, string allowedNames, XNamespace? attributeNamespace = null) {
        XNamespace ns = attributeNamespace ?? Word;
        foreach (XAttribute attribute in source.Attributes()) {
            if (!IsMetadataAttribute(attribute) &&
                (attribute.Name.Namespace != ns || !ContainsName(allowedNames, attribute.Name.LocalName)))
                Omit(attribute.Name, source.Name.LocalName + " attribute " + attribute.Name.LocalName);
        }
    }

    private static readonly HashSet<string> BorderAttributes = Names("val sz space color shadow frame");
    private static readonly HashSet<string> WidthAttributes = Names("w type");
    private static readonly HashSet<string> RunSpacingAttributes = Names("val");

    private static bool ContainsName(string names, string name) {
        int offset = 0;
        while ((offset = names.IndexOf(name, offset, StringComparison.Ordinal)) >= 0) {
            int end = offset + name.Length;
            if ((offset == 0 || names[offset - 1] == ' ') && (end == names.Length || names[end] == ' ')) return true;
            offset = end;
        }
        return false;
    }
}
