using System.Globalization;
using System.Xml.Linq;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Word;

internal sealed partial class WordMl2003Projection {
    private const int MaximumPropertyErrors = 1000;
    private const string InvalidScalar = "Sch_AttributeValueDataTypeDetailed";
    private const string InvalidUnion = "Sch_AttributeUnionFailedEx";
    private const string MissingAttribute = "Sch_MissRequiredAttribute";

    private void AdmitMappedProperties(XElement properties) {
        List<XElement>? invalid = null;
        int errorCount = 0;
        foreach (XElement node in properties.Descendants().Reverse()) {
            _token.ThrowIfCancellationRequested();
            int elementErrors = CountPropertyErrors(node, out string? error);
            if (elementErrors == 0) continue;
            errorCount += elementErrors;
            if (errorCount >= MaximumPropertyErrors)
                throw new InvalidDataException("WordML 2003 property group exceeds the property validation error limit.");
            (invalid ??= new()).Add(node);

            string name = LegacyPropertyName(node.Name.LocalName);
            string attributes = string.Join(", ", OrderDiagnosticAttributes(node).Select(DescribeMappedAttribute));
            Omit(Word + name, name + " property with unsupported attribute values [" + attributes + "] (" + error + ")");
        }
        if (invalid == null) return;
        foreach (XElement node in invalid) {
            if (node.Parent == null) continue;
            XElement? parent = node.Parent;
            node.Remove();
            while (parent != null && parent != properties && !parent.HasElements && !parent.HasAttributes) {
                XElement? next = parent.Parent;
                if (next == null) break;
                parent.Remove();
                parent = next;
            }
        }
    }

    private static int CountPropertyErrors(XElement node, out string? firstError) {
        firstError = null;
        int count = 0;
        string name = node.Name.LocalName;
        string? required = name switch {
            "tab" => "val pos",
            "lvl" or "lvlOverride" => "ilvl",
            "fitText" or "trHeight" => null,
            "bdr" or "top" or "left" or "bottom" or "right" or "between" or "bar" or "insideH" or "insideV" or "tl2br" or "tr2bl"
                when node.Parent?.Name.LocalName is not ("tblCellMar" or "tcMar") => "val",
            "color" or "shd" or "sz" or "szCs" or "w" or "kern" or "position" or "highlight" or "effect" or "vertAlign" or "em" or
            "pStyle" or "rStyle" or "tblStyle" or "cnfStyle" or "gridSpan" or "gridBefore" or "gridAfter" or "jc" or "textDirection" or
            "textAlignment" or "outlineLvl" or "divId" or "vAlign" or "ilvl" or "numId" or "abstractNumId" or "numFmt" or "start" or
            "startOverride" or "lvlRestart" or "lvlJc" or "suff" or "multiLevelType" or "nsid" or "tmpl" or "panose1" or "charset" or
            "family" or "pitch" or "defaultTabStop" or "hyphenationZone" or "consecutiveHyphenLimit" or "characterSpacingControl" or "view" => "val",
            "zoom" => "percent",
            _ => null
        };
        if (required != null) {
            if (required == "val") {
                if (node.Attribute(ModernWord + "val") == null) count++;
            } else {
                foreach (string attribute in required.Split(' ')) {
                    if (node.Attribute(ModernWord + attribute) == null) count++;
                }
            }
        }
        if (count != 0) firstError = MissingAttribute;
        foreach (XAttribute attribute in node.Attributes()) {
            string? error = GetScalarError(name, node.Parent?.Name.LocalName, attribute.Name.LocalName, attribute.Value);
            if (error == null) continue;
            firstError ??= error;
            count++;
        }
        return count;
    }

    private static string? GetScalarError(string element, string? parent, string attribute, string value) {
        if (attribute is "color" or "fill" || (element, attribute) == ("color", "val"))
            return value == "auto" || IsHex(value, 6) ? null : InvalidUnion;

        return (element, attribute) switch {
            ("panose1", "val") => IsHex(value, 20) ? null : InvalidScalar,
            ("charset", "val") => IsHex(value, 2) ? null : InvalidScalar,
            ("nsid" or "tmpl", "val") or ("lvl", "tplc") => IsHex(value, 8) ? null : InvalidScalar,
            ("tblLook", "val") => IsHex(value, 4) ? null : InvalidScalar,
            ("cnfStyle", "val") => value.Length == 12 && value.All(character => character is '0' or '1') ? null : InvalidScalar,
            ("pgSz", "w" or "h") => InRange(value, 0, 31680),
            ("cols", "num") => InRange(value, 1, 45),
            ("tab", "pos") => InRange(value, -31680, 31680),
            ("outlineLvl", "val") => InRange(value, 0, 9),
            ("ilvl", "val") or ("lvl" or "lvlOverride", "ilvl") => InRange(value, 0, 8),
            ("w", "val") => InRange(value, 1, 600),
            ("trHeight", "val") or ("gridCol", "w") or ("pgMar", "left" or "right" or "header" or "footer" or "gutter")
                or ("cols", "space") or ("col", "w" or "space") or ("hyphenationZone" or "kern" or "sz" or "szCs", "val")
                or ("spacing", "before" or "after") or ("ind", "hanging" or "firstLine") => Measure(value, signed: false),
            ("pgMar", "top" or "bottom") or ("fitText", "val") or ("spacing", "line")
                or ("ind", "left" or "right" or "start" or "end") or ("position" or "spacing", "val")
                or ("tblpPr", "leftFromText" or "rightFromText" or "topFromText" or "bottomFromText" or "tblpX" or "tblpY") => Measure(value, signed: true),
            ("tblW" or "tblInd" or "tblCellSpacing" or "tcW" or "wBefore" or "wAfter", "w") => PercentageOrMeasurement(value),
            (_, "w") when parent is "tblCellMar" or "tcMar" => PercentageOrMeasurement(value),
            (_, "sz" or "space") when element == "bdr" || parent is "pBdr" or "tblBorders" or "tcBorders" or "pgBorders"
                => OfficeXmlViewValues.ParseUInt32(value).HasValue ? null : InvalidScalar,
            ("spacing", "beforeLines" or "afterLines") or ("ind", "startChars" or "endChars" or "hangingChars" or "firstLineChars")
                or ("pgSz", "code") or ("paperSrc", "first" or "other") or ("pgNumType", "start" or "chapStyle")
                or ("lnNumType", "countBy" or "start") or ("docGrid", "linePitch" or "charSpace") or ("fitText", "id")
                or ("gridSpan" or "gridBefore" or "gridAfter" or "numId" or "abstractNumId" or "start" or "startOverride" or
                    "lvlRestart" or "divId" or "tblStyleRowBandSize" or "tblStyleColBandSize" or "consecutiveHyphenLimit", "val")
                => OfficeXmlViewValues.ParseInt32(value).HasValue ? null : InvalidScalar,
            ("lnNumType", "distance") => Measure(value, signed: false),
            ("defaultTabStop", "val") => OfficeXmlViewValues.ParseUInt32(value).HasValue ? null : InvalidScalar,
            ("zoom", "percent") => IsPercentage(value) || OfficeXmlViewValues.ParseInt32(value).HasValue ? null : InvalidUnion,
            _ => null
        };
    }

    private static string? InRange(string value, int min, int max) =>
        OfficeXmlViewValues.ParseInt32(value) is int parsed && parsed >= min && parsed <= max ? null : InvalidScalar;

    private static string? Measure(string value, bool signed) {
        if (signed ? OfficeXmlViewValues.ParseInt32(value).HasValue : OfficeXmlViewValues.ParseUInt32(value).HasValue) return null;
        if (value.Length >= 3 && ContainsName("mm cm in pt pc pi", value.Substring(value.Length - 2)) &&
            TryDecimal(value.Substring(0, value.Length - 2), signed)) return null;
        return InvalidUnion;
    }

    private static string? PercentageOrMeasurement(string value) => IsPercentage(value) ? null : Measure(value, signed: true);

    private static bool IsPercentage(string value) =>
        value.Length > 1 && value[value.Length - 1] == '%' && TryDecimal(value.Substring(0, value.Length - 1), signed: true);

    private static bool TryDecimal(string value, bool signed) =>
        decimal.TryParse(value, NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out decimal parsed) &&
        (signed || parsed >= 0);

    private static bool IsHex(string value, int length) =>
        value.Length == length && value.All(character => character is >= '0' and <= '9' or >= 'a' and <= 'f' or >= 'A' and <= 'F');

    private static IEnumerable<XAttribute> OrderDiagnosticAttributes(XElement node) {
        string? order = node.Name.LocalName switch {
            "spacing" => "before beforeLines beforeAutospacing after afterLines afterAutospacing line lineRule",
            "tab" => "val leader pos",
            "ind" => "start startChars end endChars hanging hangingChars firstLine firstLineChars",
            "pgSz" => "w h orient code",
            _ => null
        };
        if (order == null) return node.Attributes();
        return node.Attributes().OrderBy(attribute => order.IndexOf(attribute.Name.LocalName, StringComparison.Ordinal));
    }

    private static string DescribeMappedAttribute(XAttribute attribute) {
        string name = attribute.Name.LocalName;
        foreach (KeyValuePair<string, string> mapping in AttributeNames) {
            if (mapping.Value == name) { name = mapping.Key; break; }
        }
        string value = attribute.Value;
        if (value.Length > 80) value = value.Substring(0, 80) + "...";
        return name + "='" + value + "'";
    }
}
