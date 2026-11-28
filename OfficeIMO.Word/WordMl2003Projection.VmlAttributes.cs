using System.Xml.Linq;
using System.Globalization;
using System.Text.RegularExpressions;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Word;

internal sealed partial class WordMl2003Projection {
    private readonly HashSet<(string Story, XName Element, string Identifier)> _vmlIdentifiers = new();
    private const string LockBooleanAttributes = "aspectratio position selection grouping ungrouping rotation cropping verticies adjusthandles text shapetype";
    private static readonly Dictionary<string, string> VmlOfficeAttributes = new(StringComparer.Ordinal) {
        ["shape"] = "spid spt preferrelative allowincell allowoverlap",
        ["shapetype"] = "spt preferrelative",
        ["imagedata"] = "title detectmouseclick",
        ["path"] = "extrusionok connecttype connectlocs connectangles",
        ["fill"] = "opacity2"
    };
    private static readonly Dictionary<string, string> VmlBooleanAttributes = new(StringComparer.Ordinal) {
        ["shape"] = "filled stroked insetpen preferrelative allowincell allowoverlap",
        ["shapetype"] = "filled stroked preferrelative",
        ["imagedata"] = "grayscale bilevel detectmouseclick",
        ["stroke"] = "on",
        ["path"] = "extrusionok gradientshapeok fillok strokeok shadowok arrowok",
        ["fill"] = "on",
        ["lock"] = LockBooleanAttributes
    };

    private void AppendVmlAttributes(XElement source, XElement target) {
        foreach (XAttribute attribute in source.Attributes()) {
            bool locked = source.Name == Office + "lock";
            if (attribute.IsNamespaceDeclaration || locked && attribute.Name == XNamespace.Xml + "space") {
                target.Add(new XAttribute(attribute));
                continue;
            }
            // The legacy src is consumed by the embedded-resource owner, never copied into the target VML.
            if (source.Name == Vml + "imagedata" && attribute.Name == XName.Get("src")) continue;
            if (!IsSupportedVmlAttribute(source.Name, attribute.Name)) {
                Omit(attribute.Name, "VML picture attribute " + attribute.Name.LocalName);
                continue;
            }
            if (!IsSupportedVmlAttributeValue(source.Name, attribute.Name, attribute.Value)) {
                string token = attribute.Value.Length <= 80 ? attribute.Value : attribute.Value.Substring(0, 80) + "...";
                Omit(attribute.Name, "VML picture attribute " + attribute.Name.LocalName + " with unsupported value '" + token + "'");
                continue;
            }
            target.Add(new XAttribute(attribute));
        }
    }

    private static bool IsSupportedVmlAttribute(XName element, XName attribute) {
        if (element == Office + "lock")
            return attribute == Vml + "ext" || attribute.Namespace == XNamespace.None && ContainsName(LockBooleanAttributes, attribute.LocalName);
        string name = element.LocalName;
        if (!VmlAttributes.TryGetValue(name, out string? allowed) || !ContainsName(allowed, attribute.LocalName)) return false;
        bool office = VmlOfficeAttributes.TryGetValue(name, out string? qualified) && ContainsName(qualified, attribute.LocalName);
        return attribute.Namespace == (office ? Office : XNamespace.None);
    }

    private static bool IsSupportedVmlAttributeValue(XName element, XName attribute, string value) {
        string? tokens = (element.LocalName, attribute.LocalName) switch {
            ("stroke", "joinstyle") => "round bevel miter",
            ("stroke", "endcap") => "flat square round",
            ("path", "connecttype") => "none rect segments custom",
            ("fill", "type") => "solid gradient gradientRadial tile pattern frame",
            ("fill", "method") => "none linear sigma any",
            ("lock", "ext") => "view edit backwardCompatible",
            _ => null
        };
        if (tokens != null) return value == "linear sigma" && (element.LocalName, attribute.LocalName) == ("fill", "method")
            || ContainsName(tokens, value);
        if (attribute == Office + "spt" && (element == Vml + "shape" || element == Vml + "shapetype"))
            return OfficeXmlViewValues.ParseInt32(value) is >= 0 and <= 202;
        if (element == Vml + "fill" && attribute == XName.Get("angle"))
            return decimal.TryParse(value, NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint |
                NumberStyles.AllowLeadingWhite | NumberStyles.AllowTrailingWhite, CultureInfo.InvariantCulture, out decimal angle)
                && angle is >= -32767M and <= 32767M;
        if (element == Vml + "fill" && attribute == XName.Get("focus"))
            return value.Length == 0 || value.Length <= 5 && Regex.IsMatch(value, @"^-?(\d{1,2}|100)%$", RegexOptions.CultureInvariant);
        if (element == Vml + "imagedata" && attribute == XName.Get("blacklevel"))
            return IsVmlNumberInRange(value, -0.5, 0.5);
        if (element == Vml + "stroke" && attribute == XName.Get("weight"))
            return IsVmlNumberInRange(value, 0, 20116800);
        return !VmlBooleanAttributes.TryGetValue(element.LocalName, out string? booleans) ||
            !ContainsName(booleans, attribute.LocalName) || value is "t" or "f" or "true" or "false";
    }

    private static bool IsVmlNumberInRange(string value, double minimum, double maximum) {
        // Signed plain numbers obey the declared range; native unit/fixed-point strings stay literal.
        return !OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(value, NumberStyles.AllowDecimalPoint | NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out double number)
            || number >= minimum && number <= maximum;
    }

    private void CheckVmlIdentifiers(XElement picture) {
        foreach (XElement element in picture.Elements()) {
            string? identifier = (string?)element.Attribute("id");
            if (identifier != null && identifier.Length != 0 && !_vmlIdentifiers.Add((_story.PartName, element.Name, identifier)))
                throw new InvalidDataException("WordML 2003 contains a duplicate VML " + element.Name.LocalName + " identifier within one story.");
        }
    }
}
