using System.Xml.Linq;
using V = DocumentFormat.OpenXml.Vml;
using O = DocumentFormat.OpenXml.Vml.Office;

namespace OfficeIMO.Word;

internal sealed partial class WordMl2003Projection {
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
            if (attribute.IsNamespaceDeclaration || locked && attribute.Name != Vml + "ext" && IsMetadataAttribute(attribute)) {
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
        Func<string, bool>? valid = (element.LocalName, attribute.LocalName) switch {
            ("stroke", "joinstyle") => IsEnum<V.StrokeJoinStyleValues>,
            ("stroke", "endcap") => IsEnum<V.StrokeEndCapValues>,
            ("path", "connecttype") => IsEnum<O.ConnectValues>,
            ("fill", "type") => IsEnum<V.FillTypeValues>,
            ("fill", "method") => IsEnum<V.FillMethodValues>,
            ("lock", "ext") => IsEnum<V.ExtensionHandlingBehaviorValues>,
            _ => null
        };
        if (valid != null) return valid(value);
        if (!VmlBooleanAttributes.TryGetValue(element.LocalName, out string? booleans) || !ContainsName(booleans, attribute.LocalName)) return true;
        try {
            // Assigning InnerText alone defers parsing. The SDK Value getter owns the TrueFalse vocabulary.
            _ = new TrueFalseValue { InnerText = value }.Value;
            return true;
        } catch (FormatException) {
            return false;
        }
    }
}
