using System.Globalization;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

/// <summary>
/// Reads the few style facts that read-only passes need (identifier, name, type, default flag, base style,
/// paragraph numbering and caps toggles) straight from a styles part that has not been loaded, without building
/// its DOM. Word-authored styles parts are often hundreds of kilobytes of latent and table styles that text
/// extraction never touches. When the part is already loaded, or the XML holds anything the Open XML SDK might
/// read differently, callers use the loaded <see cref="Styles"/> element instead.
/// </summary>
internal static class WordStyleDefinitionsReader {
    private const string WordprocessingNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>One top-level <c>w:style</c>; each value mirrors the first matching child, as the SDK's typed accessors read it.</summary>
    internal sealed class StyleDefinition {
        internal StyleDefinition(string? styleId, string? type, bool isDefault) {
            StyleId = styleId;
            Type = type;
            IsDefault = isDefault;
        }

        internal string? StyleId { get; }
        internal string? Type { get; }
        internal bool IsDefault { get; }
        internal string? Name { get; set; }
        internal string? BasedOn { get; set; }
        internal int? NumberId { get; set; }
        internal int? Level { get; set; }
        internal bool HasRunProperties { get; set; }
        internal bool? Caps { get; set; }
        internal bool? SmallCaps { get; set; }
        internal bool IsParagraphStyle => string.Equals(Type, "paragraph", StringComparison.Ordinal);
    }

    /// <summary>The top-level style definitions of a styles part in document order, with its default run caps.</summary>
    internal sealed class StyleDefinitionSet {
        internal StyleDefinitionSet(List<StyleDefinition> styles) => Styles = styles;

        internal IReadOnlyList<StyleDefinition> Styles { get; }
        internal bool? DefaultCaps { get; set; }
        internal bool? DefaultSmallCaps { get; set; }
    }

    /// <summary>
    /// Returns the style definitions of <paramref name="part"/>, or null when the part is already loaded (its DOM
    /// may carry unsaved edits) or cannot be read with the same meaning as the DOM.
    /// </summary>
    internal static StyleDefinitionSet? TryRead(StylesPart? part) {
        if (part == null || part.IsRootElementLoaded) return null;
        try {
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                IgnoreComments = true,
                IgnoreProcessingInstructions = true,
                IgnoreWhitespace = true,
                CloseInput = false
            });
            if (reader.MoveToContent() != XmlNodeType.Element || !IsWord(reader, "styles")) return null;
            var styles = new List<StyleDefinition>();
            var set = new StyleDefinitionSet(styles);
            bool valid = true;
            bool seenDefaults = false;
            ReadChildren(reader, () => {
                if (!valid) return false;
                if (IsWord(reader, "style")) {
                    StyleDefinition? style = ReadStyle(reader);
                    if (style == null) valid = false;
                    else styles.Add(style);
                    return style != null;
                }
                if (!seenDefaults && IsWord(reader, "docDefaults")) {
                    seenDefaults = true;
                    valid = ReadDocumentDefaults(reader, set);
                    return true;
                }
                return false;
            });
            return valid ? set : null;
        } catch (XmlException) {
            return null;
        }
    }

    private static StyleDefinition? ReadStyle(XmlReader reader) {
        string? type = reader.GetAttribute("type", WordprocessingNamespace);
        if (type != null && type is not ("paragraph" or "character" or "table" or "numbering")) return null;
        if (!TryReadOnOff(reader.GetAttribute("default", WordprocessingNamespace), out bool isDefault)) return null;
        var style = new StyleDefinition(reader.GetAttribute("styleId", WordprocessingNamespace), type, isDefault);
        bool valid = true, seenName = false, seenBasedOn = false, seenParagraph = false;
        ReadChildren(reader, () => {
            if (reader.NamespaceURI != WordprocessingNamespace) return false;
            switch (reader.LocalName) {
                case "name" when !seenName:
                    seenName = true;
                    style.Name = reader.GetAttribute("val", WordprocessingNamespace);
                    return false;
                case "basedOn" when !seenBasedOn:
                    seenBasedOn = true;
                    style.BasedOn = reader.GetAttribute("val", WordprocessingNamespace);
                    return false;
                case "pPr" when !seenParagraph:
                    seenParagraph = true;
                    valid &= ReadParagraphProperties(reader, style);
                    return true;
                case "rPr" when !style.HasRunProperties:
                    style.HasRunProperties = true;
                    valid &= ReadCaps(reader, out bool? caps, out bool? smallCaps);
                    style.Caps = caps;
                    style.SmallCaps = smallCaps;
                    return true;
                default:
                    return false;
            }
        });
        return valid ? style : null;
    }

    // w:pPr/w:numPr/(w:numId, w:ilvl), first of each.
    private static bool ReadParagraphProperties(XmlReader reader, StyleDefinition style) {
        bool valid = true, seenNumbering = false;
        ReadChildren(reader, () => {
            if (seenNumbering || !IsWord(reader, "numPr")) return false;
            seenNumbering = true;
            bool seenId = false, seenLevel = false;
            ReadChildren(reader, () => {
                if (reader.NamespaceURI != WordprocessingNamespace) return false;
                if (reader.LocalName == "numId" && !seenId) {
                    seenId = true;
                    valid &= TryReadInt32(reader.GetAttribute("val", WordprocessingNamespace), out int? value);
                    style.NumberId = value;
                } else if (reader.LocalName == "ilvl" && !seenLevel) {
                    seenLevel = true;
                    valid &= TryReadInt32(reader.GetAttribute("val", WordprocessingNamespace), out int? value);
                    style.Level = value;
                }
                return false;
            });
            return true;
        });
        return valid;
    }

    // w:docDefaults/w:rPrDefault/w:rPr, first of each, as Styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle reads it.
    private static bool ReadDocumentDefaults(XmlReader reader, StyleDefinitionSet set) {
        bool valid = true, seenDefault = false;
        ReadChildren(reader, () => {
            if (seenDefault || !IsWord(reader, "rPrDefault")) return false;
            seenDefault = true;
            bool seenProperties = false;
            ReadChildren(reader, () => {
                if (seenProperties || !IsWord(reader, "rPr")) return false;
                seenProperties = true;
                valid &= ReadCaps(reader, out bool? caps, out bool? smallCaps);
                set.DefaultCaps = caps;
                set.DefaultSmallCaps = smallCaps;
                return true;
            });
            return true;
        });
        return valid;
    }

    // First w:caps and w:smallCaps children; a toggle present without w:val is on.
    private static bool ReadCaps(XmlReader reader, out bool? caps, out bool? smallCaps) {
        bool? capsValue = null, smallCapsValue = null;
        bool valid = true, seenCaps = false, seenSmallCaps = false;
        ReadChildren(reader, () => {
            if (reader.NamespaceURI != WordprocessingNamespace) return false;
            if (reader.LocalName == "caps" && !seenCaps) {
                seenCaps = true;
                valid &= TryReadToggle(reader.GetAttribute("val", WordprocessingNamespace), out capsValue);
            } else if (reader.LocalName == "smallCaps" && !seenSmallCaps) {
                seenSmallCaps = true;
                valid &= TryReadToggle(reader.GetAttribute("val", WordprocessingNamespace), out smallCapsValue);
            }
            return false;
        });
        caps = capsValue;
        smallCaps = smallCapsValue;
        return valid;
    }

    /// <summary>
    /// Visits the direct child elements of the current element and leaves the reader after its end. The visitor
    /// returns true when it consumed the child itself (through <see cref="ReadChildren"/>); otherwise the child is skipped.
    /// </summary>
    private static void ReadChildren(XmlReader reader, Func<bool> visitChild) {
        if (reader.IsEmptyElement) {
            reader.Read();
            return;
        }
        int depth = reader.Depth;
        reader.Read();
        while (!reader.EOF && reader.Depth > depth) {
            if (reader.NodeType != XmlNodeType.Element) {
                reader.Read();
                continue;
            }
            if (!visitChild()) reader.Skip();
        }
        if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth) reader.Read();
    }

    private static bool IsWord(XmlReader reader, string localName) =>
        reader.NamespaceURI == WordprocessingNamespace && reader.LocalName == localName;

    // Only plain values are read here; anything the SDK might normalize differently falls back to the DOM.
    private static bool TryReadInt32(string? text, out int? value) {
        value = null;
        if (text == null) return true;
        if (text.Length == 0 || text.Length > 11) return false;
        int start = text[0] == '-' ? 1 : 0;
        if (start == text.Length) return false;
        for (int index = start; index < text.Length; index++) {
            if (text[index] < '0' || text[index] > '9') return false;
        }
        if (!int.TryParse(text, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out int parsed)) return false;
        value = parsed;
        return true;
    }

    private static bool TryReadOnOff(string? text, out bool value) {
        value = false;
        if (text == null) return true;
        if (!TryReadToggle(text, out bool? toggle)) return false;
        value = toggle == true;
        return true;
    }

    // A present toggle element without w:val is on.
    private static bool TryReadToggle(string? text, out bool? value) {
        switch (text) {
            case null:
            case "1":
            case "true":
            case "on":
                value = true;
                return true;
            case "0":
            case "false":
            case "off":
                value = false;
                return true;
            default:
                value = null;
                return false;
        }
    }
}
