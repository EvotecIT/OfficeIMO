using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Word;

/// <summary>
/// Reads the style definitions that read-only passes need straight from a styles part that has not been loaded,
/// without building its DOM. One pass reads the facts readers look up for every style (identifier, name, type,
/// default flag, base, linked and next styles, paragraph numbering and caps toggles) and the document defaults.
/// A style's properties are parsed into a detached <see cref="Style"/> only when a caller asks for them. Word-authored
/// styles parts are often hundreds of kilobytes of latent and table styles that a reader never touches. When the
/// part is already loaded, or the XML holds anything the Open XML SDK might read differently, callers use the
/// loaded <see cref="Styles"/> element instead.
/// </summary>
internal static class WordStyleDefinitionsReader {
    private const string WordprocessingNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>
    /// One top-level <c>w:style</c>. Each fact mirrors the first matching child, as the SDK's typed accessors read
    /// it. <see cref="Element"/> is the loaded element itself, or a detached copy parsed from the part on first use.
    /// </summary>
    internal sealed class StyleDefinition {
        private readonly StyleDefinitionSet? _set;
        private Style? _element;

        internal StyleDefinition(StyleDefinitionSet set, string? styleId, string? type, bool isDefault) {
            _set = set;
            StyleId = styleId;
            Type = type;
            IsDefault = isDefault;
        }

        /// <summary>Reads the facts of a loaded style; <see cref="Element"/> is that element.</summary>
        internal StyleDefinition(Style style) {
            _element = style;
            StyleId = style.StyleId?.Value;
            Type = style.Type?.InnerText;
            IsDefault = style.Default is { HasValue: true } isDefault && isDefault.Value;
            Name = style.StyleName?.Val?.Value;
            BasedOn = style.BasedOn?.Val?.Value;
            LinkedStyle = style.LinkedStyle?.Val?.Value;
            NextStyle = style.NextParagraphStyle?.Val?.Value;
            IsCustom = style.CustomStyle is { HasValue: true } isCustom && isCustom.Value;
            IsPrimary = style.PrimaryStyle != null;
        }

        internal string? StyleId { get; }
        internal string? Type { get; }
        internal bool IsDefault { get; }
        internal string? Name { get; set; }
        internal string? BasedOn { get; set; }
        internal string? LinkedStyle { get; set; }
        internal string? NextStyle { get; set; }
        internal bool IsCustom { get; set; }
        internal bool IsPrimary { get; set; }
        internal int? NumberId { get; set; }
        internal int? Level { get; set; }
        internal bool HasRunProperties { get; set; }
        internal bool? Caps { get; set; }
        internal bool? SmallCaps { get; set; }
        internal bool IsParagraphStyle => string.Equals(Type, "paragraph", StringComparison.Ordinal);

        /// <summary>The style's markup, once <see cref="StyleDefinitionSet"/> has read it for a parse.</summary>
        internal string? Markup { get; set; }

        /// <summary>
        /// The style element. A streamed definition is parsed once, on first use, into a detached element that
        /// callers must treat as read-only.
        /// </summary>
        internal Style Element => Volatile.Read(ref _element) ?? _set!.Materialize(this);

        internal bool IsMaterialized => Volatile.Read(ref _element) != null;

        internal void SetElement(Style element) => Volatile.Write(ref _element, element);

        /// <summary>Whether this style applies to <paramref name="type"/>; a style without a type matches any.</summary>
        internal bool Matches(string type) => Type == null || string.Equals(Type, type, StringComparison.Ordinal);
    }

    /// <summary>The top-level style definitions of a styles part in document order, with its document defaults.</summary>
    internal sealed class StyleDefinitionSet {
        private readonly StylesPart _part;
        private readonly List<StyleDefinition> _styles = new();
        private readonly object _gate = new();
        private string? _documentDefaultsXml;
        private DocDefaults? _documentDefaults;
        private bool _markupRead;

        internal StyleDefinitionSet(StylesPart part) => _part = part;

        internal IReadOnlyList<StyleDefinition> Styles => _styles;
        internal bool? DefaultCaps { get; private set; }
        internal bool? DefaultSmallCaps { get; private set; }

        /// <summary>The first <c>w:docDefaults</c>, parsed on first use into a detached, read-only element.</summary>
        internal DocDefaults? DocumentDefaults {
            get {
                if (_documentDefaultsXml == null) return null;
                DocDefaults? defaults = Volatile.Read(ref _documentDefaults);
                if (defaults != null) return defaults;
                lock (_gate) {
                    return _documentDefaults ??= (DocDefaults)ParseDetached(new[] { _documentDefaultsXml })[0];
                }
            }
        }

        internal void Add(StyleDefinition style) => _styles.Add(style);

        internal void SetDocumentDefaults(string markup, bool? caps, bool? smallCaps) {
            _documentDefaultsXml = markup;
            DefaultCaps = caps;
            DefaultSmallCaps = smallCaps;
        }

        internal Style Materialize(StyleDefinition style) {
            lock (_gate) {
                if (style.IsMaterialized) return style.Element;
                ReadMarkup();
                var element = (Style)ParseDetached(new[] { style.Markup! })[0];
                style.SetElement(element);
                style.Markup = null;
                return element;
            }
        }

        /// <summary>
        /// Reads each style's markup with a second pass over the part, which still holds what the facts were read
        /// from: a later DOM load does not change the stored part until it is saved.
        /// </summary>
        private void ReadMarkup() {
            if (_markupRead) return;
            var settings = new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                CloseInput = false,
                IgnoreComments = true,
                IgnoreProcessingInstructions = true,
                IgnoreWhitespace = true
            };
            long maxCharacters = _part.OpenXmlPackage.MaxCharactersInPart;
            if (maxCharacters > 0) settings.MaxCharactersInDocument = maxCharacters;
            using Stream stream = _part.GetStream(FileMode.Open, FileAccess.Read);
            using XmlReader reader = XmlReader.Create(stream, settings);
            reader.MoveToContent();
            int index = 0;
            ReadChildren(reader, () => {
                if (!IsWord(reader, "style")) return false;
                if (index >= _styles.Count || !string.Equals(reader.GetAttribute("styleId", WordprocessingNamespace), _styles[index].StyleId, StringComparison.Ordinal)) {
                    throw new InvalidDataException("The styles part changed after its styles were read.");
                }
                StyleDefinition style = _styles[index++];
                string markup = reader.ReadOuterXml();
                if (!style.IsMaterialized) style.Markup = markup;
                return true;
            });
            if (index != _styles.Count) throw new InvalidDataException("The styles part changed after its styles were read.");
            _markupRead = true;
        }
    }

    [ThreadStatic] private static StyleDefinitionsPart? t_parsingPart;

    /// <summary>
    /// Parses top-level children of <c>w:styles</c> exactly as loading a styles part does, by loading them as the
    /// root of a small styles part in a private in-memory package and detaching them. Element constructors and
    /// <see cref="OpenXmlReader"/> would not do: they read some attributes, such as <c>w:ind/@w:left</c>, as their
    /// Office 2010 counterparts, while loading a part keeps them as authored. The elements are fully loaded, so
    /// concurrent readers of a read-only document never share a half-parsed element.
    /// </summary>
    private static IReadOnlyList<OpenXmlElement> ParseDetached(IReadOnlyList<string> markup) {
        StyleDefinitionsPart part = t_parsingPart ??= CreateParsingPart();
        try {
            using (Stream stream = part.GetStream(FileMode.Create, FileAccess.Write))
            using (var writer = new StreamWriter(stream, new UTF8Encoding(false))) {
                writer.Write("<w:styles xmlns:w=\"");
                writer.Write(WordprocessingNamespace);
                writer.Write("\">");
                foreach (string element in markup) writer.Write(element);
                writer.Write("</w:styles>");
            }
            List<OpenXmlElement> elements = part.Styles?.ChildElements.ToList() ?? new List<OpenXmlElement>();
            if (elements.Count != markup.Count) throw new InvalidDataException("A streamed style definition could not be parsed again.");
            foreach (OpenXmlElement element in elements) element.Remove();
            return elements;
        } finally {
            part.UnloadRootElement();
        }
    }

    private static StyleDefinitionsPart CreateParsingPart() {
        var package = WordprocessingDocument.Create(new MemoryStream(), WordprocessingDocumentType.Document);
        return package.AddMainDocumentPart().AddNewPart<StyleDefinitionsPart>();
    }

    /// <summary>
    /// Returns the style definitions of <paramref name="part"/>, or null when <see cref="CreateStreamingSettings"/>
    /// declines the part or its XML cannot be read with the same meaning as the DOM.
    /// </summary>
    internal static StyleDefinitionSet? TryRead(StylesPart? part) {
        XmlReaderSettings? settings = part == null ? null : CreateStreamingSettings(part);
        if (part == null || settings == null) return null;
        settings.IgnoreComments = true;
        settings.IgnoreProcessingInstructions = true;
        settings.IgnoreWhitespace = true;
        try {
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using XmlReader reader = XmlReader.Create(stream, settings);
            if (reader.MoveToContent() != XmlNodeType.Element || !IsWord(reader, "styles")) return null;
            var set = new StyleDefinitionSet(part);
            bool valid = true;
            bool seenDefaults = false;
            ReadChildren(reader, () => {
                if (!valid) return false;
                if (IsWord(reader, "style")) {
                    StyleDefinition? style = ReadStyle(reader, set);
                    if (style == null) valid = false;
                    else set.Add(style);
                    return style != null;
                }
                if (!seenDefaults && IsWord(reader, "docDefaults")) {
                    seenDefaults = true;
                    string markup = reader.ReadOuterXml();
                    valid = ReadDocumentDefaults(markup, settings, out bool? caps, out bool? smallCaps);
                    set.SetDocumentDefaults(markup, caps, smallCaps);
                    return true;
                }
                return false;
            });
            return valid ? set : null;
        } catch (XmlException) {
            return null;
        }
    }

    /// <summary>Reads only the document defaults without constructing an index of the styles.</summary>
    internal static bool TryReadDocumentDefaults(StylesPart part, out OfficeXmlNode defaults) {
        defaults = default;
        XmlReaderSettings? settings = CreateStreamingSettings(part);
        if (settings == null) return false;
        settings.IgnoreComments = true;
        settings.IgnoreProcessingInstructions = true;
        settings.IgnoreWhitespace = true;
        try {
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using XmlReader reader = XmlReader.Create(stream, settings);
            if (reader.MoveToContent() != XmlNodeType.Element || !IsWord(reader, "styles")) return false;
            string? markup = null;
            ReadChildren(reader, () => {
                if (markup != null || !IsWord(reader, "docDefaults")) return false;
                markup = reader.ReadOuterXml();
                return true;
            });
            if (markup == null) return true;
            byte[] bytes = Encoding.UTF8.GetBytes(markup);
            OfficeXmlTree tree = OfficeXmlTree.Parse(bytes, bytes.Length, new OfficeXmlNamespaces.Table());
            defaults = new OfficeXmlNode(tree, tree.Root);
            return true;
        } catch (Exception exception) when (exception is XmlException or InvalidDataException) {
            return false;
        }
    }

    /// <summary>
    /// Counts every element of an unloaded styles part, root included, as many elements as its DOM would hold;
    /// null when <see cref="CreateStreamingSettings"/> declines the part or its XML cannot be read.
    /// </summary>
    internal static long? TryCountElements(StylesPart? part) => TryCount(part, _ => true);

    /// <summary>
    /// Counts the elements below the root of an unloaded styles part whose local name matches
    /// <paramref name="localName"/> ignoring case, as a search of its DOM's descendants would; null when the part
    /// cannot be streamed.
    /// </summary>
    internal static long? TryCountDescendantsByLocalName(StylesPart? part, string localName) =>
        TryCount(part, reader => reader.Depth > 0 && string.Equals(reader.LocalName, localName, StringComparison.OrdinalIgnoreCase));

    // A comment or processing instruction becomes a child node of the loaded DOM, which walks count; such a part
    // is left to the DOM so streamed counts never fall below loaded ones.
    private static long? TryCount(StylesPart? part, Func<XmlReader, bool> counts) {
        XmlReaderSettings? settings = part == null ? null : CreateStreamingSettings(part);
        if (part == null || settings == null) return null;
        settings.IgnoreComments = false;
        settings.IgnoreProcessingInstructions = false;
        settings.IgnoreWhitespace = true;
        try {
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using XmlReader reader = XmlReader.Create(stream, settings);
            long count = 0;
            while (reader.Read()) {
                if (reader.NodeType is XmlNodeType.Comment or XmlNodeType.ProcessingInstruction) return null;
                if (reader.NodeType == XmlNodeType.Element && counts(reader)) count++;
            }
            return count;
        } catch (XmlException) {
            return null;
        }
    }

    private static readonly ConditionalWeakTable<OpenXmlPackage, object> HandedOutPackages = new();

    /// <summary>
    /// Stops streaming every part of <paramref name="package"/>, including parts added or replaced later. Once the
    /// raw package has been handed out, its parts can be rewritten without OfficeIMO noticing, so they are read
    /// from their DOM.
    /// </summary>
    internal static void StopStreaming(OpenXmlPackage? package) {
        if (package != null) HandedOutPackages.GetValue(package, _ => new object());
    }

    /// <summary>
    /// Whether <paramref name="part"/> can be read from its stream with the same meaning as its DOM: it is not
    /// loaded, its package does not process markup compatibility, and the raw package was not handed out.
    /// </summary>
    internal static bool CanStream(OpenXmlPart part) =>
        !part.IsRootElementLoaded &&
        part.OpenXmlPackage.MarkupCompatibilityProcessSettings.ProcessMode == MarkupCompatibilityProcessMode.NoProcess &&
        !HandedOutPackages.TryGetValue(part.OpenXmlPackage, out _);

    /// <summary>
    /// Reader settings that match how the package loads <paramref name="part"/>, or null when a streaming read
    /// cannot: the part is already loaded (its DOM may carry unsaved edits), or the package was opened with
    /// markup compatibility processing, which rewrites alternate content and ignorable markup while loading.
    /// The package's per-part character limit applies as it does to the DOM; exceeding it falls back to the DOM,
    /// which reports it.
    /// </summary>
    internal static XmlReaderSettings? CreateStreamingSettings(OpenXmlPart part) {
        if (!CanStream(part)) return null;
        var settings = new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            CloseInput = false
        };
        long maxCharacters = part.OpenXmlPackage.MaxCharactersInPart;
        if (maxCharacters > 0) settings.MaxCharactersInDocument = maxCharacters;
        return settings;
    }

    private static StyleDefinition? ReadStyle(XmlReader reader, StyleDefinitionSet set) {
        string? type = reader.GetAttribute("type", WordprocessingNamespace);
        if (type != null && type is not ("paragraph" or "character" or "table" or "numbering")) return null;
        if (!TryReadOnOff(reader.GetAttribute("default", WordprocessingNamespace), out bool isDefault)) return null;
        if (!TryReadOnOff(reader.GetAttribute("customStyle", WordprocessingNamespace), out bool isCustom)) return null;
        var style = new StyleDefinition(set, reader.GetAttribute("styleId", WordprocessingNamespace), type, isDefault) {
            IsCustom = isCustom
        };
        bool valid = true, seenName = false, seenBasedOn = false, seenLink = false, seenNext = false, seenParagraph = false;
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
                case "link" when !seenLink:
                    seenLink = true;
                    style.LinkedStyle = reader.GetAttribute("val", WordprocessingNamespace);
                    return false;
                case "next" when !seenNext:
                    seenNext = true;
                    style.NextStyle = reader.GetAttribute("val", WordprocessingNamespace);
                    return false;
                case "qFormat":
                    style.IsPrimary = true;
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
    private static bool ReadDocumentDefaults(string outerXml, XmlReaderSettings settings, out bool? caps, out bool? smallCaps) {
        using var text = new StringReader(outerXml);
        using XmlReader reader = XmlReader.Create(text, settings.Clone());
        reader.MoveToContent();
        bool valid = true, seenDefault = false;
        bool? capsValue = null, smallCapsValue = null;
        ReadChildren(reader, () => {
            if (seenDefault || !IsWord(reader, "rPrDefault")) return false;
            seenDefault = true;
            bool seenProperties = false;
            ReadChildren(reader, () => {
                if (seenProperties || !IsWord(reader, "rPr")) return false;
                seenProperties = true;
                valid &= ReadCaps(reader, out capsValue, out smallCapsValue);
                return true;
            });
            return true;
        });
        caps = capsValue;
        smallCaps = smallCapsValue;
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
