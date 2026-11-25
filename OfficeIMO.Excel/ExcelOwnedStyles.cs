using System.Globalization;
using System.Text;
using System.Xml;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Excel {
    /// <summary>A literal (<c>rgb</c>) or theme (<c>theme</c>) color as a SpreadsheetML color element carries it.</summary>
    internal readonly struct ExcelStyleColor {
        internal ExcelStyleColor(string attribute, string value) {
            Attribute = attribute;
            Value = value;
        }

        internal string Attribute { get; }
        internal string Value { get; }

        /// <summary>The native color the model writes for <paramref name="color"/>.</summary>
        internal static ExcelStyleColor From(OfficeColor color, OfficeTheme? theme) {
            uint? index = OfficeIMO.OpenXml.Internal.OfficeOpenXmlThemeColorResolver.GetSpreadsheetThemeColorIndex(color.ThemeSlot);
            return index.HasValue
                ? new ExcelStyleColor("theme", index.Value.ToString(CultureInfo.InvariantCulture))
                : new ExcelStyleColor("rgb", (theme != null ? color.Resolve(theme) : color).ToArgbHex());
        }
    }

    internal enum ExcelStyleOperationKind : byte {
        Date,
        Bold,
        Italic,
        Underline,
        Strike,
        VerticalTextAlignment,
        FontName,
        FontSize,
        FontColor,
        Fill,
        Border,
        NumberFormat,
        HorizontalAlignment,
        VerticalAlignment,
        WrapText,
        ShrinkToFit,
        TextRotation,
    }

    /// <summary>
    /// The stylesheet of a workbook saved without the SDK. Fonts, fills, borders, number formats and cell formats are
    /// read into small element trees; a style change clones the cell's format, changes what the model's style method
    /// changes, and reuses an existing entry with the same content, as the model does. Existing entries are never
    /// rewritten: new ones are appended and the counts updated, so everything else keeps its bytes.
    /// </summary>
    internal sealed class ExcelOwnedStyles {
        private static readonly string[] FormatAttributeOrder = {
            "numFmtId", "fontId", "fillId", "borderId", "xfId", "quotePrefix", "pivotButton", "applyNumberFormat", "applyFont",
            "applyFill", "applyBorder", "applyAlignment", "applyProtection"
        };
        private static readonly string[] FormatChildOrder = { "alignment", "protection", "extLst" };
        private static readonly string[] FontChildOrder = {
            "b", "i", "strike", "condense", "extend", "outline", "shadow", "u", "vertAlign", "sz", "color", "name", "family", "charset", "scheme"
        };
        private static readonly string[] BorderChildOrder = { "start", "left", "end", "right", "top", "bottom", "diagonal", "vertical", "horizontal" };
        private static readonly string[] AlignmentAttributeOrder = {
            "horizontal", "vertical", "textRotation", "wrapText", "indent", "relativeIndent", "justifyLastLine", "shrinkToFit", "readingOrder", "mergeCell"
        };

        private readonly OfficeOpcEditablePackage _package;
        private readonly string _partName;
        private readonly string _prefix;
        private readonly HashSet<string> _rootPrefixes;
        private readonly Collection _numberFormats = new("numFmts");
        private readonly Collection _fonts = new("fonts");
        private readonly Collection _fills = new("fills");
        private readonly Collection _borders = new("borders");
        private readonly Collection _formats = new("cellXfs");
        private readonly Dictionary<string, uint> _numberFormatIds = new(StringComparer.Ordinal);
        private readonly Dictionary<uint, string> _numberFormatCodes = new();
        // The model numbers a new format one past the highest it holds, or 164 in an empty table.
        private uint _highestNumberFormatId;
        private uint? _nextNumberFormatIdValue;
        private uint NextNumberFormatId() {
            uint next = _nextNumberFormatIdValue ?? (_numberFormats.Items.Count == 0 ? 164U : checked(_highestNumberFormatId + 1));
            _nextNumberFormatIdValue = next + 1;
            return next;
        }
        // Whether the model's completion of the stylesheet is one this class can make, whether it changes anything,
        // and whether a style change asked for it.
        private bool _canComplete;
        private bool _needsCompletion;
        private bool _completionUsed;
        private int _rootStartEnd;
        private int _rootEnd = -1;
        private Dictionary<string, (int Declared, int Actual)> _counts = new(StringComparer.Ordinal);
        private readonly List<(string Name, int Start)> _sections = new();
        private static readonly string[] SectionOrder = {
            "numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs", "tableStyles", "colors", "extLst"
        };

        private ExcelOwnedStyles(OfficeOpcEditablePackage package, string partName, string prefix, HashSet<string> rootPrefixes) {
            _package = package;
            _partName = partName;
            _prefix = prefix;
            _rootPrefixes = rootPrefixes;
        }

        internal bool Changed => _completionUsed && _needsCompletion || _numberFormats.Added.Count + _fonts.Added.Count + _fills.Added.Count + _borders.Added.Count + _formats.Added.Count > 0;

        /// <summary>A collection of style entries and where new entries go.</summary>
        private sealed class Collection {
            internal Collection(string name) => Name = name;
            internal string Name { get; }
            internal readonly List<Node> Items = new();
            internal readonly Dictionary<string, int> Indexes = new(StringComparer.Ordinal);
            internal readonly List<Node> Added = new();
            internal bool Seen;
            internal int Declared = -1;
            internal int End = -1;

            internal void AddExisting(Node node) {
                string key = node.Key;
                if (!Indexes.ContainsKey(key)) Indexes[key] = Items.Count;
                Items.Add(node);
            }

            /// <summary>Finds an entry with the same content or appends <paramref name="node"/>.</summary>
            internal int IndexOf(Node node) {
                string key = node.Key;
                if (Indexes.TryGetValue(key, out int index)) return index;
                index = Items.Count;
                Items.Add(node);
                Added.Add(node);
                Indexes[key] = index;
                return index;
            }
        }

        /// <summary>A style element: a SpreadsheetML element with its attributes and children, or markup kept as written.</summary>
        private sealed class Node {
            internal string Name = string.Empty;
            internal string? Raw;
            internal readonly List<(string Name, string Value)> Attributes = new();
            internal readonly List<Node> Children = new();
            internal string? Text;

            internal string Key {
                get {
                    var key = new StringBuilder();
                    AppendKey(key);
                    return key.ToString();
                }
            }

            private void AppendKey(StringBuilder key) {
                if (Raw != null) {
                    key.Append("raw:").Append(Raw);
                    return;
                }
                key.Append('<').Append(Name);
                // Attribute order carries no meaning; the model compares entries with their attributes in schema order.
                foreach ((string name, string value) in Attributes.OrderBy(static attribute => attribute.Name, StringComparer.Ordinal)) {
                    key.Append(' ').Append(name).Append("=\"").Append(value).Append('"');
                }
                key.Append('>');
                foreach (Node child in Children) child.AppendKey(key);
                if (Text != null) key.Append(Text);
                key.Append("</>");
            }

            internal Node Clone() {
                var clone = new Node { Name = Name, Raw = Raw, Text = Text };
                clone.Attributes.AddRange(Attributes);
                foreach (Node child in Children) clone.Children.Add(child.Clone());
                return clone;
            }

            internal string? Attribute(string name) {
                foreach ((string attributeName, string value) in Attributes) {
                    if (attributeName == name) return value;
                }
                return null;
            }

            internal void SetAttribute(string name, string? value) {
                int index = Attributes.FindIndex(attribute => attribute.Name == name);
                if (value == null) {
                    if (index >= 0) Attributes.RemoveAt(index);
                } else if (index >= 0) {
                    Attributes[index] = (name, value);
                } else {
                    Attributes.Add((name, value));
                }
            }

            internal Node? Child(string name) => Children.Find(child => child.Raw == null && child.Name == name);

            /// <summary>Replaces every child named <paramref name="name"/> with <paramref name="replacement"/>, placed by <paramref name="order"/>.</summary>
            internal void SetChild(string name, Node? replacement, string[] order) {
                Children.RemoveAll(child => child.Raw == null && child.Name == name);
                if (replacement == null) return;
                int rank = Array.IndexOf(order, name);
                int position = Children.Count;
                for (int index = 0; index < Children.Count; index++) {
                    int other = Children[index].Raw == null ? Array.IndexOf(order, Children[index].Name) : -1;
                    if (other > rank || other < 0 && Children[index].Raw != null) {
                        position = index;
                        break;
                    }
                }
                Children.Insert(position, replacement);
            }

            internal static Node Element(string name, params (string Name, string Value)[] attributes) {
                var node = new Node { Name = name };
                node.Attributes.AddRange(attributes);
                return node;
            }
        }

        /// <summary>Reads the stylesheet, or returns null when it needs the model.</summary>
        internal static ExcelOwnedStyles? Load(OfficeOpcEditablePackage package, string? partName) {
            if (partName == null) return null;
            OfficePooledBytes content = package.ReadPart(partName);
            using (content) {
                var reader = new OfficeXmlPullReader(content.Array, content.Length, processMarkupCompatibility: false);
                if (reader.Transcoded) return null;
                ExcelOwnedStyles? styles = null;
                var counts = new Dictionary<string, (int Declared, int Actual)>(StringComparer.Ordinal);
                bool firstFontComplete = false, hasNoneFill = false, hasGrayFill = false;
                Collection? current = null;
                string? section = null;
                while (reader.Read()) {
                    int depth = reader.Depth;
                    if (reader.NodeKind == OfficeXmlNodeKind.EndElement) {
                        if (depth == 1 && styles != null) styles._rootEnd = reader.MarkupStart;
                        if (depth == 2) {
                            if (current != null && reader.MarkupEnd > reader.MarkupStart) current.End = reader.MarkupStart;
                            current = null;
                            section = null;
                        }
                        continue;
                    }
                    if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
                    if (depth == 1) {
                        if (!reader.IsElement(OfficeXmlNamespaces.SpreadsheetML, "styleSheet")
                            || reader.ElementUriId != OfficeXmlNamespaces.TransitionalUriId(OfficeXmlNamespaces.SpreadsheetML)) {
                            return null;
                        }
                        var prefixes = new HashSet<string>(StringComparer.Ordinal);
                        for (int index = 0; index < reader.AttributeCount; index++) {
                            if (reader.GetAttributeNamespace(index, out _) == OfficeXmlNamespaces.Xmlns) prefixes.Add(reader.GetAttributeLocalName(index));
                        }
                        styles = new ExcelOwnedStyles(package, partName, reader.ElementPrefix.Length == 0 ? string.Empty : reader.ElementPrefix + ":", prefixes) {
                            _rootStartEnd = reader.MarkupEnd
                        };
                        continue;
                    }
                    if (styles == null) return null;
                    if (depth == 2) {
                        // Compatibility markup can hold any section; which one applies is the model's to decide.
                        if (reader.NamespaceId != OfficeXmlNamespaces.SpreadsheetML) return null;
                        section = reader.LocalName;
                        styles._sections.Add((section ?? string.Empty, reader.MarkupStart));
                        current = section switch {
                            "numFmts" => styles._numberFormats,
                            "fonts" => styles._fonts,
                            "fills" => styles._fills,
                            "borders" => styles._borders,
                            "cellXfs" => styles._formats,
                            _ => null
                        };
                        if (section == null) continue;
                        if (counts.ContainsKey(section)) return null;
                        int declared = reader.TryGetAttribute(OfficeXmlNamespaces.None, "count", out string countText)
                            && int.TryParse(countText, NumberStyles.None, CultureInfo.InvariantCulture, out int parsed) ? parsed : -1;
                        counts[section] = (declared, 0);
                        if (current != null) {
                            current.Seen = true;
                            current.Declared = declared;
                        }
                        continue;
                    }
                    if (depth != 3 || section == null) continue;
                    // Compatibility markup between entries hides which entries count, and so every index after it.
                    if (reader.NamespaceId != OfficeXmlNamespaces.SpreadsheetML && section != "extLst") return null;
                    string child = reader.NamespaceId == OfficeXmlNamespaces.SpreadsheetML ? reader.LocalName : string.Empty;
                    bool counted = section switch {
                        "fonts" => child == "font",
                        "fills" => child == "fill",
                        "borders" => child == "border",
                        "cellStyleXfs" or "cellXfs" => child == "xf",
                        "cellStyles" => child == "cellStyle",
                        "dxfs" => child == "dxf",
                        "tableStyles" => child == "tableStyle",
                        "numFmts" => child == "numFmt",
                        _ => false
                    };
                    var count = counts[section];
                    if (counted) counts[section] = (count.Declared, count.Actual + 1);
                    if (current == null || !counted) {
                        reader.SkipElement();
                        continue;
                    }
                    Node? node = ReadNode(reader, styles._rootPrefixes);
                    if (node == null) return null;
                    if (section == "fonts" && current.Items.Count == 0) firstFontComplete = node.Child("sz") != null && node.Child("name") != null;
                    if (section == "fills" && node.Child("patternFill")?.Attribute("patternType") is { } pattern) {
                        hasNoneFill |= pattern == "none";
                        hasGrayFill |= pattern == "gray125";
                    }
                    if (section == "numFmts") styles.ReadNumberFormat(node);
                    current.AddExisting(node);
                }
                if (styles == null || !styles._formats.Seen || styles._formats.End < 0) return null;
                // The model completes a stylesheet before it changes a format. Missing cell styles, differential formats
                // and table styles, and wrong counts, are completed here as the model completes them; a stylesheet
                // without fonts, fills, borders or formats to build on stays with the model.
                bool canComplete = firstFontComplete && hasNoneFill && hasGrayFill;
                foreach (string required in new[] { "fonts", "fills", "borders", "cellStyleXfs", "cellXfs" }) {
                    if (!counts.TryGetValue(required, out var count) || count.Actual == 0) canComplete = false;
                }
                foreach (Collection collection in new[] { styles._fonts, styles._fills, styles._borders }) {
                    if (collection.End < 0) canComplete = false;
                }
                if (styles._numberFormats.Seen && styles._numberFormats.End < 0) canComplete = false;
                if (counts.TryGetValue("cellStyles", out var cellStyles) && cellStyles.Actual == 0) canComplete = false;
                if (styles._rootEnd < 0) canComplete = false;
                bool needsCompletion = false;
                foreach (var count in counts) {
                    if (count.Key is "fonts" or "fills" or "borders" or "cellStyleXfs" or "cellXfs" or "cellStyles" or "dxfs" or "tableStyles" or "numFmts"
                        && count.Value.Declared != count.Value.Actual) {
                        needsCompletion = true;
                    }
                }
                foreach (string optional in new[] { "cellStyles", "dxfs", "tableStyles" }) {
                    if (!counts.ContainsKey(optional)) needsCompletion = true;
                }
                styles._canComplete = canComplete;
                styles._needsCompletion = needsCompletion;
                styles._counts = counts;
                return styles;
            }
        }

        private void ReadNumberFormat(Node node) {
            string? idText = node.Attribute("numFmtId");
            string? code = node.Attribute("formatCode");
            if (idText == null || !uint.TryParse(idText, NumberStyles.None, CultureInfo.InvariantCulture, out uint id)) return;
            if (!_numberFormatCodes.ContainsKey(id)) _numberFormatCodes[id] = code ?? string.Empty;
            if (code != null) {
                if (!_numberFormatIds.ContainsKey(code)) _numberFormatIds[code] = id;
                _highestNumberFormatId = Math.Max(_highestNumberFormatId, id);
            }
        }

        /// <summary>Reads the element the reader is on; null for markup the owned model cannot carry.</summary>
        private static Node? ReadNode(OfficeXmlPullReader reader, HashSet<string> rootPrefixes) {
            int start = reader.MarkupStart;
            bool foreign = reader.NamespaceId != OfficeXmlNamespaces.SpreadsheetML;
            var node = new Node { Name = reader.LocalName };
            for (int index = 0; index < reader.AttributeCount; index++) {
                int namespaceId = reader.GetAttributeNamespace(index, out _);
                if (namespaceId == OfficeXmlNamespaces.Xmlns) return null;
                string name = reader.GetAttributeLocalName(index);
                if (namespaceId != OfficeXmlNamespaces.None) {
                    string prefix = reader.GetAttributePrefix(index);
                    if (!rootPrefixes.Contains(prefix)) return null;
                    name = prefix + ":" + name;
                }
                node.Attributes.Add((name, reader.GetAttributeValue(index)));
            }
            if (foreign) {
                // Markup from other vocabularies is kept as written; its prefix is declared on the stylesheet.
                if (!rootPrefixes.Contains(reader.ElementPrefix)) return null;
                int depth = reader.Depth;
                if (!reader.IsEmptyElement) {
                    while (reader.Read()) {
                        if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == depth) break;
                        if (reader.NodeKind == OfficeXmlNodeKind.StartElement) {
                            for (int index = 0; index < reader.AttributeCount; index++) {
                                if (reader.GetAttributeNamespace(index, out _) == OfficeXmlNamespaces.Xmlns) return null;
                            }
                        }
                    }
                } else {
                    reader.Read();
                }
                return new Node { Name = node.Name, Raw = Encoding.UTF8.GetString(reader.Data, start, reader.MarkupEnd - start) };
            }
            if (reader.IsEmptyElement) {
                reader.Read();
                return node;
            }
            int elementDepth = reader.Depth;
            while (reader.Read()) {
                if (reader.NodeKind == OfficeXmlNodeKind.EndElement && reader.Depth == elementDepth) return node;
                if (reader.NodeKind == OfficeXmlNodeKind.Text) {
                    if (!reader.IsWhitespace) node.Text = (node.Text ?? string.Empty) + reader.GetText();
                } else if (reader.NodeKind == OfficeXmlNodeKind.StartElement) {
                    Node? child = ReadNode(reader, rootPrefixes);
                    if (child == null) return null;
                    node.Children.Add(child);
                }
            }
            return null;
        }

        /// <summary>The style a cell with style <paramref name="baseStyle"/> takes after <paramref name="operation"/>, or null when the model must apply it.</summary>
        internal uint? Apply(uint baseStyle, ExcelStyleOperation operation) {
            // The same change to the same format always gives the same format, as the model's derived-format cache has it.
            if (_applied.TryGetValue((baseStyle, operation), out uint known)) return known;
            uint? result = ApplyCore(baseStyle, operation);
            if (result.HasValue) _applied[(baseStyle, operation)] = result.Value;
            return result;
        }

        private readonly Dictionary<(uint, ExcelStyleOperation), uint> _applied = new();

        private uint? ApplyCore(uint baseStyle, ExcelStyleOperation operation) {
            if (operation.Kind == ExcelStyleOperationKind.Date) {
                if (baseStyle != 0U && HasCalendarFormat(baseStyle)) return baseStyle;
                return DeriveFormat(baseStyle, format => {
                    format.SetAttribute("numFmtId", "14");
                    format.SetAttribute("applyNumberFormat", "1");
                    return true;
                });
            }
            if (!_canComplete) return null;
            _completionUsed = true;
            switch (operation.Kind) {
                case ExcelStyleOperationKind.Bold: return DeriveFont(baseStyle, font => font.SetChild("b", operation.Flag ? Node.Element("b") : null, FontChildOrder));
                case ExcelStyleOperationKind.Italic: return DeriveFont(baseStyle, font => font.SetChild("i", operation.Flag ? Node.Element("i") : null, FontChildOrder));
                case ExcelStyleOperationKind.Strike: return DeriveFont(baseStyle, font => font.SetChild("strike", operation.Flag ? Node.Element("strike") : null, FontChildOrder));
                case ExcelStyleOperationKind.Underline:
                    return DeriveFont(baseStyle, font => font.SetChild("u", operation.Text == null ? null
                        : operation.Text.Length == 0 ? Node.Element("u") : Node.Element("u", ("val", operation.Text)), FontChildOrder));
                case ExcelStyleOperationKind.VerticalTextAlignment:
                    return DeriveFont(baseStyle, font => font.SetChild("vertAlign", Node.Element("vertAlign", ("val", operation.Text!)), FontChildOrder));
                case ExcelStyleOperationKind.FontName:
                    return DeriveFont(baseStyle, font => {
                        font.SetChild("scheme", null, FontChildOrder);
                        font.SetChild("name", Node.Element("name", ("val", operation.Text!.Trim())), FontChildOrder);
                    });
                case ExcelStyleOperationKind.FontSize:
                    return DeriveFont(baseStyle, font => font.SetChild("sz", Node.Element("sz", ("val", XmlConvert.ToString(operation.Number))), FontChildOrder));
                case ExcelStyleOperationKind.FontColor:
                    return DeriveFont(baseStyle, font => font.SetChild("color", Node.Element("color", (operation.Color!.Value.Attribute, operation.Color.Value.Value)), FontChildOrder));
                case ExcelStyleOperationKind.Fill: {
                    ExcelStyleColor color = operation.Color!.Value;
                    var pattern = Node.Element("patternFill", ("patternType", "solid"));
                    pattern.Children.Add(Node.Element("fgColor", (color.Attribute, color.Value)));
                    pattern.Children.Add(Node.Element("bgColor", (color.Attribute, color.Value)));
                    var fill = Node.Element("fill");
                    fill.Children.Add(pattern);
                    int fillId = _fills.IndexOf(fill);
                    return DeriveFormat(baseStyle, format => {
                        format.SetAttribute("fillId", fillId.ToString(CultureInfo.InvariantCulture));
                        format.SetAttribute("applyFill", "1");
                        return true;
                    });
                }
                case ExcelStyleOperationKind.Border: {
                    Node? baseBorder = BaseEntry(_borders, baseStyle, "borderId");
                    if (baseBorder != null && baseBorder.Raw != null) return null;
                    Node border = baseBorder?.Clone() ?? Node.Element("border");
                    foreach ((ExcelCellBorders flag, string side) in new[] {
                        (ExcelCellBorders.Left, "left"), (ExcelCellBorders.Top, "top"), (ExcelCellBorders.Right, "right"), (ExcelCellBorders.Bottom, "bottom") }) {
                        if (!operation.Borders.HasFlag(flag)) continue;
                        Node edge = Node.Element(side, ("style", operation.Text!));
                        if (operation.Color is { } color) edge.Children.Add(Node.Element("color", (color.Attribute, color.Value)));
                        border.SetChild(side, edge, BorderChildOrder);
                    }
                    int borderId = _borders.IndexOf(border);
                    return DeriveFormat(baseStyle, format => {
                        format.SetAttribute("borderId", borderId.ToString(CultureInfo.InvariantCulture));
                        format.SetAttribute("applyBorder", "1");
                        return true;
                    });
                }
                case ExcelStyleOperationKind.NumberFormat: {
                    string code = operation.Text!;
                    if (!_numberFormatIds.TryGetValue(code, out uint id)) {
                        id = NextNumberFormatId();
                        _numberFormatIds[code] = id;
                        _numberFormatCodes[id] = code;
                        _numberFormats.IndexOf(Node.Element("numFmt", ("numFmtId", id.ToString(CultureInfo.InvariantCulture)), ("formatCode", code)));
                    }
                    return DeriveFormat(baseStyle, format => {
                        format.SetAttribute("numFmtId", id.ToString(CultureInfo.InvariantCulture));
                        format.SetAttribute("applyNumberFormat", "1");
                        return true;
                    });
                }
                case ExcelStyleOperationKind.HorizontalAlignment: return DeriveAlignment(baseStyle, alignment => alignment.SetAttribute("horizontal", operation.Text));
                case ExcelStyleOperationKind.VerticalAlignment: return DeriveAlignment(baseStyle, alignment => alignment.SetAttribute("vertical", operation.Text));
                case ExcelStyleOperationKind.WrapText: return DeriveAlignment(baseStyle, alignment => alignment.SetAttribute("wrapText", operation.Flag ? "1" : null));
                case ExcelStyleOperationKind.ShrinkToFit: return DeriveAlignment(baseStyle, alignment => alignment.SetAttribute("shrinkToFit", operation.Flag ? "1" : "0"));
                case ExcelStyleOperationKind.TextRotation:
                    return DeriveAlignment(baseStyle, alignment => alignment.SetAttribute("textRotation", ((int)operation.Number).ToString(CultureInfo.InvariantCulture)));
                default:
                    return null;
            }
        }

        private uint? DeriveFont(uint baseStyle, Action<Node> mutate) {
            Node? baseFont = BaseEntry(_fonts, baseStyle, "fontId");
            if (baseFont != null && baseFont.Raw != null) return null;
            Node font = baseFont?.Clone() ?? Node.Element("font");
            mutate(font);
            if (font.Children.Exists(static child => child.Raw != null)) return null;
            font.Children.Sort((left, right) => Array.IndexOf(FontChildOrder, left.Name).CompareTo(Array.IndexOf(FontChildOrder, right.Name)));
            int fontId = _fonts.IndexOf(font);
            return DeriveFormat(baseStyle, format => {
                format.SetAttribute("fontId", fontId.ToString(CultureInfo.InvariantCulture));
                format.SetAttribute("applyFont", "1");
                return true;
            });
        }

        private uint? DeriveAlignment(uint baseStyle, Action<Node> mutate) =>
            DeriveFormat(baseStyle, format => {
                Node alignment = format.Child("alignment")?.Clone() ?? Node.Element("alignment");
                mutate(alignment);
                format.SetChild("alignment", alignment, FormatChildOrder);
                format.SetAttribute("applyAlignment", "1");
                return true;
            });

        /// <summary>The entry of <paramref name="collection"/> a cell format refers to, as the model finds it; null when out of range.</summary>
        private Node? BaseEntry(Collection collection, uint baseStyle, string attribute) {
            Node? format = baseStyle < (uint)_formats.Items.Count ? _formats.Items[(int)baseStyle] : null;
            string? idText = format?.Attribute(attribute);
            uint id = idText != null && uint.TryParse(idText, NumberStyles.None, CultureInfo.InvariantCulture, out uint parsed) ? parsed : 0U;
            return id < (uint)collection.Items.Count ? collection.Items[(int)id] : null;
        }

        private uint? DeriveFormat(uint baseStyle, Func<Node, bool> mutate) {
            Node format;
            if (baseStyle < (uint)_formats.Items.Count) {
                format = _formats.Items[(int)baseStyle].Clone();
                if (format.Raw != null) return null;
            } else {
                format = Node.Element("xf", ("numFmtId", "0"), ("fontId", "0"), ("fillId", "0"), ("borderId", "0"), ("xfId", "0"));
            }
            if (!mutate(format)) return null;
            // Deriving a format completes the stylesheet first, as the model does.
            if (!_canComplete) return null;
            _completionUsed = true;
            return (uint)_formats.IndexOf(format);
        }

        private bool HasCalendarFormat(uint styleIndex) {
            if (styleIndex >= (uint)_formats.Items.Count) return false;
            string? numberFormatText = _formats.Items[(int)styleIndex].Attribute("numFmtId");
            if (numberFormatText == null || !uint.TryParse(numberFormatText, NumberStyles.None, CultureInfo.InvariantCulture, out uint numberFormatId)
                || numberFormatId == 0U) {
                return false;
            }
            string? code = _numberFormatCodes.TryGetValue(numberFormatId, out string? custom) ? custom : ExcelBuiltInNumberFormats.GetCode(numberFormatId);
            return ExcelNumberFormatClassifier.ClassifySerialDateStyle(numberFormatId, code) == ExcelSerialDateStyle.Calendar;
        }

        /// <summary>Writes the stylesheet with its new entries.</summary>
        internal void Write() {
            if (!Changed) return;
            OfficePooledBytes content = _package.ReadPart(_partName);
            byte[] result;
            int resultLength;
            using (content) {
                var reader = new OfficeXmlPullReader(content.Array, content.Length, processMarkupCompatibility: false);
                var edits = new ExcelPartEdits();
                bool complete = _completionUsed && _needsCompletion;
                while (reader.Read()) {
                    if (reader.NodeKind != OfficeXmlNodeKind.StartElement || reader.Depth != 2 || reader.NamespaceId != OfficeXmlNamespaces.SpreadsheetML) continue;
                    string name = reader.LocalName;
                    Collection? collection = name switch {
                        "numFmts" => _numberFormats,
                        "fonts" => _fonts,
                        "fills" => _fills,
                        "borders" => _borders,
                        "cellXfs" => _formats,
                        _ => null
                    };
                    if (!_counts.TryGetValue(name, out var count)) continue;
                    int actual = collection?.Items.Count ?? count.Actual;
                    bool counted = name is "fonts" or "fills" or "borders" or "cellStyleXfs" or "cellXfs" or "cellStyles" or "dxfs" or "tableStyles" or "numFmts";
                    if (counted && (collection != null && collection.Added.Count > 0 || complete && count.Declared != actual)) {
                        edits.SetAttribute(reader, "count", actual.ToString(CultureInfo.InvariantCulture));
                    }
                }
                if (complete) {
                    // Missing sections go where the schema puts them, as the model adds them.
                    InsertMissingSection(edits, "cellStyles", "<" + _prefix + "cellStyles count=\"1\"><" + _prefix + "cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\" /></" + _prefix + "cellStyles>");
                    InsertMissingSection(edits, "dxfs", "<" + _prefix + "dxfs count=\"0\" />");
                    InsertMissingSection(edits, "tableStyles", "<" + _prefix + "tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleLight16\" />");
                }
                foreach (Collection collection in new[] { _numberFormats, _fonts, _fills, _borders, _formats }) {
                    if (collection.Added.Count == 0) continue;
                    var xml = new StringBuilder();
                    foreach (Node node in collection.Added) AppendNode(xml, node);
                    if (collection.Seen) {
                        edits.Insert(collection.End, xml.ToString());
                    } else {
                        // The number formats come first in a stylesheet; the model adds them when the first one is needed.
                        edits.Insert(_rootStartEnd, "<" + _prefix + collection.Name + " count=\"" + collection.Items.Count.ToString(CultureInfo.InvariantCulture)
                            + "\">" + xml + "</" + _prefix + collection.Name + ">");
                    }
                }
                result = edits.Apply(content.Array, content.Length, out resultLength);
            }
            _package.ReplacePart(_partName, result, resultLength);
        }

        private void InsertMissingSection(ExcelPartEdits edits, string name, string xml) {
            if (_counts.ContainsKey(name)) return;
            int rank = Array.IndexOf(SectionOrder, name);
            foreach ((string section, int start) in _sections) {
                int other = Array.IndexOf(SectionOrder, section);
                if (other > rank) {
                    edits.Insert(start, xml);
                    return;
                }
            }
            edits.Insert(_rootEnd, xml);
        }

        private void AppendNode(StringBuilder xml, Node node) {
            if (node.Raw != null) {
                xml.Append(node.Raw);
                return;
            }
            xml.Append('<').Append(_prefix).Append(node.Name);
            string[]? order = node.Name switch { "xf" => FormatAttributeOrder, "alignment" => AlignmentAttributeOrder, _ => null };
            IEnumerable<(string Name, string Value)> attributes = order == null
                ? node.Attributes
                : node.Attributes.OrderBy(attribute => Array.IndexOf(order, attribute.Name) is int rank && rank >= 0 ? rank : int.MaxValue);
            foreach ((string name, string value) in attributes) {
                xml.Append(' ').Append(name).Append("=\"").Append(ExcelPartEdits.Escape(value)).Append('"');
            }
            if (node.Children.Count == 0 && node.Text == null) {
                xml.Append(" />");
                return;
            }
            xml.Append('>');
            foreach (Node child in node.Children) AppendNode(xml, child);
            if (node.Text != null) xml.Append(ExcelPartEdits.Escape(node.Text));
            xml.Append("</").Append(_prefix).Append(node.Name).Append('>');
        }
    }
}
