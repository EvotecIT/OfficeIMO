using System.Globalization;
using System.Text;
using System.Xml;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word {
    /// <summary>
    /// How the comparer reads WordprocessingML: elements are <see cref="OfficeXmlNode"/>s from <see cref="WordPartView"/>
    /// parts, read with the meaning the SDK's typed members gave them.
    /// </summary>
    public static partial class WordDocumentComparer {
        private const int W = OfficeXmlNamespaces.WordprocessingML;
        private const int Dml = OfficeXmlNamespaces.DrawingML;
        private const int Pic = OfficeXmlNamespaces.Picture;
        private const int Wp = OfficeXmlNamespaces.WordprocessingDrawing;
        private const int Vml = OfficeXmlNamespaces.Vml;
        private const int Mc = OfficeXmlNamespaces.MarkupCompatibility;
        private const string WordprocessingNamespaceUri = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        // ---- Packages ---------------------------------------------------------------------------------------------

        /// <summary>The main document part of the package that holds <paramref name="part"/>.</summary>
        private static WordPartView? GetMainDocumentPart(WordPartView? part) => part?.Package.Main;

        private static OfficeXmlNode DocumentRoot(WordPartView? mainPart) => mainPart?.RootNamed(W, "document") ?? default;

        private static OfficeXmlNode BodyOf(WordPartView? mainPart) => DocumentRoot(mainPart).Child(W, "body");

        private static IEnumerable<WordPartView> HeaderParts(WordPartView mainPart) => mainPart.Parts(WordPackageView.HeaderRelationshipType);

        private static IEnumerable<WordPartView> FooterParts(WordPartView mainPart) => mainPart.Parts(WordPackageView.FooterRelationshipType);

        private static WordPartView? FootnotesPart(WordPartView mainPart) => mainPart.FirstPart(WordPackageView.FootnotesRelationshipType);

        private static WordPartView? EndnotesPart(WordPartView mainPart) => mainPart.FirstPart(WordPackageView.EndnotesRelationshipType);

        private static OfficeXmlNode HeaderRoot(WordPartView part) => part.RootNamed(W, "hdr");

        private static OfficeXmlNode FooterRoot(WordPartView part) => part.RootNamed(W, "ftr");

        private static OfficeXmlNode FootnotesRoot(WordPartView? part) => part?.RootNamed(W, "footnotes") ?? default;

        private static OfficeXmlNode EndnotesRoot(WordPartView? part) => part?.RootNamed(W, "endnotes") ?? default;

        private static OfficeXmlNode NumberingRoot(WordPartView? mainPart) =>
            mainPart?.FirstPart(WordPackageView.NumberingRelationshipType)?.RootNamed(W, "numbering") ?? default;

        private static OfficeXmlNode SettingsRoot(WordPartView mainPart) =>
            mainPart.FirstPart(WordPackageView.SettingsRelationshipType)?.RootNamed(W, "settings") ?? default;

        // ---- Element names --------------------------------------------------------------------------------------

        private static bool IsW(OfficeXmlNode node, string localName) => node.Is(W, localName);

        /// <summary>A run-level <c>w:tab</c>; the same name in <c>w:tabs</c> is a tab stop.</summary>
        private static bool IsTabChar(OfficeXmlNode node) => node.Is(W, "tab") && !node.Parent.Is(W, "tabs");

        private static bool HasAncestorW(OfficeXmlNode node, string localName) => !WordXmlNodes.Ancestor(node, W, localName).IsNone;

        private static OfficeXmlNode AncestorW(OfficeXmlNode node, string localName) => WordXmlNodes.Ancestor(node, W, localName);

        private static List<OfficeXmlNode> ChildrenW(OfficeXmlNode node, string localName) =>
            node.IsNone ? new List<OfficeXmlNode>() : node.Children(W, localName).ToList();

        private static string InnerText(OfficeXmlNode node) => WordXmlNodes.InnerText(node);

        /// <summary>The <c>w:val</c> of a WordprocessingML element, or null.</summary>
        private static string? WVal(OfficeXmlNode node) => node.IsNone ? null : node.Attribute(W, "val");

        private static string? WAttr(OfficeXmlNode node, string localName) => node.IsNone ? null : node.Attribute(W, localName);

        /// <summary>The <c>w:val</c> of the first child with this name, or null.</summary>
        private static string? ChildVal(OfficeXmlNode node, string localName) => WordXmlEdits.ChildValue(node, localName);

        /// <summary>An element whose <c>w:val</c> is the paragraph or run style: its value, or null.</summary>
        private static string? StyleIdOf(OfficeXmlNode properties, string styleElement) => ChildVal(properties, styleElement);

        // ---- Typed values ---------------------------------------------------------------------------------------

        /// <summary>An integer attribute as the SDK's typed value reads it: null when absent, a format error when invalid.</summary>
        private static int? ParseInt32(string? text) => text == null ? null : XmlConvert.ToInt32(text);

        private static long? ParseInt64(string? text) => text == null ? null : XmlConvert.ToInt64(text);

        private static uint? ParseUInt32(string? text) => text == null ? null : XmlConvert.ToUInt32(text);

        private static string? Int32Text(string? text) => ParseInt32(text)?.ToString(CultureInfo.InvariantCulture);

        private static string? Int64Text(string? text) => ParseInt64(text)?.ToString(CultureInfo.InvariantCulture);

        /// <summary>An on/off attribute value: null when absent, a format error for a value the SDK doesn't parse.</summary>
        private static bool? ParseOnOff(string? text) {
            switch (text) {
                case null: return null;
                case "true": case "1": case "on": return true;
                case "false": case "0": case "off": return false;
                default: throw new FormatException("The text value is invalid. It can be only 'true', 'false', 'on', 'off', '0', '1'.");
            }
        }

        /// <summary>An enumeration attribute: null when absent, the token when the SDK maps it, a format error otherwise.</summary>
        private static string? EnumToken(string? text, params string[] tokens) {
            if (text == null) return null;
            foreach (string token in tokens) {
                if (string.Equals(token, text, StringComparison.Ordinal)) return token;
            }
            throw new FormatException("The text value is not a valid enumeration value.");
        }

        /// <summary>An on/off property element: present with no value or a true value is on.</summary>
        private static bool IsOnOffEnabled(OfficeXmlNode onOff) => WordXmlEdits.ReadOnOff(onOff);

        // ---- Document order -------------------------------------------------------------------------------------

        private static IEnumerable<OrderedElement> EnumerateDescendantsWithOrder(OfficeXmlNode container, int orderBase) {
            if (container.IsNone) {
                yield break;
            }

            int order = orderBase;
            foreach (OfficeXmlNode element in EnumerateComparableDescendants(container)) {
                yield return new OrderedElement(element, order);
                order++;
            }
        }

        /// <summary>
        /// The elements below <paramref name="container"/> in document order, each <c>mc:AlternateContent</c> read through
        /// its selected branch. An element inside an extension list (<c>extLst</c>) comes back as
        /// <see cref="OfficeXmlNode.None"/>: it still takes its place in document order, but the SDK read such markup as
        /// unknown elements, so no typed match finds it.
        /// </summary>
        private static IEnumerable<OfficeXmlNode> EnumerateComparableDescendants(OfficeXmlNode container) {
            if (container.IsNone) yield break;
            var pending = new Stack<(OfficeXmlNode Next, bool InExtension)>();
            OfficeXmlNode current = FirstComparableChild(container);
            bool inExtension = false;
            while (true) {
                if (current.IsNone) {
                    if (pending.Count == 0) yield break;
                    (current, inExtension) = pending.Pop();
                    continue;
                }

                yield return inExtension ? default : current;
                OfficeXmlNode next = current.NextSibling;
                OfficeXmlNode child = FirstComparableChild(current);
                if (child.IsNone) {
                    current = next;
                } else {
                    pending.Push((next, inExtension));
                    inExtension = inExtension || IsExtensionList(current);
                    current = child;
                }
            }
        }

        /// <summary>An extension list, whose content the SDK read as unknown elements.</summary>
        private static bool IsExtensionList(OfficeXmlNode element) => string.Equals(element.LocalName, "extLst", StringComparison.Ordinal);

        /// <summary>Whether the element sits inside an extension list.</summary>
        private static bool IsInsideExtensionList(OfficeXmlNode element) {
            for (OfficeXmlNode ancestor = element.Parent; !ancestor.IsNone; ancestor = ancestor.Parent) {
                if (IsExtensionList(ancestor)) return true;
            }

            return false;
        }

        /// <summary>
        /// The elements below <paramref name="element"/> in document order that the SDK's typed lookups could find: the
        /// content of extension lists is left out.
        /// </summary>
        private static IEnumerable<OfficeXmlNode> TypedDescendants(OfficeXmlNode element) {
            if (element.IsNone) yield break;
            OfficeXmlNode node = element.FirstChild;
            while (!node.IsNone) {
                yield return node;
                OfficeXmlNode child = IsExtensionList(node) ? default : node.FirstChild;
                if (!child.IsNone) {
                    node = child;
                    continue;
                }

                while (!node.IsNone && node != element && node.NextSibling.IsNone) node = node.Parent;
                if (node.IsNone || node == element) yield break;
                node = node.NextSibling;
            }
        }

        private static List<OfficeXmlNode> DescendantsOf(OfficeXmlNode element, int namespaceId, string localName) {
            var found = new List<OfficeXmlNode>();
            foreach (OfficeXmlNode node in TypedDescendants(element)) {
                if (node.Is(namespaceId, localName)) found.Add(node);
            }

            return found;
        }

        private static OfficeXmlNode FirstComparableChild(OfficeXmlNode element) {
            if (element.Is(Mc, "AlternateContent")) {
                OfficeXmlNode branch = WordAlternateContentResolver.SelectBranch(element);
                return branch.IsNone ? default : branch.FirstChild;
            }

            return element.FirstChild;
        }

        /// <summary>The elements below <paramref name="element"/> in document order, without branch selection.</summary>
        private static IEnumerable<OfficeXmlNode> Descendants(OfficeXmlNode element) => WordXmlNodes.Descendants(element);

        private static List<OfficeXmlNode> DescendantsW(OfficeXmlNode element, string localName) => DescendantsOf(element, W, localName);

        private static bool HasDescendant(OfficeXmlNode element, int namespaceId, string localName) => !FirstDescendant(element, namespaceId, localName).IsNone;

        private static OfficeXmlNode FirstDescendant(OfficeXmlNode element, int namespaceId, string localName) {
            foreach (OfficeXmlNode node in TypedDescendants(element)) {
                if (node.Is(namespaceId, localName)) return node;
            }

            return default;
        }

        // ---- Signatures -----------------------------------------------------------------------------------------

        /// <summary>
        /// Writes an element so that two elements write the same text exactly when they hold the same markup: names by
        /// namespace URI, attributes in a fixed order, and the text of leaf elements. It takes the place of comparing
        /// the SDK's serialized markup, and reads the same whichever layer holds the element.
        /// </summary>
        private static string GetElementSignature(OfficeXmlNode element, SignatureRules? rules = null) {
            if (element.IsNone) return string.Empty;
            var builder = new StringBuilder();
            AppendElementSignature(element, builder, rules ?? SignatureRules.None, new List<WordXmlAttribute>());
            return builder.ToString();
        }

        private static void AppendElementSignature(OfficeXmlNode element, StringBuilder builder, SignatureRules rules, List<WordXmlAttribute> attributes) {
            builder.Append('<').Append(element.NamespaceUri).Append(' ').Append(element.LocalName);
            attributes.Clear();
            WordXmlNodes.AddAttributes(element, attributes);
            rules.Rewrite(element, attributes);
            if (attributes.Count > 1) attributes.Sort(CompareAttributes);
            foreach (WordXmlAttribute attribute in attributes) {
                builder.Append(' ').Append(attribute.NamespaceUri).Append(' ').Append(attribute.LocalName).Append('=')
                    .Append(attribute.Value.Length.ToString(CultureInfo.InvariantCulture)).Append(':').Append(attribute.Value);
            }

            builder.Append('>');
            OfficeXmlNode child = element.FirstChild;
            if (child.IsNone) {
                string text = element.InnerText;
                if (text.Length > 0) builder.Append('"').Append(text.Length.ToString(CultureInfo.InvariantCulture)).Append(':').Append(text);
            } else {
                for (; !child.IsNone; child = child.NextSibling) {
                    if (rules.Skip(child)) continue;
                    AppendElementSignature(child, builder, rules, attributes);
                }
            }

            builder.Append("</>");
        }

        private static int CompareAttributes(WordXmlAttribute x, WordXmlAttribute y) {
            int result = string.CompareOrdinal(x.NamespaceUri, y.NamespaceUri);
            return result != 0 ? result : string.CompareOrdinal(x.LocalName, y.LocalName);
        }

        /// <summary>Changes a signature makes to the markup it writes, as the SDK path edited a copy before serializing it.</summary>
        private class SignatureRules {
            internal static readonly SignatureRules None = new();

            /// <summary>Whether the child element is left out, with its subtree.</summary>
            internal virtual bool Skip(OfficeXmlNode element) => false;

            /// <summary>Removes, adds or changes attributes of the element before they are written.</summary>
            internal virtual void Rewrite(OfficeXmlNode element, List<WordXmlAttribute> attributes) {
            }

            protected static void RemoveAll(List<WordXmlAttribute> attributes, Func<WordXmlAttribute, bool> match) => attributes.RemoveAll(attribute => match(attribute));

            /// <summary>Sets an attribute in the WordprocessingML-style typed member way: replaced when present, added otherwise.</summary>
            protected static void Set(List<WordXmlAttribute> attributes, string namespaceUri, string localName, string value) {
                for (int index = 0; index < attributes.Count; index++) {
                    if (string.Equals(attributes[index].LocalName, localName, StringComparison.Ordinal) &&
                        string.Equals(attributes[index].NamespaceUri, namespaceUri, StringComparison.Ordinal)) {
                        attributes[index] = new WordXmlAttribute(attributes[index].Prefix, localName, namespaceUri, value);
                        return;
                    }
                }

                attributes.Add(new WordXmlAttribute(string.Empty, localName, namespaceUri, value));
            }

            protected static void Remove(List<WordXmlAttribute> attributes, string namespaceUri, string localName) =>
                attributes.RemoveAll(attribute => string.Equals(attribute.LocalName, localName, StringComparison.Ordinal) &&
                                                  string.Equals(attribute.NamespaceUri, namespaceUri, StringComparison.Ordinal));
        }

        private const string RelationshipsNamespaceUri = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string DrawingNamespaceUri = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private const string WordprocessingDrawingNamespaceUri = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        private const string PictureNamespaceUri = "http://schemas.openxmlformats.org/drawingml/2006/picture";
        private const string VmlNamespaceUri = "urn:schemas-microsoft-com:vml";

        /// <summary>The markup of a drawing's identity attributes reset the way the SDK path reset them: id 0, empty name.</summary>
        private static void ResetDrawingProperties(OfficeXmlNode element, List<WordXmlAttribute> attributes) {
            if (element.Is(Wp, "docPr") || element.Is(Pic, "cNvPr")) {
                SetSignatureAttribute(attributes, string.Empty, "id", "0");
                SetSignatureAttribute(attributes, string.Empty, "name", string.Empty);
            }
        }

        private static void SetSignatureAttribute(List<WordXmlAttribute> attributes, string namespaceUri, string localName, string value) {
            for (int index = 0; index < attributes.Count; index++) {
                if (string.Equals(attributes[index].LocalName, localName, StringComparison.Ordinal) &&
                    string.Equals(attributes[index].NamespaceUri, namespaceUri, StringComparison.Ordinal)) {
                    attributes[index] = new WordXmlAttribute(attributes[index].Prefix, localName, namespaceUri, value);
                    return;
                }
            }

            attributes.Add(new WordXmlAttribute(string.Empty, localName, namespaceUri, value));
        }

        // ---- Styles ---------------------------------------------------------------------------------------------

        /// <summary>A top-level <c>w:style</c> with the facts the comparer reads.</summary>
        private sealed class ComparisonStyle {
            internal ComparisonStyle(OfficeXmlNode element) {
                Element = element;
                StyleId = element.Attribute(W, "styleId");
                Type = element.Attribute(W, "type");
                IsDefault = TryReadOnOff(element.Attribute(W, "default"));
                BasedOn = WVal(element.Child(W, "basedOn"));
            }

            internal OfficeXmlNode Element { get; }
            internal string? StyleId { get; }
            internal string? Type { get; }
            internal bool IsDefault { get; }
            internal string? BasedOn { get; }
            internal bool IsParagraphStyle => string.Equals(Type, "paragraph", StringComparison.Ordinal);

            private static bool TryReadOnOff(string? text) => text is "true" or "1" or "on";
        }

        /// <summary>The styles of one styles part, in document order, and its document defaults.</summary>
        private sealed class ComparisonStyles {
            internal static readonly ComparisonStyles Empty = new(default);

            internal ComparisonStyles(OfficeXmlNode root) {
                var styles = new List<ComparisonStyle>();
                if (!root.IsNone) {
                    foreach (OfficeXmlNode style in root.Children(W, "style")) styles.Add(new ComparisonStyle(style));
                    DocumentDefaults = root.Child(W, "docDefaults");
                }

                Definitions = styles;
            }

            internal IReadOnlyList<ComparisonStyle> Definitions { get; }

            internal OfficeXmlNode DocumentDefaults { get; }
        }
    }
}
