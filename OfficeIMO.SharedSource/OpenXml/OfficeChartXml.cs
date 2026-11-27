using System.Xml;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.OpenXml.Internal {
    /// <summary>
    /// Element helpers the shared chart reader and writer use over <see cref="OfficeXmlNode"/>. Typed values read the way
    /// the Open XML SDK reads them: a missing attribute is null and a value that doesn't parse throws
    /// <see cref="FormatException"/> (or <see cref="OverflowException"/>), so a chart reads the same on both paths.
    /// </summary>
    internal static class OfficeChartXml {
        internal const int NsC = OfficeXmlNamespaces.Chart;
        internal const int NsA = OfficeXmlNamespaces.DrawingML;
        internal const int NsR = OfficeXmlNamespaces.Relationships;

        internal const string ChartStyleUri = "http://schemas.microsoft.com/office/drawing/2012/chartStyle";
        internal const string C14Uri = "http://schemas.microsoft.com/office/drawing/2007/8/2/chart";
        internal const string C15Uri = "http://schemas.microsoft.com/office/drawing/2012/chart";
        internal const string C16Uri = "http://schemas.microsoft.com/office/drawing/2014/chart";
        internal const string C16R2Uri = "http://schemas.microsoft.com/office/drawing/2015/06/chart";
        internal const string C16R3Uri = "http://schemas.microsoft.com/office/drawing/2017/03/chart";

        // ---- Typed attribute reads -------------------------------------------------------------------------------

        /// <summary>The <c>val</c> attribute as written, or null.</summary>
        internal static string? Val(this OfficeXmlNode node) => node.Attribute("val");

        /// <summary>The <c>val</c> attribute of the first child with this name, or null when either is missing.</summary>
        internal static string? ChildVal(this OfficeXmlNode node, int namespaceId, string localName) =>
            node.Child(namespaceId, localName).Attribute("val");

        internal static bool? BoolAttr(this OfficeXmlNode node, string localName) {
            string? text = node.Attribute(localName);
            return text == null ? null : XmlConvert.ToBoolean(text);
        }

        internal static double? DoubleAttr(this OfficeXmlNode node, string localName) {
            string? text = node.Attribute(localName);
            return text == null ? null : XmlConvert.ToDouble(text);
        }

        internal static uint? UIntAttr(this OfficeXmlNode node, string localName) {
            string? text = node.Attribute(localName);
            return text == null ? null : XmlConvert.ToUInt32(text);
        }

        internal static int? IntAttr(this OfficeXmlNode node, string localName) {
            string? text = node.Attribute(localName);
            return text == null ? null : XmlConvert.ToInt32(text);
        }

        internal static long? LongAttr(this OfficeXmlNode node, string localName) {
            string? text = node.Attribute(localName);
            return text == null ? null : XmlConvert.ToInt64(text);
        }

        internal static ushort? UShortAttr(this OfficeXmlNode node, string localName) {
            string? text = node.Attribute(localName);
            return text == null ? null : XmlConvert.ToUInt16(text);
        }

        internal static byte? ByteAttr(this OfficeXmlNode node, string localName) {
            string? text = node.Attribute(localName);
            return text == null ? null : XmlConvert.ToByte(text);
        }

        internal static sbyte? SByteAttr(this OfficeXmlNode node, string localName) {
            string? text = node.Attribute(localName);
            return text == null ? null : XmlConvert.ToSByte(text);
        }

        internal static bool? BoolVal(this OfficeXmlNode node) => node.BoolAttr("val");
        internal static double? DoubleVal(this OfficeXmlNode node) => node.DoubleAttr("val");
        internal static uint? UIntVal(this OfficeXmlNode node) => node.UIntAttr("val");
        internal static int? IntVal(this OfficeXmlNode node) => node.IntAttr("val");
        internal static long? LongVal(this OfficeXmlNode node) => node.LongAttr("val");
        internal static ushort? UShortVal(this OfficeXmlNode node) => node.UShortAttr("val");
        internal static byte? ByteVal(this OfficeXmlNode node) => node.ByteAttr("val");
        internal static sbyte? SByteVal(this OfficeXmlNode node) => node.SByteAttr("val");

        // ---- Value text, as the SDK writes it ---------------------------------------------------------------------

        internal static string Text(bool value) => value ? "1" : "0";
        internal static string Text(double value) => OfficeXmlViewValues.Format(value)!;
        internal static string Text(int value) => value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        internal static string Text(uint value) => value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        internal static string Text(long value) => value.ToString(System.Globalization.CultureInfo.InvariantCulture);

        // ---- Structure ----------------------------------------------------------------------------------------------

        /// <summary>Whether the element has child elements.</summary>
        internal static bool HasChildElements(this OfficeXmlNode node) => !node.FirstChild.IsNone;

        /// <summary>Creates a child element with this name and appends it.</summary>
        internal static OfficeXmlNode AppendNew(this OfficeXmlNode parent, int namespaceId, string localName) =>
            parent.Append(parent.CreateElement(namespaceId, localName));

        /// <summary>Creates a child element with a <c>val</c> attribute and appends it.</summary>
        internal static OfficeXmlNode AppendVal(this OfficeXmlNode parent, int namespaceId, string localName, string value) {
            OfficeXmlNode child = parent.Append(parent.CreateElement(namespaceId, localName));
            child.SetAttribute("val", value);
            return child;
        }

        /// <summary>Creates a detached element with a <c>val</c> attribute, built in <paramref name="context"/>'s part.</summary>
        internal static OfficeXmlNode CreateVal(this OfficeXmlNode context, int namespaceId, string localName, string value) {
            OfficeXmlNode child = context.CreateElement(namespaceId, localName);
            child.SetAttribute("val", value);
            return child;
        }

        /// <summary>
        /// Places a child as the SDK's <c>AddChild(child, true)</c> and typed property setters do: every child in the same
        /// schema slot (each member of a choice, each occurrence of a repeated element) is removed, then the child goes in
        /// schema order. Returns the child as it is now held.
        /// </summary>
        internal static OfficeXmlNode SetOrdered(OfficeXmlNode parent, OfficeXmlNode child, OfficeXmlChildOrder order) {
            int rank = order.RankOf(child.NamespaceId, child.LocalName);
            if (rank >= 0) {
                OfficeXmlNode sibling = parent.FirstChild;
                while (!sibling.IsNone) {
                    OfficeXmlNode next = sibling.NextSibling;
                    if (order.RankOf(sibling.NamespaceId, sibling.LocalName) == rank) sibling.Remove();
                    sibling = next;
                }
            }
            return parent.InsertOrdered(child, order);
        }

        /// <summary>Removes every child element.</summary>
        internal static void RemoveAllChildElements(this OfficeXmlNode node) {
            OfficeXmlNode child = node.FirstChild;
            while (!child.IsNone) {
                OfficeXmlNode next = child.NextSibling;
                child.Remove();
                child = next;
            }
        }

        /// <summary>Replaces the element with <paramref name="replacement"/> and returns it as it is now held.</summary>
        internal static OfficeXmlNode ReplaceWith(this OfficeXmlNode node, OfficeXmlNode replacement) {
            OfficeXmlNode parent = node.Parent;
            OfficeXmlNode inserted = parent.InsertBefore(replacement, node);
            node.Remove();
            return inserted;
        }

        /// <summary>
        /// The element after <paramref name="node"/> in document order within <paramref name="root"/>'s subtree, or none.
        /// Walks without recursion, so deeply nested markup can't exhaust the stack.
        /// </summary>
        internal static OfficeXmlNode NextInSubtree(this OfficeXmlNode node, OfficeXmlNode root) {
            OfficeXmlNode child = node.FirstChild;
            if (!child.IsNone) return child;
            for (OfficeXmlNode current = node; !current.IsNone && current != root; current = current.Parent) {
                OfficeXmlNode sibling = current.NextSibling;
                if (!sibling.IsNone) return sibling;
            }
            return default;
        }

        /// <summary>The first element with this name below this one, in document order, or none.</summary>
        internal static OfficeXmlNode FirstDescendant(this OfficeXmlNode node, int namespaceId, string localName) {
            for (OfficeXmlNode current = node.FirstChild; !current.IsNone; current = current.NextInSubtree(node)) {
                if (current.Is(namespaceId, localName)) return current;
            }
            return default;
        }

        /// <summary>Every element below this one, in document order.</summary>
        internal static List<OfficeXmlNode> AllDescendants(this OfficeXmlNode node) {
            var found = new List<OfficeXmlNode>();
            CollectDescendants(node, null, found);
            return found;
        }

        /// <summary>The elements below this one that match, in document order.</summary>
        internal static List<OfficeXmlNode> DescendantsWhere(this OfficeXmlNode node, Func<OfficeXmlNode, bool> predicate) {
            var found = new List<OfficeXmlNode>();
            CollectDescendants(node, predicate, found);
            return found;
        }

        private static void CollectDescendants(OfficeXmlNode node, Func<OfficeXmlNode, bool>? predicate, List<OfficeXmlNode> found) {
            for (OfficeXmlNode current = node.FirstChild; !current.IsNone; current = current.NextInSubtree(node)) {
                if (predicate == null || predicate(current)) found.Add(current);
            }
        }

        /// <summary>The nearest ancestor with this name, the element itself excluded, or none.</summary>
        internal static OfficeXmlNode Ancestor(this OfficeXmlNode node, int namespaceId, string localName) =>
            node.IsNone ? default : node.Parent.AncestorOrSelf(namespaceId, localName);

        /// <summary>Whether the element is the <c>c:</c> element with this name.</summary>
        internal static bool IsChart(this OfficeXmlNode node, string localName) => node.Is(NsC, localName);

        /// <summary>Whether the element is the <c>a:</c> element with this name.</summary>
        internal static bool IsDrawing(this OfficeXmlNode node, string localName) => node.Is(NsA, localName);
    }
}
