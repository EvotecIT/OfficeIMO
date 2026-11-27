using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel.OpenXml {
    /// <summary>
    /// The element idioms chart, drawing and image code uses over <see cref="OfficeXmlNode"/>: the <c>val</c> child
    /// elements DrawingML charts store most settings in, building small subtrees, and the namespaces those parts use.
    /// Values are written as the schema spells them, booleans as <c>1</c> and <c>0</c>, as the Open XML SDK writes them.
    /// </summary>
    internal static class ExcelChartXml {
        internal const int C = OfficeXmlNamespaces.Chart;
        internal const int A = OfficeXmlNamespaces.DrawingML;
        internal const int Xdr = OfficeXmlNamespaces.SpreadsheetDrawing;
        internal const int R = OfficeXmlNamespaces.Relationships;
        internal const int Pic = OfficeXmlNamespaces.Picture;
        internal const int Cdr = OfficeXmlNamespaces.ChartDrawing;
        internal const int Mc = OfficeXmlNamespaces.MarkupCompatibility;

        internal const string ChartUri = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        internal const string DrawingUri = "http://schemas.openxmlformats.org/drawingml/2006/main";
        internal const string SpreadsheetDrawingUri = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
        internal const string RelationshipsUri = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        internal const string C14Uri = "http://schemas.microsoft.com/office/drawing/2007/8/2/chart";
        internal const string C15Uri = "http://schemas.microsoft.com/office/drawing/2012/chart";
        internal const string C16Uri = "http://schemas.microsoft.com/office/drawing/2014/chart";
        internal const string C16R2Uri = "http://schemas.microsoft.com/office/drawing/2015/06/chart";
        internal const string C16R3Uri = "http://schemas.microsoft.com/office/drawing/2017/03/chart";
        internal const string ChartExUri = "http://schemas.microsoft.com/office/drawing/2014/chartex";
        internal const string ChartStyleUri = "http://schemas.microsoft.com/office/drawing/2012/chartStyle";
        internal const string A14Uri = "http://schemas.microsoft.com/office/drawing/2010/main";
        internal const string A16Uri = "http://schemas.microsoft.com/office/drawing/2014/main";

        // ---- val children ---------------------------------------------------------------------------------------



        internal static int? Int32Val(this OfficeXmlNode element) => OfficeXmlViewValues.ParseInt32(element.Val());
        internal static uint? UInt32Val(this OfficeXmlNode element) => OfficeXmlViewValues.ParseUInt32(element.Val());
        internal static long? Int64Val(this OfficeXmlNode element) => OfficeXmlViewValues.ParseInt64(element.Val());

        internal static bool? ChildBoolVal(this OfficeXmlNode parent, int namespaceId, string localName) =>
            OfficeXmlViewValues.ParseBoolean(parent.ChildVal(namespaceId, localName));
        internal static int? ChildInt32Val(this OfficeXmlNode parent, int namespaceId, string localName) =>
            OfficeXmlViewValues.ParseInt32(parent.ChildVal(namespaceId, localName));
        internal static uint? ChildUInt32Val(this OfficeXmlNode parent, int namespaceId, string localName) =>
            OfficeXmlViewValues.ParseUInt32(parent.ChildVal(namespaceId, localName));
        internal static double? ChildDoubleVal(this OfficeXmlNode parent, int namespaceId, string localName) =>
            OfficeXmlViewValues.ParseDouble(parent.ChildVal(namespaceId, localName));

        /// <summary>Sets the <c>val</c> attribute; null removes it.</summary>
        internal static OfficeXmlNode SetVal(this OfficeXmlNode element, string? value) {
            element.SetAttribute("val", value);
            return element;
        }

        internal static OfficeXmlNode SetVal(this OfficeXmlNode element, bool? value) => element.SetVal(OfficeXmlViewValues.Format(value));
        internal static OfficeXmlNode SetVal(this OfficeXmlNode element, int? value) => element.SetVal(OfficeXmlViewValues.Format(value));
        internal static OfficeXmlNode SetVal(this OfficeXmlNode element, uint? value) => element.SetVal(OfficeXmlViewValues.Format(value));
        internal static OfficeXmlNode SetVal(this OfficeXmlNode element, long? value) => element.SetVal(OfficeXmlViewValues.Format(value));
        internal static OfficeXmlNode SetVal(this OfficeXmlNode element, double? value) => element.SetVal(OfficeXmlViewValues.Format(value));

        // ---- Building -------------------------------------------------------------------------------------------

        /// <summary>A new detached element of this part, to fill and insert, as <c>new C.X()</c> makes one.</summary>
        internal static OfficeXmlNode New(this OfficeXmlNode context, int namespaceId, string localName) =>
            context.CreateElement(namespaceId, localName);

        /// <summary>A new detached element with a <c>val</c> attribute, as <c>new C.X { Val = ... }</c> makes one.</summary>
        internal static OfficeXmlNode NewVal(this OfficeXmlNode context, int namespaceId, string localName, string? value) =>
            context.CreateElement(namespaceId, localName).SetVal(value);

        internal static OfficeXmlNode NewVal(this OfficeXmlNode context, int namespaceId, string localName, bool value) =>
            context.CreateElement(namespaceId, localName).SetVal(value);

        internal static OfficeXmlNode NewVal(this OfficeXmlNode context, int namespaceId, string localName, int value) =>
            context.CreateElement(namespaceId, localName).SetVal(value);

        internal static OfficeXmlNode NewVal(this OfficeXmlNode context, int namespaceId, string localName, uint value) =>
            context.CreateElement(namespaceId, localName).SetVal(value);

        internal static OfficeXmlNode NewVal(this OfficeXmlNode context, int namespaceId, string localName, long value) =>
            context.CreateElement(namespaceId, localName).SetVal(value);

        internal static OfficeXmlNode NewVal(this OfficeXmlNode context, int namespaceId, string localName, double value) =>
            context.CreateElement(namespaceId, localName).SetVal(value);

        /// <summary>A new detached element with text content, as <c>new C.Formula(text)</c> makes one.</summary>
        internal static OfficeXmlNode NewText(this OfficeXmlNode context, int namespaceId, string localName, string text) {
            OfficeXmlNode element = context.CreateElement(namespaceId, localName);
            element.SetInnerText(text);
            return element;
        }

        /// <summary>Appends detached elements in order and returns this element, as the SDK's child constructors do.</summary>
        internal static OfficeXmlNode With(this OfficeXmlNode element, params OfficeXmlNode[] children) {
            foreach (OfficeXmlNode child in children) {
                if (!child.IsNone) element.Append(child);
            }
            return element;
        }

        /// <summary>Sets an attribute without a namespace and returns this element.</summary>
        internal static OfficeXmlNode Attr(this OfficeXmlNode element, string localName, string? value) {
            element.SetAttribute(localName, value);
            return element;
        }

        internal static OfficeXmlNode Attr(this OfficeXmlNode element, string localName, bool? value) => element.Attr(localName, OfficeXmlViewValues.Format(value));
        internal static OfficeXmlNode Attr(this OfficeXmlNode element, string localName, int? value) => element.Attr(localName, OfficeXmlViewValues.Format(value));
        internal static OfficeXmlNode Attr(this OfficeXmlNode element, string localName, uint? value) => element.Attr(localName, OfficeXmlViewValues.Format(value));
        internal static OfficeXmlNode Attr(this OfficeXmlNode element, string localName, long? value) => element.Attr(localName, OfficeXmlViewValues.Format(value));
        internal static OfficeXmlNode Attr(this OfficeXmlNode element, string localName, double? value) => element.Attr(localName, OfficeXmlViewValues.Format(value));

        // ---- Editing --------------------------------------------------------------------------------------------

        /// <summary>
        /// Puts <paramref name="child"/> where the schema orders it, replacing an existing child of the same name, as an
        /// SDK typed property setter does (<c>chart.AutoTitleDeleted = new ...</c>). Returns the child as held.
        /// </summary>
        internal static OfficeXmlNode SetOrdered(this OfficeXmlNode parent, OfficeXmlNode child, OfficeXmlChildOrder order) =>
            OfficeChartXml.SetOrdered(parent, child, order);

        /// <summary>
        /// Adds <paramref name="child"/> where the schema orders it, after existing children of the same name, as the
        /// SDK's <c>AddChild(child, true)</c> does. Returns the child as held.
        /// </summary>
        internal static OfficeXmlNode AddOrdered(this OfficeXmlNode parent, OfficeXmlNode child, OfficeXmlChildOrder order) =>
            parent.InsertOrdered(child, order);

        /// <summary>Removes the first child with this name, when there is one.</summary>
        internal static void RemoveChild(this OfficeXmlNode parent, int namespaceId, string localName) {
            OfficeXmlNode child = parent.Child(namespaceId, localName);
            if (!child.IsNone) child.Remove();
        }

        /// <summary>Removes every child element.</summary>
        internal static void RemoveAllChildren(this OfficeXmlNode parent) {
            OfficeXmlNode child = parent.FirstChild;
            while (!child.IsNone) {
                OfficeXmlNode next = child.NextSibling;
                child.Remove();
                child = next;
            }
        }


        /// <summary>The number of child elements.</summary>
        internal static int ChildCount(this OfficeXmlNode parent) => parent.Children().Count;

        /// <summary>The zero-based position of this element among its parent's child elements, or -1.</summary>
        internal static int IndexInParent(this OfficeXmlNode element) {
            OfficeXmlNode parent = element.Parent;
            if (parent.IsNone) return -1;
            int index = 0;
            for (OfficeXmlNode sibling = parent.FirstChild; !sibling.IsNone; sibling = sibling.NextSibling) {
                if (sibling.Equals(element)) return index;
                index++;
            }
            return -1;
        }



        /// <summary>The text of a child element as an invariant integer, as SDK marker elements read.</summary>
        internal static int? ChildInt32Text(this OfficeXmlNode parent, int namespaceId, string localName) {
            string? text = parent.ChildText(namespaceId, localName);
            return text != null && int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value) ? value : null;
        }

        /// <summary>The <c>r:id</c>, <c>r:embed</c> or other relationship attribute of an element.</summary>
        internal static string? RelationshipAttribute(this OfficeXmlNode element, string localName) =>
            element.IsNone ? null : element.Attribute(R, localName);

        /// <summary>
        /// The element after <paramref name="node"/> in document order within <paramref name="root"/>'s subtree, never entering
        /// an extension list. The SDK reads unknown extension content as unknown elements, so its typed lookups never matched
        /// markup there. Unlike the shared chart walk it still enters mc:AlternateContent, whose choices the SDK reads as typed
        /// elements. Walks without recursion, so deeply nested markup can't exhaust the stack.
        /// </summary>
        internal static OfficeXmlNode NextOutsideExtensions(this OfficeXmlNode node, OfficeXmlNode root) {
            if (!IsExtensionList(node)) {
                OfficeXmlNode child = node.FirstChild;
                if (!child.IsNone) return child;
            }
            for (OfficeXmlNode current = node; !current.IsNone && current != root; current = current.Parent) {
                OfficeXmlNode sibling = current.NextSibling;
                if (!sibling.IsNone) return sibling;
            }
            return default;
        }

        /// <summary>The first element below this one, in document order and outside extension lists, that matches, or none.</summary>
        internal static OfficeXmlNode FirstDescendantOutsideExtensions(this OfficeXmlNode node, Func<OfficeXmlNode, bool> match) {
            if (node.IsNone) return default;
            for (OfficeXmlNode current = node.FirstChild; !current.IsNone; current = current.NextOutsideExtensions(node)) {
                if (!IsExtensionList(current) && match(current)) return current;
            }
            return default;
        }

        /// <summary>The first element with this name below this one, in document order and outside extension lists, or none.</summary>
        internal static OfficeXmlNode FirstDescendantOutsideExtensions(this OfficeXmlNode node, int namespaceId, string localName) =>
            node.FirstDescendantOutsideExtensions(current => current.Is(namespaceId, localName));

        /// <summary>The elements with this name below this one, in document order and outside extension lists.</summary>
        internal static List<OfficeXmlNode> DescendantsOutsideExtensions(this OfficeXmlNode node, int namespaceId, string localName) {
            var found = new List<OfficeXmlNode>();
            if (node.IsNone) return found;
            for (OfficeXmlNode current = node.FirstChild; !current.IsNone; current = current.NextOutsideExtensions(node)) {
                if (current.Is(namespaceId, localName)) found.Add(current);
            }
            return found;
        }

        private static bool IsExtensionList(OfficeXmlNode node) => string.Equals(node.LocalName, "extLst", StringComparison.Ordinal);
    }
}
