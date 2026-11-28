using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word {
    /// <summary>
    /// Element helpers for chart markup. New elements are always created from the element they go under, so an SDK-held
    /// chart gets the same typed elements it got when the chart was built from SDK classes.
    /// </summary>
    public partial class WordChart {
        private const int NsC = OfficeXmlNamespaces.Chart;
        private const int NsA = OfficeXmlNamespaces.DrawingML;
        private const int NsWp = OfficeXmlNamespaces.WordprocessingDrawing;
        private const int NsR = OfficeXmlNamespaces.Relationships;
        private const int NsMc = OfficeXmlNamespaces.MarkupCompatibility;
        private const string ChartNamespaceUri = "http://schemas.openxmlformats.org/drawingml/2006/chart";

        /// <summary>Creates a child of <paramref name="parent"/> and appends it.</summary>
        private static OfficeXmlNode AddC(OfficeXmlNode parent, string localName) => parent.AppendNew(NsC, localName);

        /// <summary>Creates a child of <paramref name="parent"/> with a <c>val</c> attribute and appends it.</summary>
        private static OfficeXmlNode AddC(OfficeXmlNode parent, string localName, string value) => parent.AppendVal(NsC, localName, value);

        private static OfficeXmlNode AddC(OfficeXmlNode parent, string localName, bool value) => parent.AppendVal(NsC, localName, OfficeChartXml.Text(value));

        private static OfficeXmlNode AddC(OfficeXmlNode parent, string localName, uint value) => parent.AppendVal(NsC, localName, OfficeChartXml.Text(value));

        private static OfficeXmlNode AddA(OfficeXmlNode parent, string localName) => parent.AppendNew(NsA, localName);

        /// <summary>Creates a detached child for <paramref name="context"/>, the element it will go under.</summary>
        private static OfficeXmlNode NewC(OfficeXmlNode context, string localName) => context.CreateElement(NsC, localName);

        /// <summary>Creates a detached child for <paramref name="context"/> with a <c>val</c> attribute.</summary>
        private static OfficeXmlNode NewC(OfficeXmlNode context, string localName, string value) => context.CreateVal(NsC, localName, value);

        private static OfficeXmlNode NewC(OfficeXmlNode context, string localName, bool value) => context.CreateVal(NsC, localName, OfficeChartXml.Text(value));

        private static OfficeXmlNode NewC(OfficeXmlNode context, string localName, uint value) => context.CreateVal(NsC, localName, OfficeChartXml.Text(value));

        /// <summary>Appends a leaf element holding <paramref name="text"/>, such as <c>c:v</c> or <c>c:formatCode</c>.</summary>
        private static OfficeXmlNode AddText(OfficeXmlNode parent, int namespaceId, string localName, string text) {
            OfficeXmlNode child = parent.AppendNew(namespaceId, localName);
            child.SetInnerText(text);
            return child;
        }

        /// <summary>The chart's <c>c:plotArea</c>, or none.</summary>
        private static OfficeXmlNode PlotAreaOf(OfficeXmlNode chart) => chart.IsNone ? default : chart.Child(NsC, "plotArea");

        /// <summary>The first child of the chart's plot area with this name, or none.</summary>
        private OfficeXmlNode PlotChild(string localName) {
            OfficeXmlNode plot = PlotAreaOf(_chart);
            return plot.IsNone ? default : plot.Child(NsC, localName);
        }

        /// <summary>
        /// The relationship id of the chart a drawing shows: the <c>c:chart</c> at
        /// wp:inline or wp:anchor/a:graphic/a:graphicData/c:chart of the drawing's own frame. An <c>mc:AlternateContent</c>
        /// at any step contributes the branch a first-edition consumer selects. A chart anywhere else, such as in vendor
        /// or extension markup, is not.
        /// </summary>
        private static string? ChartReferenceId(OfficeXmlNode drawing) {
            OfficeXmlNode graphic = SelectedChild(DrawingFrame(drawing), NsA, "graphic");
            OfficeXmlNode reference = SelectedChild(SelectedChild(graphic, NsA, "graphicData"), NsC, "chart");
            return reference.IsNone ? null : reference.Attribute(NsR, "id");
        }

        /// <summary>The drawing's wp:inline or wp:anchor frame, looking through <c>mc:AlternateContent</c>; or none.</summary>
        private static OfficeXmlNode DrawingFrame(OfficeXmlNode drawing) {
            for (OfficeXmlNode child = drawing.IsNone ? default : drawing.FirstChild; !child.IsNone; child = child.NextSibling) {
                if (child.Is(NsWp, "inline") || child.Is(NsWp, "anchor")) return child;
                if (!child.Is(NsMc, "AlternateContent")) continue;
                OfficeXmlNode frame = DrawingFrame(SelectedBranch(child));
                if (!frame.IsNone) return frame;
            }
            return default;
        }

        /// <summary>The first child with this name, looking through <c>mc:AlternateContent</c>; or none.</summary>
        private static OfficeXmlNode SelectedChild(OfficeXmlNode parent, int namespaceId, string localName) {
            for (OfficeXmlNode child = parent.IsNone ? default : parent.FirstChild; !child.IsNone; child = child.NextSibling) {
                if (child.Is(namespaceId, localName)) return child;
                if (!child.Is(NsMc, "AlternateContent")) continue;
                OfficeXmlNode found = SelectedChild(SelectedBranch(child), namespaceId, localName);
                if (!found.IsNone) return found;
            }
            return default;
        }

        /// <summary>
        /// The branch of an <c>mc:AlternateContent</c> a first-edition consumer selects: the first <c>mc:Choice</c> whose
        /// <c>Requires</c> namespaces are all ECMA-376 namespaces, otherwise the <c>mc:Fallback</c>; or none.
        /// </summary>
        private static OfficeXmlNode SelectedBranch(OfficeXmlNode alternate) {
            for (OfficeXmlNode branch = alternate.FirstChild; !branch.IsNone; branch = branch.NextSibling) {
                if (branch.Is(NsMc, "Choice") && RequirementsUnderstood(branch)) return branch;
                if (branch.Is(NsMc, "Fallback")) return branch;
            }
            return default;
        }

        private static bool RequirementsUnderstood(OfficeXmlNode choice) {
            string? requires = choice.Attribute(OfficeXmlNamespaces.None, "Requires");
            if (string.IsNullOrWhiteSpace(requires)) return false;
            foreach (string prefix in requires!.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)) {
                string? uri = choice.LookupNamespaceUri(prefix);
                if (uri == null || !OfficeXmlNamespaces.TryGetKnownId(uri, out _)) return false;
            }
            return true;
        }

        /// <summary>
        /// A series title's text: the literal <c>c:v</c> of its <c>c:tx</c>, or the first cached value of its string
        /// reference. Values anywhere else, such as in vendor markup, are not the title.
        /// </summary>
        internal static string? SeriesTitleText(OfficeXmlNode text) {
            if (text.IsNone) return null;
            OfficeXmlNode literal = text.Child(NsC, "v");
            if (!literal.IsNone) return literal.InnerText;
            OfficeXmlNode reference = text.Child(NsC, "strRef");
            OfficeXmlNode cache = reference.IsNone ? default : reference.Child(NsC, "strCache");
            if (cache.IsNone) return null;
            foreach (OfficeXmlNode point in cache.Children(NsC, "pt")) {
                OfficeXmlNode value = point.Child(NsC, "v");
                if (!value.IsNone) return value.InnerText;
            }
            return null;
        }

        /// <summary>Whether a chart child is an axis.</summary>
        private static bool IsAxis(OfficeXmlNode element) =>
            element.Is(NsC, "catAx") || element.Is(NsC, "valAx") || element.Is(NsC, "dateAx") || element.Is(NsC, "serAx");

        /// <summary>Whether a plot area child is a chart group, such as <c>c:barChart</c>.</summary>
        private static bool IsChartGroup(OfficeXmlNode element) =>
            element.NamespaceId == NsC && element.LocalName.EndsWith("Chart", StringComparison.Ordinal);

        /// <summary>The plotted series: the <c>c:ser</c> children of the plot area's chart groups, in document order.</summary>
        private static IEnumerable<OfficeXmlNode> PlottedSeries(OfficeXmlNode plot) {
            for (OfficeXmlNode group = plot.IsNone ? default : plot.FirstChild; !group.IsNone; group = group.NextSibling) {
                if (!IsChartGroup(group)) continue;
                foreach (OfficeXmlNode series in group.Children(NsC, "ser")) yield return series;
            }
        }

        /// <summary>Raises the shared axis id seed to at least <paramref name="max"/> without lowering it.</summary>
        private static void RaiseAxisIdSeed(uint max) {
            // Another document may allocate axes while this document is scanned.
            // Raise the shared seed without overwriting those allocations.
            int current = System.Threading.Volatile.Read(ref _axisIdSeed);
            while (max > (uint)current) {
                int observed = System.Threading.Interlocked.CompareExchange(ref _axisIdSeed, (int)max, current);
                if (observed == current) break;
                current = observed;
            }
        }
    }
}
