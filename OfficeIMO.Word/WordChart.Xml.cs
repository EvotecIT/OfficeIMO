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
        private OfficeXmlNode PlotChild(string localName) => PlotAreaOf(_chart).IsNone ? default : PlotAreaOf(_chart).Child(NsC, localName);

        /// <summary>
        /// The relationship id of the chart a drawing shows. The reference sits in the drawing's graphic data; extension
        /// markup is not searched, as the SDK read what it holds as unknown elements.
        /// </summary>
        private static string? ChartReferenceId(OfficeXmlNode drawing) {
            if (drawing.IsNone) return null;
            List<OfficeXmlNode> references = drawing.ContentDescendants(NsC, "chart");
            return references.Count == 0 ? null : references[0].Attribute(NsR, "id");
        }

        /// <summary>Whether a chart child is an axis.</summary>
        private static bool IsAxis(OfficeXmlNode element) =>
            element.Is(NsC, "catAx") || element.Is(NsC, "valAx") || element.Is(NsC, "dateAx") || element.Is(NsC, "serAx");

        /// <summary>Whether a plot area child is a chart group, such as <c>c:barChart</c>.</summary>
        private static bool IsChartGroup(OfficeXmlNode element) => element.LocalName.EndsWith("Chart", StringComparison.Ordinal);

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
