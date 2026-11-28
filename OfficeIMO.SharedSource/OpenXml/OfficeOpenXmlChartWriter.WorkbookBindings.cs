using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal.OpenXml;
using static OfficeIMO.OpenXml.Internal.OfficeChartXml;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartWriter {
        /// <summary>
        /// Prepares titles, custom labels and custom error bars for replacing the embedded worksheet.
        /// Every binding is checked before detaching cached content from its old workbook formula.
        /// </summary>
        internal static Action PrepareSharedWorkbookBindings(IOfficeChartPart part, IReadOnlyList<int> replacementPointCounts) {
            _ = GetSharedWorkbookBindings(part, replacementPointCounts);
            return () => {
                // Native updates can clone preserved series and axes. Resolve the current
                // text nodes after that update so cloned bindings are materialized too.
                foreach (var replacement in GetSharedWorkbookBindings(part, replacementPointCounts)) {
                    replacement.Key.ReplaceWith(replacement.Value);
                }
            };
        }

        private static List<KeyValuePair<OfficeXmlNode, OfficeXmlNode>> GetSharedWorkbookBindings(
            IOfficeChartPart part, IReadOnlyList<int> replacementPointCounts) {
            OfficeXmlNode chartSpace = part.ChartSpace;
            if (chartSpace.IsNone) throw new InvalidOperationException("Chart space not found.");
            var replacements = new List<KeyValuePair<OfficeXmlNode, OfficeXmlNode>>();
            List<OfficeXmlNode> descendants = chartSpace.DescendantsWhere(null, includeExtensionMarkup: true);
            List<OfficeXmlNode> sourceSeries = descendants.Where(element => element.LocalName == "ser").ToList();
            foreach (OfficeXmlNode text in descendants) {
                if (!IsChartText(text)) continue;
                OfficeXmlNode reference = text.Child(NsC, "strRef");
                if (reference.IsNone) continue;
                OfficeXmlNode cache = reference.Child(NsC, "strCache");
                List<OfficeXmlNode> points = cache.IsNone ? new List<OfficeXmlNode>() : TakeChildren(cache, "pt", 2);
                if (cache.IsNone || points.Count != 1 || cache.Child(NsC, "ptCount").UIntVal() != 1 ||
                    points[0].UIntAttr("idx") != 0 || points[0].Child(NsC, "v").IsNone)
                    throw new NotSupportedException("Replacing a chart workbook requires a single cached value for every formula-linked title or custom label; the existing chart and workbook are preserved.");
                OfficeXmlNode rich = text.CreateElement(NsC, "rich");
                rich.AppendNew(NsA, "bodyPr");
                rich.AppendNew(NsA, "lstStyle");
                rich.AppendNew(NsA, "p").AppendNew(NsA, "r").AppendNew(NsA, "t").SetInnerText(points[0].Child(NsC, "v").InnerText);
                replacements.Add(new KeyValuePair<OfficeXmlNode, OfficeXmlNode>(reference, rich));
            }
            int totalErrorPoints = 0;
            foreach (OfficeXmlNode reference in descendants) {
                if (!reference.Is(NsC, "numRef")) continue;
                OfficeXmlNode errorSide = reference.Parent;
                if (!errorSide.Is(NsC, "plus") && !errorSide.Is(NsC, "minus")) continue;
                OfficeXmlNode cache = reference.Child(NsC, "numCache");
                List<OfficeXmlNode> points = cache.IsNone ? new List<OfficeXmlNode>() : TakeChildren(cache, "pt", MaximumSharedChartPoints + 1);
                totalErrorPoints += points.Count;
                uint? count = cache.Child(NsC, "ptCount").UIntVal();
                OfficeXmlNode ownerSeries = default;
                for (OfficeXmlNode ancestor = reference.Parent; !ancestor.IsNone; ancestor = ancestor.Parent) {
                    if (ancestor.LocalName == "ser") {
                        ownerSeries = ancestor;
                        break;
                    }
                }
                int seriesIndex = ownerSeries.IsNone ? -1 : sourceSeries.IndexOf(ownerSeries);
                if (seriesIndex < 0 || seriesIndex >= replacementPointCounts.Count ||
                    count != (uint)replacementPointCounts[seriesIndex])
                    throw new NotSupportedException("Custom error-bar caches must match the replacement series point count.");
                if (cache.IsNone || !count.HasValue || totalErrorPoints > MaximumSharedChartPoints ||
                    count.Value > MaximumSharedChartPoints || (uint)points.Count != count.Value)
                    throw new NotSupportedException("Replacing a chart workbook requires bounded cached values for custom error bars.");
                var indexes = new HashSet<uint>();
                foreach (OfficeXmlNode point in points) {
                    uint? index = point.UIntAttr("idx");
                    if (!index.HasValue || index.Value >= count!.Value || !indexes.Add(index.Value))
                        throw new NotSupportedException("Custom error-bar caches must have bounded, unique point indexes.");
                    OfficeXmlNode valueElement = point.Child(NsC, "v");
                    if (!OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(valueElement.IsNone ? null : valueElement.InnerText, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) ||
                        double.IsNaN(value) || double.IsInfinity(value))
                        throw new NotSupportedException("Custom error-bar caches must provide a finite numeric value for every declared point.");
                }
                OfficeXmlNode literal = errorSide.CreateElement(NsC, "numLit");
                foreach (OfficeXmlNode child in cache.Children()) literal.Append(child.Clone());
                if (literal.Child(NsC, "formatCode").IsNone) {
                    OfficeXmlNode formatCode = literal.CreateElement(NsC, "formatCode");
                    formatCode.SetInnerText("General");
                    SetOrdered(literal, formatCode, CNumDataView.Order);
                }
                replacements.Add(new KeyValuePair<OfficeXmlNode, OfficeXmlNode>(reference, literal));
            }
            foreach (OfficeXmlNode formula in descendants) {
                if (formula.LocalName != "f") continue;
                bool rewrittenSource = false;
                for (OfficeXmlNode ancestor = formula.Parent; !ancestor.IsNone; ancestor = ancestor.Parent) {
                    if (IsRewrittenSeriesSource(ancestor)) {
                        rewrittenSource = true;
                        break;
                    }
                }
                OfficeXmlNode parent = formula.Parent;
                bool preservedText = parent.Is(NsC, "strRef") && IsChartText(parent.Parent);
                OfficeXmlNode grandparent = parent.Parent;
                bool preservedError = parent.Is(NsC, "numRef") && (grandparent.Is(NsC, "plus") || grandparent.Is(NsC, "minus"));
                if (!rewrittenSource && !preservedText && !preservedError)
                    throw new NotSupportedException("The chart contains a workbook-linked extension that cannot be preserved while replacing the embedded worksheet.");
            }
            return replacements;
        }

        private static List<OfficeXmlNode> TakeChildren(OfficeXmlNode parent, string localName, int maximum) {
            var children = new List<OfficeXmlNode>();
            foreach (OfficeXmlNode child in parent.Children(NsC, localName)) {
                if (children.Count == maximum) break;
                children.Add(child);
            }
            return children;
        }

        // The c:tx of a title, data label, trendline label or display units label (CT_Tx); a series' c:tx is CT_SerTx.
        private static bool IsChartText(OfficeXmlNode element) {
            if (!element.Is(NsC, "tx")) return false;
            OfficeXmlNode owner = element.Parent;
            return owner.NamespaceId == NsC && owner.LocalName is "title" or "dLbl" or "trendlineLbl" or "dispUnitsLbl";
        }

        // The series name, category, value and bubble size sources the replacement workbook rewrites: direct children of
        // a series of a chart group in the plot area, where the SDK typed them. Same-named elements kept inside extension
        // content were unknown elements to the SDK, and the replacement leaves them as they are.
        private static bool IsRewrittenSeriesSource(OfficeXmlNode element) {
            if (element.NamespaceId != NsC || element.LocalName is not ("tx" or "cat" or "val" or "xVal" or "yVal" or "bubbleSize")) return false;
            OfficeXmlNode series = element.Parent;
            return series.Is(NsC, "ser") && series.Parent.Parent.Is(NsC, "plotArea");
        }

        // SDK callers
        internal static Action PrepareSharedWorkbookBindings(ChartPart part, IReadOnlyList<int> replacementPointCounts) =>
            PrepareSharedWorkbookBindings(new OfficeSdkChartPart(part), replacementPointCounts);
    }
}
