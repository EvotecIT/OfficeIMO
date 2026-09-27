using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartWriter {
        /// <summary>
        /// Prepares titles, custom labels and custom error bars for replacing the embedded worksheet.
        /// Every binding is checked before detaching cached content from its old workbook formula.
        /// </summary>
        internal static Action PrepareSharedWorkbookBindings(ChartPart part) {
            _ = GetSharedWorkbookBindings(part);
            return () => {
                // Native updates can clone preserved series and axes. Resolve the current
                // text nodes after that update so cloned bindings are materialized too.
                foreach (var replacement in GetSharedWorkbookBindings(part)) {
                    replacement.Key.Parent!.ReplaceChild(replacement.Value, replacement.Key);
                }
            };
        }

        private static List<KeyValuePair<OpenXmlElement, OpenXmlElement>> GetSharedWorkbookBindings(ChartPart part) {
            var replacements = new List<KeyValuePair<OpenXmlElement, OpenXmlElement>>();
            foreach (C.ChartText text in part.ChartSpace!.Descendants<C.ChartText>()) {
                C.StringReference? reference = text.GetFirstChild<C.StringReference>();
                if (reference == null) continue;
                C.StringCache? cache = reference.GetFirstChild<C.StringCache>();
                List<C.StringPoint> points = cache?.Elements<C.StringPoint>().Take(2).ToList() ?? new List<C.StringPoint>();
                if (cache == null || points.Count != 1 || cache.PointCount?.Val?.Value != 1 ||
                    points[0].Index?.Value != 0 || points[0].NumericValue == null)
                    throw new NotSupportedException("Replacing a chart workbook requires a single cached value for every formula-linked title or custom label; the existing chart and workbook are preserved.");
                var rich = new C.RichText(new A.BodyProperties(), new A.ListStyle(),
                    new A.Paragraph(new A.Run(new A.Text(points[0].NumericValue!.Text ?? string.Empty))));
                replacements.Add(new KeyValuePair<OpenXmlElement, OpenXmlElement>(reference, rich));
            }
            int totalErrorPoints = 0;
            foreach (C.NumberReference reference in part.ChartSpace.Descendants<C.NumberReference>()) {
                if (!(reference.Parent is C.Plus) && !(reference.Parent is C.Minus)) continue;
                C.NumberingCache? cache = reference.NumberingCache;
                List<C.NumericPoint> points = cache?.Elements<C.NumericPoint>().Take(MaximumSharedChartPoints + 1).ToList() ?? new List<C.NumericPoint>();
                totalErrorPoints += points.Count;
                uint? count = cache?.PointCount?.Val?.Value;
                if (cache == null || !count.HasValue || totalErrorPoints > MaximumSharedChartPoints ||
                    count.Value > MaximumSharedChartPoints || (uint)points.Count != count.Value)
                    throw new NotSupportedException("Replacing a chart workbook requires bounded cached values for custom error bars.");
                var indexes = new HashSet<uint>();
                foreach (C.NumericPoint point in points) {
                    uint? index = point.Index?.Value;
                    if (!index.HasValue || index.Value >= count!.Value || !indexes.Add(index.Value))
                        throw new NotSupportedException("Custom error-bar caches must have bounded, unique point indexes.");
                    if (!double.TryParse(point.NumericValue?.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) ||
                        double.IsNaN(value) || double.IsInfinity(value))
                        throw new NotSupportedException("Custom error-bar caches must provide a finite numeric value for every declared point.");
                }
                var literal = new C.NumberLiteral(cache.ChildElements.Select(child => child.CloneNode(true)));
                if (literal.FormatCode == null) literal.AddChild(new C.FormatCode { Text = "General" }, true);
                replacements.Add(new KeyValuePair<OpenXmlElement, OpenXmlElement>(reference, literal));
            }
            foreach (OpenXmlElement formula in part.ChartSpace.Descendants().Where(element => element.LocalName == "f")) {
                bool rewrittenSource = formula.Ancestors().Any(element => element is C.SeriesText || element is C.CategoryAxisData ||
                    element is C.Values || element is C.XValues || element is C.YValues || element is C.BubbleSize);
                bool preservedText = formula.Parent is C.StringReference && formula.Parent.Parent is C.ChartText;
                bool preservedError = formula.Parent is C.NumberReference && (formula.Parent.Parent is C.Plus || formula.Parent.Parent is C.Minus);
                if (!rewrittenSource && !preservedText && !preservedError)
                    throw new NotSupportedException("The chart contains a workbook-linked extension that cannot be preserved while replacing the embedded worksheet.");
            }
            return replacements;
        }
    }
}
