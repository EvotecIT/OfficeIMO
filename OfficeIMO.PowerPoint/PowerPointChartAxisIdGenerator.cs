using System;
using OfficeIMO.OpenXml.Internal;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.PowerPoint {
    internal static class PowerPointChartAxisIdGenerator {
        // PowerPoint seeds chart axis identifiers starting at 48650112 when it
        // generates charts. We mirror that behaviour so documents created with
        // OfficeIMO match what the desktop client produces.
        private const long BaseAxisId = 48650112L;

        internal static void Initialize(PresentationPart presentationPart) {
            if (presentationPart == null) {
                return;
            }

            long max = BaseAxisId;

            foreach (SlidePart slidePart in presentationPart.SlideParts) {
                foreach (ChartPart chartPart in slidePart.ChartParts) {
                    ChartSpace? chartSpace = chartPart.ChartSpace;
                    Chart? chart = chartSpace?.GetFirstChild<Chart>();
                    if (chart == null) {
                        continue;
                    }

                    foreach (AxisId axisId in chart.Descendants<AxisId>()) {
                        if (axisId.Val?.Value is uint value && value > max) {
                            max = value;
                        }
                    }
                }
            }

            AdvanceSeedToAtLeast(max);
        }

        internal static uint GetNextId() => OfficeOpenXmlChartWriter.GetNextAxisId();

        internal static void AdvanceSeedToAtLeast(long minimumSeed) =>
            OfficeOpenXmlChartWriter.AdvanceAxisSeed(minimumSeed);
    }
}