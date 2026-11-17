using OfficeIMO.OpenXml.Internal;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Features;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.PowerPoint {
    internal static class PowerPointChartAxisIdGenerator {
        // PowerPoint seeds chart axis identifiers starting at 48650112 when it
        // generates charts. We mirror that behaviour so documents created with
        // OfficeIMO match what the desktop client produces.
        private const long BaseAxisId = 48650112L;

        /// <summary>
        /// Raises the shared seed above the axis identifiers already in this package before it
        /// allocates new ones. Loading does not pay for this: it parses every chart part, and
        /// reading or converting a presentation never allocates an axis identifier.
        /// </summary>
        internal static void EnsureSeeded(OpenXmlPackage package) {
            if (package is not PresentationDocument document) return;
            PowerPointDocument? owner = document.Features.Get<PowerPointDocument>();
            if (owner != null) {
                if (owner.ChartAxisSeedApplied) return;
                owner.ChartAxisSeedApplied = true;
            }
            Initialize(document.PresentationPart!);
        }

        private static void Initialize(PresentationPart presentationPart) {
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
                        // Imported IDs near UInt32.MaxValue cannot become the shared
                        // process seed: one generated chart would exhaust the allocator.
                        if (axisId.Val?.Value is uint value && value <= int.MaxValue && value > max) {
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
