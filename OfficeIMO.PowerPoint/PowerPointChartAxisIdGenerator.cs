using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    internal static class PowerPointChartAxisIdGenerator {
        // PowerPoint seeds chart axis identifiers starting at 48650112 when it
        // generates charts. We mirror that behaviour so documents created with
        // OfficeIMO match what the desktop client produces.
        private const long BaseAxisId = 48650112L;

        private const string ChartRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";

        /// <summary>
        /// Raises the shared seed above the axis identifiers already in this presentation before it
        /// allocates new ones. Loading does not pay for this: it parses every chart part, and
        /// reading or converting a presentation never allocates an axis identifier.
        /// </summary>
        internal static void EnsureSeeded(PowerPointDocument? document) {
            if (document == null || document.ChartAxisSeedApplied) return;
            document.ChartAxisSeedApplied = true;
            Initialize(document.PresentationPartRef);
        }

        private static void Initialize(PowerPointPartRef presentationPart) {
            if (presentationPart.IsNone) {
                return;
            }

            long max = BaseAxisId;

            foreach (PowerPointPartRef slidePart in presentationPart.RelatedParts(PowerPointPartRef.SlideRelationshipType)) {
                foreach (PowerPointPartRef chartPart in slidePart.RelatedParts(ChartRelationshipType)) {
                    OfficeXmlNode chart = chartPart.Root.Child(OfficeXmlNamespaces.Chart, "chart");
                    if (chart.IsNone) {
                        continue;
                    }

                    foreach (OfficeXmlNode axisId in chart.Descendants(OfficeXmlNamespaces.Chart, "axId")) {
                        // Imported IDs near UInt32.MaxValue cannot become the shared
                        // process seed: one generated chart would exhaust the allocator.
                        if (axisId.UInt32Attribute("val") is uint value && value <= int.MaxValue && value > max) {
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
