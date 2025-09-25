using System;
using System.Threading;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.PowerPoint {
    internal static class PowerPointChartAxisIdGenerator {
        private const int BaseAxisId = 48650112;
        private static int _axisIdSeed = BaseAxisId;

        internal static void Initialize(PresentationPart presentationPart) {
            if (presentationPart == null) {
                return;
            }

            uint max = (uint)Math.Max(_axisIdSeed, BaseAxisId);

            foreach (SlidePart slidePart in presentationPart.SlideParts) {
                foreach (ChartPart chartPart in slidePart.ChartParts) {
                    ChartSpace? chartSpace = chartPart.ChartSpace;
                    Chart? chart = chartSpace?.GetFirstChild<Chart>();
                    if (chart == null) {
                        continue;
                    }

                    foreach (AxisId axisId in chart.Descendants<AxisId>()) {
                        uint? value = axisId.Val?.Value;
                        if (value.HasValue && value.Value > max) {
                            max = value.Value;
                        }
                    }
                }
            }

            Interlocked.Exchange(ref _axisIdSeed, (int)max);
        }

        internal static uint GetNextId() {
            int next = Interlocked.Increment(ref _axisIdSeed);
            return (uint)next;
        }
    }
}
