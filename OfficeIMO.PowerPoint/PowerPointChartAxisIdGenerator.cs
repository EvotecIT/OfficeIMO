using System;
using System.Threading;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.PowerPoint {
    internal static class PowerPointChartAxisIdGenerator {
        // PowerPoint seeds chart axis identifiers starting at 48650112 when it
        // generates charts. We mirror that behaviour so documents created with
        // OfficeIMO match what the desktop client produces.
        private const long BaseAxisId = 48650112L;
        private static long _axisIdSeed = BaseAxisId;

        internal static void Initialize(PresentationPart presentationPart) {
            if (presentationPart == null) {
                return;
            }

            long currentSeed = Interlocked.Read(ref _axisIdSeed);
            long max = Math.Max(currentSeed, BaseAxisId);

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

            Interlocked.Exchange(ref _axisIdSeed, max);
        }

        internal static uint GetNextId() {
            long next = Interlocked.Increment(ref _axisIdSeed);
            if (next < 0 || next > uint.MaxValue) {
                throw new InvalidOperationException("Chart axis id seed exceeded the range of valid UInt32 values.");
            }

            return (uint)next;
        }
    }
}
