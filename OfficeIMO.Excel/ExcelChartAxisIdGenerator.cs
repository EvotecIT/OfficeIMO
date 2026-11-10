using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Excel {
    internal static class ExcelChartAxisIdGenerator {
        private const long BaseAxisId = 48650112L;
        private static long _axisIdSeed = BaseAxisId;
        private static readonly ConditionalWeakTable<OpenXmlPackage, object> ScannedPackages = new();

        /// <summary>
        /// Raises the seed above the axis ids already in a package before the first chart in it
        /// takes a new id. Scanning on first use spares opening every worksheet's relationships
        /// when a workbook is loaded and no chart is ever built or updated.
        /// </summary>
        internal static void EnsureInitialized(OpenXmlPackage? package) {
            if (package is not SpreadsheetDocument document) return;
            lock (ScannedPackages) {
                if (ScannedPackages.TryGetValue(document, out _)) return;
                Initialize(document);
                ScannedPackages.Add(document, ScannedPackages);
            }
        }

        private static void Initialize(SpreadsheetDocument document) {
            if (document.WorkbookPart == null) {
                return;
            }

            long currentSeed = Interlocked.Read(ref _axisIdSeed);
            long max = Math.Max(currentSeed, BaseAxisId);

            foreach (var worksheetPart in document.WorkbookPart.WorksheetParts) {
                var drawings = worksheetPart.DrawingsPart;
                if (drawings == null) {
                    continue;
                }

                foreach (ChartPart chartPart in drawings.ChartParts) {
                    var chart = chartPart.ChartSpace?.GetFirstChild<Chart>();
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
