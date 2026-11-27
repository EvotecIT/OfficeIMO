using System.Runtime.CompilerServices;
using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel {
    internal static partial class ExcelChartAxisIdGenerator {
        private const long BaseAxisId = 48650112L;
        private static long _axisIdSeed = BaseAxisId;
        // Keyed by the package's workbook part, which lives as long as the package.
        private static readonly ConditionalWeakTable<object, object> ScannedPackages = new();

        /// <summary>
        /// Raises the seed above the axis ids already in a package before the first chart in it
        /// takes a new id. Scanning on first use spares opening every worksheet's relationships
        /// when a workbook is loaded and no chart is ever built or updated.
        /// </summary>
        /// <param name="part">Any part of the spreadsheet package.</param>
        internal static void EnsureInitialized(ExcelPartRef part) {
            ExcelPartRef workbook = part.PackageWorkbook;
            if (workbook.IsNone || workbook.SdkPart is not { } key) return;
            lock (ScannedPackages) {
                if (ScannedPackages.TryGetValue(key, out _)) return;
                Initialize(workbook);
                ScannedPackages.Add(key, ScannedPackages);
            }
        }

        /// <summary>Whether any worksheet of the workbook has a chart in its drawing.</summary>
        internal static bool HasWorksheetCharts(ExcelPartRef workbook) {
            foreach (ExcelPartRef worksheetPart in workbook.PartsOfKind(ExcelPartKind.Worksheet)) {
                if (worksheetPart.Drawings.FirstPartOfKind(ExcelPartKind.Chart).Exists) return true;
            }
            return false;
        }

        private static void Initialize(ExcelPartRef workbook) {
            long currentSeed = Interlocked.Read(ref _axisIdSeed);
            long max = Math.Max(currentSeed, BaseAxisId);

            foreach (ExcelPartRef worksheetPart in workbook.PartsOfKind(ExcelPartKind.Worksheet)) {
                ExcelPartRef drawings = worksheetPart.Drawings;
                if (drawings.IsNone) {
                    continue;
                }

                foreach (ExcelPartRef chartPart in drawings.PartsOfKind(ExcelPartKind.Chart)) {
                    OfficeXmlNode chart = chartPart.Root.Child(C, "chart");
                    if (chart.IsNone) {
                        continue;
                    }

                    foreach (OfficeXmlNode axisId in chart.DescendantsOutsideExtensions(C, "axId")) {
                        if (axisId.UInt32Val() is uint value && value > max) {
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
