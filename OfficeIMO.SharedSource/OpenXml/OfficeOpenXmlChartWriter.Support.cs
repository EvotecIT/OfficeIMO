using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartWriter {
        private const string ChartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        private const string DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private const string RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static long _axisIdSeed = 48650112;

        internal static byte[] BuildWorkbook(OfficeChartData data, OfficeChartKind kind) {
            ValidateSharedChartData(data, kind);
            return kind switch {
                OfficeChartKind.Bubble => BuildBubbleChartWorkbook(data),
                OfficeChartKind.Scatter => BuildScatterWorkbook(NormalizeScatterData(data)),
                _ => BuildCategoryWorkbook(data)
            };
        }

        internal static IReadOnlyList<double> ParseNumericCategories(IReadOnlyList<string> categories) =>
            ParseScatterCategories(categories);

        // Native scatter caches and worksheet columns require explicit numeric X values.
        // Appearance is applied separately from the caller's original shared series.
        private static OfficeChartData NormalizeScatterData(OfficeChartData data) {
            IReadOnlyList<double>? sharedX = data.Series.Any(series => series.XValues == null)
                ? ParseScatterCategories(data.Categories) : null;
            return new OfficeChartData(data.Categories, data.Series.Select(series =>
                new OfficeChartSeries(series.Name, series.Values, series.XValues ?? sharedX!)));
        }

        internal static uint GetNextAxisId() {
            long next = Interlocked.Increment(ref _axisIdSeed);
            if (next < 0 || next > uint.MaxValue)
                throw new InvalidOperationException("Chart axis identifiers exceeded the UInt32 range.");
            return (uint)next;
        }

        internal static void AdvanceAxisSeed(long minimum) {
            while (true) {
                long current = Interlocked.Read(ref _axisIdSeed);
                if (minimum <= current) return;
                if (Interlocked.CompareExchange(ref _axisIdSeed, minimum, current) == current) return;
            }
        }

        private static long FromPoints(double points) => (long)Math.Round(points * 12700);
    }
}
