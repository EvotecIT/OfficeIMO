using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartWriter {
        internal const string SharedWorkbookContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        /// <summary>Selects the supported workbook before any native chart or package mutation.</summary>
        internal static EmbeddedPackagePart? GetSharedEmbeddedWorkbook(ChartPart part) {
            var packages = part.GetPartsOfType<EmbeddedPackagePart>().Take(2).ToList();
            if (packages.Count > 1)
                throw new NotSupportedException("Charts with multiple embedded packages require an explicit package selection.");
            EmbeddedPackagePart? package = packages.FirstOrDefault();
            if (package != null && !string.Equals(package.ContentType, SharedWorkbookContentType, StringComparison.OrdinalIgnoreCase))
                throw new NotSupportedException("Shared chart data updates require an XLSX embedded workbook; the existing package format is preserved.");
            return package;
        }
    }
}
