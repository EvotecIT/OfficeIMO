using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartWriter {
        internal const string SharedWorkbookContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        /// <summary>Selects the supported workbook before any native chart or package mutation.</summary>
        internal static EmbeddedPackagePart? GetSharedEmbeddedWorkbook(ChartPart part) {
            var packages = part.GetPartsOfType<EmbeddedPackagePart>().Take(2).ToList();
            if (packages.Count > 1)
                throw new NotSupportedException("Charts with multiple embedded packages require an explicit package selection.");
            EmbeddedPackagePart? package = packages.FirstOrDefault();
            var references = part.ChartSpace?.Elements<C.ExternalData>().Take(2).ToList();
            if (references?.Count > 1)
                throw new NotSupportedException("Charts with multiple workbook references require an explicit workbook selection.");
            if (references?.Count == 1) {
                string? id = references[0].Id?.Value;
                if (string.IsNullOrWhiteSpace(id) || package == null ||
                    !part.Parts.Any(pair => pair.RelationshipId == id && ReferenceEquals(pair.OpenXmlPart, package)))
                    throw new NotSupportedException("Chart data updates require a referenced embedded XLSX workbook; external or unresolved workbook links are preserved without mutation.");
            }
            if (package != null && !string.Equals(package.ContentType, SharedWorkbookContentType, StringComparison.OrdinalIgnoreCase))
                throw new NotSupportedException("Shared chart data updates require an XLSX embedded workbook; the existing package format is preserved.");
            return package;
        }
    }
}
