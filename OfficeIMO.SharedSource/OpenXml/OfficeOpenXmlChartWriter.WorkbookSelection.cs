using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal.OpenXml;
using static OfficeIMO.OpenXml.Internal.OfficeChartXml;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartWriter {
        internal const string SharedWorkbookContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        /// <summary>Selects the supported workbook before any native chart or package mutation.</summary>
        internal static IOfficeChartEmbeddedPackage? GetSharedEmbeddedWorkbook(IOfficeChartPart part) {
            IReadOnlyList<IOfficeChartEmbeddedPackage> packages = part.EmbeddedPackages(2);
            if (packages.Count > 1)
                throw new NotSupportedException("Charts with multiple embedded packages require an explicit package selection.");
            IOfficeChartEmbeddedPackage? package = packages.Count == 0 ? null : packages[0];
            OfficeXmlNode chartSpace = part.ChartSpace;
            if (!chartSpace.IsNone) {
                OfficeXmlNode reference = chartSpace.Child(NsC, "externalData");
                if (!reference.IsNone && !reference.NextSiblingNamed(NsC, "externalData").IsNone)
                    throw new NotSupportedException("Charts with multiple workbook references require an explicit workbook selection.");
                if (!reference.IsNone) {
                    string? id = reference.Attribute(NsR, "id");
                    if (string.IsNullOrWhiteSpace(id) || package == null || package.RelationshipId != id)
                        throw new NotSupportedException("Chart data updates require a referenced embedded XLSX workbook; external or unresolved workbook links are preserved without mutation.");
                }
            }
            if (package != null && !string.Equals(package.ContentType, SharedWorkbookContentType, StringComparison.OrdinalIgnoreCase))
                throw new NotSupportedException("Shared chart data updates require an XLSX embedded workbook; the existing package format is preserved.");
            return package;
        }

        // SDK callers
        internal static EmbeddedPackagePart? GetSharedEmbeddedWorkbook(ChartPart part) =>
            GetSharedEmbeddedWorkbook(new OfficeSdkChartPart(part)) is { } package ? (EmbeddedPackagePart)part.GetPartById(package.RelationshipId) : null;
    }
}
