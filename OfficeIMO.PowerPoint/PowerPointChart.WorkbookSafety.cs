using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private const string ChartWorkbookContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        /// <summary>
        /// Whether a package embedded in a chart is the chart's own data workbook, referenced from the chart's
        /// <c>c:externalData</c>, and simple enough for OfficeIMO to rewrite without changing its meaning.
        /// </summary>
        internal static bool IsSafeChartWorkbookPart(IOfficeChartPart chartPart, IOfficeChartEmbeddedPackage package) {
            ArgumentNullException.ThrowIfNull(chartPart);
            ArgumentNullException.ThrowIfNull(package);
            if (!string.Equals(package.ContentType, ChartWorkbookContentType, StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            string relationshipId = package.RelationshipId;
            OfficeXmlNode chartSpace = chartPart.ChartSpace;
            bool isReferenced = !chartSpace.IsNone
                && chartSpace.Descendants(OfficeXmlNamespaces.Chart, "externalData")
                    .Any(externalData => string.Equals(
                        externalData.Attribute(OfficeXmlNamespaces.Relationships, "id"), relationshipId, StringComparison.Ordinal));
            if (!isReferenced) {
                return false;
            }

            try {
                using Stream stream = package.OpenRead();
                byte[] workbookBytes = OfficeOpenXmlChartWorkbookSecurity.ReadAndValidate(stream);
                return PowerPointChartWorkbookEditor.IsSafelyEditable(workbookBytes);
            } catch (Exception exception) when (IsUnreadableChartWorkbook(exception)) {
                return false;
            }
        }

        // The failures that mean the workbook can't be read as a package, as the SDK-based check treats them.
        private static bool IsUnreadableChartWorkbook(Exception exception) {
            if (exception is InvalidDataException or IOException) return true;
            for (Type? type = exception.GetType(); type != null; type = type.BaseType) {
                if (type.FullName == "System.IO.FileFormatException" || type.Name == "OpenXmlPackageException") return true;
            }
            return false;
        }
    }
}
