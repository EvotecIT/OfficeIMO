using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Enables or disables workbook refresh-on-open metadata for supported workbook data features.
        /// OfficeIMO writes the metadata request; Excel-compatible applications perform the actual refresh on open.
        /// </summary>
        /// <param name="enabled">Whether refresh-on-open should be enabled.</param>
        /// <param name="pivotTables">Update pivot cache definitions.</param>
        /// <param name="connections">Update workbook connection metadata parts.</param>
        /// <param name="savePivotSourceData">Optional pivot cache source-data setting to apply together with the refresh request.</param>
        /// <returns>Counts of metadata entries updated.</returns>
        public ExcelRefreshOnOpenResult SetRefreshOnOpen(
            bool enabled = true,
            bool pivotTables = true,
            bool connections = true,
            bool? savePivotSourceData = null) {
            int pivotCacheCount = 0;
            int connectionCount = 0;

            if (pivotTables) {
                foreach (var cachePart in WorkbookPartRoot.GetPartsOfType<PivotTableCacheDefinitionPart>()) {
                    var cacheDefinition = cachePart.PivotCacheDefinition;
                    if (cacheDefinition == null) {
                        continue;
                    }

                    cacheDefinition.RefreshOnLoad = enabled;
                    if (savePivotSourceData.HasValue) {
                        cacheDefinition.SaveData = savePivotSourceData.Value;
                    }

                    cacheDefinition.Save();
                    pivotCacheCount++;
                }
            }

            if (connections) {
                foreach (OpenXmlPart part in EnumerateWorkbookConnectionParts()) {
                    connectionCount += SetConnectionPartRefreshOnOpen(part, enabled);
                }
            }

            if (pivotCacheCount > 0 || connectionCount > 0) {
                MarkPackageDirty();
            }

            return new ExcelRefreshOnOpenResult(enabled, pivotCacheCount, connectionCount);
        }

        private static int SetConnectionPartRefreshOnOpen(OpenXmlPart part, bool enabled) {
            XDocument document;
            document = XDocument.Parse(ReadOpenXmlPartText(part), LoadOptions.PreserveWhitespace);

            int count = 0;
            foreach (XElement connection in document.Descendants().Where(element => element.Name.LocalName == "connection")) {
                connection.SetAttributeValue("refreshOnLoad", enabled ? "1" : "0");
                count++;
            }

            if (count == 0) {
                return 0;
            }

            WriteOpenXmlPartText(part, document.ToString(SaveOptions.DisableFormatting));

            return count;
        }
    }
}
