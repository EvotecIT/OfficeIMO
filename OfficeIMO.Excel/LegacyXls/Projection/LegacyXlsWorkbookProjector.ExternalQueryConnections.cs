using OfficeIMO.Excel.LegacyXls.Model;
using System.Globalization;
using System.Xml.Linq;

namespace OfficeIMO.Excel.LegacyXls.Projection {
    internal static partial class LegacyXlsWorkbookProjector {
        private static void ProjectExternalQueryConnections(LegacyXlsWorkbook workbook, ExcelDocument document) {
            if (workbook.ExternalQueryConnections.Count == 0) {
                return;
            }

            document.AddWorkbookConnectionMetadata(BuildExternalQueryConnectionMetadataXml(workbook.ExternalQueryConnections));
        }

        private static string BuildExternalQueryConnectionMetadataXml(IReadOnlyList<LegacyXlsExternalQueryConnection> connections) {
            XNamespace main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var root = new XElement(main + "connections");
            uint connectionId = 1U;
            foreach (LegacyXlsExternalQueryConnection connection in connections) {
                root.Add(BuildExternalQueryConnectionElement(main, connection, connectionId));
                connectionId++;
            }

            root.SetAttributeValue("count", connections.Count.ToString(CultureInfo.InvariantCulture));
            return new XDocument(root).ToString(SaveOptions.DisableFormatting);
        }

        private static XElement BuildExternalQueryConnectionElement(XNamespace main, LegacyXlsExternalQueryConnection connection, uint connectionId) {
            string name = BuildExternalQueryConnectionName(connection, connectionId);
            var element = new XElement(main + "connection",
                new XAttribute("id", connectionId.ToString(CultureInfo.InvariantCulture)),
                new XAttribute("name", name),
                new XAttribute("type", connection.DataSourceType.ToString(CultureInfo.InvariantCulture)),
                new XAttribute("refreshedVersion", connection.RefreshedVersion.ToString(CultureInfo.InvariantCulture)),
                new XAttribute("minRefreshableVersion", connection.RefreshableMinimumVersion.ToString(CultureInfo.InvariantCulture)),
                new XAttribute("description", BuildExternalQueryConnectionDescription(connection)));

            if (connection.MaintainConnection) {
                element.SetAttributeValue("keepAlive", "1");
            }

            if (connection.RefreshIntervalMinutes > 0) {
                element.SetAttributeValue("interval", connection.RefreshIntervalMinutes.ToString(CultureInfo.InvariantCulture));
            }

            return element;
        }

        private static string BuildExternalQueryConnectionName(LegacyXlsExternalQueryConnection connection, uint connectionId) {
            string baseName = string.IsNullOrWhiteSpace(connection.SheetName)
                ? "LegacyXlsQuery"
                : connection.SheetName!.Trim() + "Query";

            return baseName + connectionId.ToString(CultureInfo.InvariantCulture);
        }

        private static string BuildExternalQueryConnectionDescription(LegacyXlsExternalQueryConnection connection) {
            var parts = new List<string> {
                "Legacy XLS DBQueryExt metadata",
                "Source=" + connection.SourceTypeName,
                "EditVersion=" + connection.EditVersion.ToString(CultureInfo.InvariantCulture),
                "RefreshedVersion=" + connection.RefreshedVersion.ToString(CultureInfo.InvariantCulture),
                "RefreshableMinimumVersion=" + connection.RefreshableMinimumVersion.ToString(CultureInfo.InvariantCulture)
            };

            if (connection.RefreshIntervalMinutes > 0) {
                parts.Add("RefreshIntervalMinutes=" + connection.RefreshIntervalMinutes.ToString(CultureInfo.InvariantCulture));
            }

            if (connection.ConnectionFlagNames.Count > 0) {
                parts.Add("Flags=" + string.Join(",", connection.ConnectionFlagNames));
            }

            if (connection.QueryOptionNames.Count > 0) {
                parts.Add("QueryOptions=" + string.Join(",", connection.QueryOptionNames));
            }

            return string.Join("; ", parts);
        }
    }
}
