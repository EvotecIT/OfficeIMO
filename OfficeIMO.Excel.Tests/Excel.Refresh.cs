using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ExcelRefreshOnOpen_UpdatesPivotCachesAndConnections() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelRefreshOnOpen.PivotConnections.xlsx");
            const string connectionXml = "<connections xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\"><connection id=\"1\" name=\"SalesConnection\" type=\"5\" refreshedVersion=\"7\"/></connections>";

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 2, "Sales");
                sheet.CellValue(2, 1, "East");
                sheet.CellValue(2, 2, 10);
                sheet.CellValue(3, 1, "West");
                sheet.CellValue(3, 2, 20);
                sheet.AddPivotTable(
                    sourceRange: "A1:B3",
                    destinationCell: "E2",
                    name: "SalesPivot",
                    rowFields: new[] { "Region" },
                    dataFields: new[] { new ExcelPivotDataField("Sales") },
                    options: new ExcelPivotTableOptions {
                        RefreshOnOpen = false,
                        SaveSourceData = true
                    });
                document.AddWorkbookConnectionMetadata(connectionXml);

                ExcelRefreshOnOpenResult result = document.SetRefreshOnOpen(savePivotSourceData: false);
                Assert.True(result.Enabled);
                Assert.Equal(1, result.PivotCacheCount);
                Assert.Equal(1, result.ConnectionCount);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var cacheDefinition = spreadsheet.WorkbookPart!.PivotTableCacheDefinitionParts.Single().PivotCacheDefinition!;
                Assert.True(cacheDefinition.RefreshOnLoad!.Value);
                Assert.False(cacheDefinition.SaveData!.Value);

                string connectionText = ReadSinglePackagePartText(spreadsheet.WorkbookPart!, "connections");
                XDocument connections = XDocument.Parse(connectionText);
                XElement connection = connections.Descendants().Single(element => element.Name.LocalName == "connection");
                Assert.Equal("1", connection.Attribute("refreshOnLoad")?.Value);
            }
        }

        [Fact]
        public void Test_ExcelRefreshOnOpen_UpdatesTypedConnectionsAndAllocatesNextConnectionId() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelRefreshOnOpen.TypedConnections.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorkSheet("Data");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                ConnectionsPart connectionsPart = spreadsheet.WorkbookPart!.AddNewPart<ConnectionsPart>();
                connectionsPart.Connections = new Connections(
                    new Connection { Id = 1U, Name = "ExistingOne", Type = 5, RefreshedVersion = 7 },
                    new Connection { Id = 2U, Name = "ExistingTwo", Type = 5, RefreshedVersion = 7 });
                connectionsPart.Connections.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                ExcelRefreshOnOpenResult refresh = document.SetRefreshOnOpen(pivotTables: false, connections: true);
                Assert.Equal(2, refresh.ConnectionCount);

                ExcelPowerQueryMetadataResult metadata = document.AddPowerQueryMetadata(new ExcelPowerQueryMetadataOptions {
                    Name = "AddedQuery",
                    WorksheetName = "Data",
                    CommandText = "let Source = Excel.CurrentWorkbook() in Source",
                    RefreshOnOpen = true
                });
                Assert.Equal(3U, metadata.ConnectionId);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                ConnectionsPart connectionsPart = spreadsheet.WorkbookPart!.GetPartsOfType<ConnectionsPart>()
                    .First(part => part.Connections?.Elements<Connection>().Any(connection => connection.Name?.Value == "ExistingOne") == true);
                Assert.All(connectionsPart.Connections!.Elements<Connection>(), connection => Assert.True(connection.RefreshOnLoad!.Value));
                Assert.DoesNotContain(spreadsheet.WorkbookPart!.Parts.Select(pair => pair.OpenXmlPart), part => part is ExtendedPart && part.ContentType.IndexOf("connections", StringComparison.OrdinalIgnoreCase) >= 0);

                XElement addedConnection = EnumerateWorkbookConnectionXml(spreadsheet.WorkbookPart!)
                    .Select(XDocument.Parse)
                    .SelectMany(document => document.Descendants().Where(element => element.Name.LocalName == "connection"))
                    .Single(element => element.Attribute("name")?.Value == "AddedQuery");
                Assert.Equal("3", addedConnection.Attribute("id")?.Value);
            }
        }
    }
}
