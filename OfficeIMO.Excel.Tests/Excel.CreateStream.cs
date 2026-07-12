using System.Data;
using System.IO;
using System.IO.Compression;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Create_ToMemoryStream_ExplicitPersistenceDoesNotWriteUntilSave() {
            using var memory = new MemoryStream();
            using (var document = ExcelDocument.Create(memory)) {
                document.AddWorkSheet("Explicit").CellValue(1, 1, "Pending");
            }

            Assert.Equal(0, memory.Length);

            using (var document = ExcelDocument.Create(memory)) {
                document.AddWorkSheet("Explicit").CellValue(1, 1, "Saved");
                document.Save();
            }

            using var reader = ExcelDocumentReader.Open(memory.ToArray());
            Assert.Equal("Saved", reader.GetSheet("Explicit").ReadRange("A1:A1")[0, 0]);
        }

        [Fact]
        public void Create_ToPath_ExplicitPersistenceDoesNotCreateFileUntilSave() {
            string filePath = Path.Combine(_directoryWithFiles, "CreateExplicit.xlsx");
            File.Delete(filePath);

            using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorkSheet("Explicit");
                Assert.False(File.Exists(filePath));
                document.Save();
            }

            Assert.True(File.Exists(filePath));
            File.Delete(filePath);
        }

        [Fact]
        public void Create_ToMemoryStream_SaveOnDisposeWritesPackage() {
            using var memory = new MemoryStream();
            using (var document = ExcelDocument.Create(memory, new ExcelCreateOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose })) {
                var sheet = document.AddWorkSheet("StreamData");
                sheet.CellValue(1, 1, "Hello Stream");
            }

            Assert.True(memory.Length > 0);
            memory.Position = 0;

            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            Assert.NotNull(spreadsheet.WorkbookPart);
            var sheets = spreadsheet.WorkbookPart!.Workbook!.Sheets!.OfType<Sheet>().ToList();
            var sheetInfo = Assert.Single(sheets);
            Assert.Equal("StreamData", sheetInfo.Name?.Value);
        }

        [Fact]
        public void Save_LoadedUnchangedStream_WritesReadablePackage() {
            byte[] originalBytes = CreateStreamSaveWorkbookBytes();

            using var input = new MemoryStream(originalBytes, writable: false);
            using var output = new MemoryStream();
            using (var document = ExcelDocument.Load(input)) {
                document.Save(output);
            }

            using var reader = ExcelDocumentReader.Open(output.ToArray());
            object?[,] values = reader.GetSheet("StreamData").ReadRange("A1:A1");
            Assert.Equal("Hello Stream", values[0, 0]);
        }

        [Fact]
        public void Save_LoadedStreamAfterCellEdit_WritesEditedPackage() {
            byte[] originalBytes = CreateStreamSaveWorkbookBytes();

            using var input = new MemoryStream(originalBytes, writable: false);
            using var output = new MemoryStream();
            using (var document = ExcelDocument.Load(input)) {
                var sheet = document.GetSheet("StreamData");
                sheet.CellValue(2, 1, "Edited");
                document.Save(output);
            }

            using var reader = ExcelDocumentReader.Open(output.ToArray());
            object?[,] values = reader.GetSheet("StreamData").ReadRange("A1:A2");
            Assert.Equal("Hello Stream", values[0, 0]);
            Assert.Equal("Edited", values[1, 0]);
        }

        [Fact]
        public void Save_LoadedStreamAfterThemeEdit_WritesEditedTheme() {
            byte[] originalBytes = CreateStreamSaveWorkbookBytes();

            using var input = new MemoryStream(originalBytes, writable: false);
            using var output = new MemoryStream();
            using (var document = ExcelDocument.Load(input)) {
                document.SetWorkbookThemeName("Stream Theme");
                document.Save(output);
            }

            using var reloaded = ExcelDocument.Load(new MemoryStream(output.ToArray()), new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            Assert.Equal("Stream Theme", reloaded.GetWorkbookTheme().Name);
        }

        [Fact]
        public void ThemeEdit_LoadedPathWithExplicitPersistence_DoesNotWriteUntilSaved() {
            string filePath = Path.Combine(_directoryWithFiles, "ThemeAutoSaveFalse.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorkSheet("Data");
                document.SetWorkbookThemeName("Original Theme");
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                document.SetWorkbookThemeName("Unsaved Theme");
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Equal("Original Theme", document.GetWorkbookTheme().Name);
            }

            using (var document = ExcelDocument.Load(filePath)) {
                document.SetWorkbookThemeName("Saved Theme");
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Equal("Saved Theme", document.GetWorkbookTheme().Name);
            }
        }

        [Fact]
        public void Create_ToMemoryStream_WithTableAutoFitAndDate_WritesReadablePackage() {
            var rows = new[] {
                new StreamSalesRow(1, "North", new DateTime(2024, 1, 2), 123.45d),
                new StreamSalesRow(2, "South", new DateTime(2024, 1, 3), 234.56d)
            };

            using var memory = new MemoryStream();
            using (var document = ExcelDocument.Create(memory, new ExcelCreateOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose })) {
                document.Execution.SaveWorksheetAfterAutoFit = false;
                var sheet = document.AddWorkSheet("Data");
                sheet.InsertObjects(rows,
                    ("Id", item => item.Id),
                    ("Region", item => item.Region),
                    ("CreatedOn", item => item.CreatedOn),
                    ("Amount", item => item.Amount));
                sheet.AddTable("A1:D3", hasHeader: true, name: "SalesData", style: OfficeIMO.Excel.TableStyle.TableStyleMedium2, includeAutoFilter: true);
                sheet.AutoFitColumns();
            }

            Assert.True(memory.Length > 0);
            memory.Position = 0;

            using (var spreadsheet = SpreadsheetDocument.Open(memory, false)) {
                var errors = new OpenXmlValidator().Validate(spreadsheet).ToList();
                Assert.Empty(errors);

                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                TableDefinitionPart tablePart = Assert.Single(wsPart.TableDefinitionParts);
                Assert.Equal("SalesData", tablePart.Table.Name);
                Assert.Equal("A1:D3", tablePart.Table.Reference!.Value);
                Assert.NotNull(tablePart.Table.GetFirstChild<AutoFilter>());

                var dateCell = wsPart.Worksheet.Descendants<Cell>().First(cell => cell.CellReference?.Value == "C2");
                Assert.True(dateCell.DataType == null || dateCell.DataType.Value == CellValues.Number);
                Assert.NotNull(dateCell.StyleIndex);
                Assert.NotNull(wsPart.Worksheet.GetFirstChild<Columns>());
            }

            using var reader = ExcelDocumentReader.Open(memory.ToArray());
            object?[,] values = reader.GetSheet("Data").ReadRange("A1:D3");
            Assert.Equal("Region", values[0, 1]);
            Assert.Equal("North", values[1, 1]);
            Assert.Equal(123.45d, Assert.IsType<double>(values[1, 3]));
        }

        [Fact]
        public void Create_ToMemoryStream_DataTableExportWithTableVisuals_WritesReadablePackage() {
            DataTable orders = new DataTable("Orders");
            orders.Columns.Add("Id", typeof(int));
            orders.Columns.Add("Customer", typeof(string));
            orders.Columns.Add("CreatedOn", typeof(DateTime));
            orders.Columns.Add("Status", typeof(string));
            orders.Columns.Add("Amount", typeof(double));
            orders.Rows.Add(1, "Contoso", new DateTime(2024, 2, 1), "Open", 120.5d);
            orders.Rows.Add(2, "Fabrikam", new DateTime(2024, 2, 2), "Closed", 250.75d);
            orders.Rows.Add(3, "Northwind", new DateTime(2024, 2, 3), "Open", 87.25d);

            using var memory = new MemoryStream();
            using (var document = ExcelDocument.Create(memory, new ExcelCreateOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose })) {
                document.Execution.SaveWorksheetAfterAutoFit = false;
                var sheet = document.AddWorkSheet("Orders");
                string range = sheet.InsertDataTableAsTable(
                    orders,
                    tableName: "OrdersExport",
                    style: OfficeIMO.Excel.TableStyle.TableStyleMedium9,
                    includeAutoFilter: true);
                sheet.AddAutoFilter(range, new Dictionary<uint, IEnumerable<string>> {
                    { 3U, new[] { "Open" } }
                });
                sheet.ValidationList("D2:D4", new[] { "Open", "Closed", "Hold" });
                sheet.Freeze(topRows: 1, leftCols: 0);
                sheet.AutoFitColumns();
            }

            Assert.True(memory.Length > 0);
            memory.Position = 0;

            using (var spreadsheet = SpreadsheetDocument.Open(memory, false)) {
                var errors = new OpenXmlValidator().Validate(spreadsheet).ToList();
                Assert.Empty(errors);

                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                Assert.NotNull(wsPart.Worksheet.GetFirstChild<SheetViews>());
                Assert.NotNull(wsPart.Worksheet.GetFirstChild<DataValidations>());
                Assert.NotNull(wsPart.Worksheet.GetFirstChild<Columns>());

                TableDefinitionPart tablePart = Assert.Single(wsPart.TableDefinitionParts);
                Assert.Equal("OrdersExport", tablePart.Table.Name);
                Assert.Equal("TableStyleMedium9", tablePart.Table.TableStyleInfo?.Name?.Value);
                AutoFilter tableFilter = Assert.Single(tablePart.Table.Elements<AutoFilter>());
                FilterColumn filterColumn = Assert.Single(tableFilter.Elements<FilterColumn>());
                Assert.Equal(3U, filterColumn.ColumnId?.Value);
            }

            using var reader = ExcelDocumentReader.Open(memory.ToArray());
            object?[,] values = reader.GetSheet("Orders").ReadRange("A1:E4");
            Assert.Equal("Customer", values[0, 1]);
            Assert.Equal("Contoso", values[1, 1]);
            Assert.Equal(250.75d, Assert.IsType<double>(values[2, 4]));
        }

        [Fact]
        public void Create_ToMemoryStream_ObjectExportSortedFilteredAndFrozen_WritesReadablePackage() {
            var rows = new[] {
                new StreamSalesRow(1, "North", new DateTime(2024, 1, 2), 123.45d),
                new StreamSalesRow(2, "West", new DateTime(2024, 1, 3), 345.67d),
                new StreamSalesRow(3, "South", new DateTime(2024, 1, 4), 234.56d)
            };

            using var memory = new MemoryStream();
            using (var document = ExcelDocument.Create(memory, new ExcelCreateOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose })) {
                document.Execution.SaveWorksheetAfterAutoFit = false;
                var sheet = document.AddWorkSheet("Sales");
                sheet.InsertObjects(rows,
                    ("Id", item => item.Id),
                    ("Region", item => item.Region),
                    ("CreatedOn", item => item.CreatedOn),
                    ("Amount", item => item.Amount));
                sheet.SortUsedRangeByHeader("Amount", ascending: false);
                sheet.AddTable("A1:D4", hasHeader: true, name: "SalesExport", style: OfficeIMO.Excel.TableStyle.TableStyleMedium4, includeAutoFilter: true);
                sheet.AddAutoFilter("A1:D4", new Dictionary<uint, IEnumerable<string>> {
                    { 1U, new[] { "West", "South" } }
                });
                sheet.Freeze(topRows: 1, leftCols: 0);
                sheet.AutoFitColumns();
            }

            Assert.True(memory.Length > 0);
            memory.Position = 0;

            using (var spreadsheet = SpreadsheetDocument.Open(memory, false)) {
                var errors = new OpenXmlValidator().Validate(spreadsheet).ToList();
                Assert.Empty(errors);

                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                Assert.NotNull(wsPart.Worksheet.GetFirstChild<SheetViews>());
                TableDefinitionPart tablePart = Assert.Single(wsPart.TableDefinitionParts);
                Assert.Equal("TableStyleMedium4", tablePart.Table.TableStyleInfo?.Name?.Value);
                AutoFilter tableFilter = Assert.Single(tablePart.Table.Elements<AutoFilter>());
                FilterColumn filterColumn = Assert.Single(tableFilter.Elements<FilterColumn>());
                Assert.Equal(1U, filterColumn.ColumnId?.Value);
            }

            using var reader = ExcelDocumentReader.Open(memory.ToArray());
            object?[,] values = reader.GetSheet("Sales").ReadRange("A1:D4");
            Assert.Equal("West", values[1, 1]);
            Assert.Equal(345.67d, Assert.IsType<double>(values[1, 3]));
            Assert.Equal("South", values[2, 1]);
        }

        [Fact]
        public void Create_ToMemoryStream_WithUnsupportedHyperlink_FallsBackToReadablePackage() {
            using var memory = new MemoryStream();
            using (var document = ExcelDocument.Create(memory, new ExcelCreateOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose })) {
                var sheet = document.AddWorkSheet("Links");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 1, "OfficeIMO");
                sheet.SetHyperlink(2, 1, "https://github.com/evotecit/OfficeIMO", display: "OfficeIMO", style: false);
            }

            Assert.True(memory.Length > 0);
            memory.Position = 0;

            using (var spreadsheet = SpreadsheetDocument.Open(memory, false)) {
                var errors = new OpenXmlValidator().Validate(spreadsheet).ToList();
                Assert.Empty(errors);

                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                Hyperlinks hyperlinks = Assert.Single(wsPart.Worksheet.Elements<Hyperlinks>());
                Hyperlink hyperlink = Assert.Single(hyperlinks.Elements<Hyperlink>());
                Assert.Equal("A2", hyperlink.Reference?.Value);
                Assert.Single(wsPart.HyperlinkRelationships);
            }

            using var reader = ExcelDocumentReader.Open(memory.ToArray());
            object?[,] values = reader.GetSheet("Links").ReadRange("A1:A2");
            Assert.Equal("OfficeIMO", values[1, 0]);
        }

        [Fact]
        public void Create_ToMemoryStream_DataSetExportWithMultipleTables_ReadsBackByTableName() {
            DataSet dataSet = new DataSet("Export");
            DataTable orders = new DataTable("Orders");
            orders.Columns.Add("Id", typeof(int));
            orders.Columns.Add("Customer", typeof(string));
            orders.Columns.Add("Amount", typeof(double));
            orders.Rows.Add(1, "Contoso", 120.5d);
            orders.Rows.Add(2, "Fabrikam", 250.75d);

            DataTable customers = new DataTable("Customers");
            customers.Columns.Add("Id", typeof(int));
            customers.Columns.Add("Name", typeof(string));
            customers.Columns.Add("Active", typeof(bool));
            customers.Rows.Add(10, "Contoso", true);
            customers.Rows.Add(20, "Fabrikam", false);

            dataSet.Tables.Add(orders);
            dataSet.Tables.Add(customers);

            using var memory = new MemoryStream();
            using (var document = ExcelDocument.Create(memory, new ExcelCreateOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose })) {
                var results = document.InsertDataSet(dataSet, autoFit: true);
                Assert.Equal(2, results.Count);
                Assert.Equal("Orders", results[0].TableName);
                Assert.Equal("Customers", results[1].TableName);
            }

            Assert.True(memory.Length > 0);
            memory.Position = 0;

            using (var package = new ZipArchive(memory, ZipArchiveMode.Read, leaveOpen: true)) {
                Assert.NotNull(package.GetEntry("xl/worksheets/sheet1.xml"));
                Assert.NotNull(package.GetEntry("xl/worksheets/sheet2.xml"));
                Assert.NotNull(package.GetEntry("xl/tables/table1.xml"));
                Assert.NotNull(package.GetEntry("xl/tables/table2.xml"));
                Assert.Null(package.GetEntry("xl/sharedStrings.xml"));
            }

            memory.Position = 0;
            using (var spreadsheet = SpreadsheetDocument.Open(memory, false)) {
                var errors = new OpenXmlValidator().Validate(spreadsheet).ToList();
                Assert.Empty(errors);
                Assert.Equal(2, spreadsheet.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>().Count());
                Assert.Equal(2, spreadsheet.WorkbookPart.WorksheetParts.Sum(part => part.TableDefinitionParts.Count()));
            }

            using var reader = ExcelDocumentReader.Open(memory.ToArray());
            var tables = reader.GetTables().OrderBy(table => table.Name).ToList();
            Assert.Equal(new[] { "Customers", "Orders" }, tables.Select(table => table.Name).ToArray());

            DataTable importedOrders = reader.ReadTableAsDataTable("Orders");
            Assert.Equal(2, importedOrders.Rows.Count);
            Assert.Equal("Customer", importedOrders.Columns[1].ColumnName);
            Assert.Equal("Fabrikam", importedOrders.Rows[1]["Customer"]);

            var importedCustomers = reader.ReadTableObjects<CustomerImportRow>("Customers").ToList();
            Assert.Equal(2, importedCustomers.Count);
            Assert.Equal(10, importedCustomers[0].Id);
            Assert.Equal("Contoso", importedCustomers[0].Name);
            Assert.True(importedCustomers[0].Active);

            var streamedCustomers = reader.ReadTableObjectsStream<CustomerImportRow>("Customers").ToList();
            Assert.Equal(importedCustomers.Count, streamedCustomers.Count);
            Assert.Equal("Fabrikam", streamedCustomers[1].Name);
        }

        [Fact]
        public void WriteDataSet_ToMemoryStream_DirectExport_ReadsBackByTableName() {
            DataSet dataSet = new DataSet("Export");
            DataTable orders = new DataTable("Orders");
            orders.Columns.Add("Id", typeof(int));
            orders.Columns.Add("Customer", typeof(string));
            orders.Columns.Add("Amount", typeof(double));
            orders.Rows.Add(1, "Contoso", 120.5d);
            orders.Rows.Add(2, "Fabrikam", 250.75d);

            DataTable customers = new DataTable("Customers");
            customers.Columns.Add("Id", typeof(int));
            customers.Columns.Add("Name", typeof(string));
            customers.Columns.Add("Active", typeof(bool));
            customers.Rows.Add(10, "Contoso", true);
            customers.Rows.Add(20, "Fabrikam", false);

            dataSet.Tables.Add(orders);
            dataSet.Tables.Add(customers);

            using var memory = new MemoryStream();
            var results = ExcelDocument.WriteDataSet(memory, dataSet);
            Assert.Equal(2, results.Count);
            Assert.Equal("Orders", results[0].TableName);
            Assert.Equal("Customers", results[1].TableName);

            using (var package = new ZipArchive(memory, ZipArchiveMode.Read, leaveOpen: true)) {
                Assert.NotNull(package.GetEntry("xl/worksheets/sheet1.xml"));
                Assert.NotNull(package.GetEntry("xl/worksheets/sheet2.xml"));
                Assert.NotNull(package.GetEntry("xl/tables/table1.xml"));
                Assert.NotNull(package.GetEntry("xl/tables/table2.xml"));
                Assert.Null(package.GetEntry("xl/sharedStrings.xml"));
            }

            memory.Position = 0;
            using (var spreadsheet = SpreadsheetDocument.Open(memory, false)) {
                var errors = new OpenXmlValidator().Validate(spreadsheet).ToList();
                Assert.Empty(errors);
            }

            using var reader = ExcelDocumentReader.Open(memory.ToArray());
            DataTable importedOrders = reader.ReadTableAsDataTable("Orders");
            Assert.Equal(2, importedOrders.Rows.Count);
            Assert.Equal("Fabrikam", importedOrders.Rows[1]["Customer"]);

            var importedCustomers = reader.ReadTableObjects<CustomerImportRow>("Customers").ToList();
            Assert.Equal(2, importedCustomers.Count);
            Assert.True(importedCustomers[0].Active);
        }

        private static byte[] CreateStreamSaveWorkbookBytes() {
            using var memory = new MemoryStream();
            using (var document = ExcelDocument.Create(memory, new ExcelCreateOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose })) {
                var sheet = document.AddWorkSheet("StreamData");
                sheet.CellValue(1, 1, "Hello Stream");
                sheet.AutoFitColumns();
            }

            return memory.ToArray();
        }

        private sealed class StreamSalesRow {
            public StreamSalesRow(int id, string region, DateTime createdOn, double amount) {
                Id = id;
                Region = region;
                CreatedOn = createdOn;
                Amount = amount;
            }

            public int Id { get; }

            public string Region { get; }

            public DateTime CreatedOn { get; }

            public double Amount { get; }
        }

        private sealed class CustomerImportRow {
            public int Id { get; set; }

            public string Name { get; set; } = string.Empty;

            public bool Active { get; set; }
        }
    }
}
