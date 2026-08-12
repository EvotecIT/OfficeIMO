using System.Data;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public class ExcelSheetComposerDataTableTests {
        [Fact]
        public void SheetComposer_DataTablePreservesSelectedSchemaAndRichWorkbookFeatures() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            try {
                DataTable members = CreateMembersTable();
                using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                    document.Compose("Summary", composer => {
                        composer.Title("Directory report");
                        composer.Finish(autoFitColumns: false);
                    });
                    document.Compose("Members", composer => {
                        string range = composer.TableFrom(
                            members,
                            title: "Members",
                            configure: options => {
                                options.Columns = new[] { "Enabled", "ADState", "Endpoint", "Missing" };
                                options.MaxCells = 1;
                            },
                            style: ExcelTableStyle.TableStyleMedium2,
                            visuals: options => options.ShowRowStripes = false);
                        Assert.Equal("A2:D4", range);
                        composer.Finish(autoFitColumns: false);
                    });
                    document.Compose("Notes", composer => {
                        composer.Paragraph("Generated after the report table.");
                        composer.Finish(autoFitColumns: false);
                    });
                    document.Save();
                }

                using (ExcelDocument document = ExcelDocument.Load(
                    filePath,
                    new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                    Assert.Equal(3, document.Sheets.Count);
                    ExcelSheet sheet = document["Members"];
                    Assert.Equal("A2:D4", sheet.GetTableRange("Members"));
                    Assert.True(sheet.TryGetCellText(2, 1, out string? enabledHeader));
                    Assert.True(sheet.TryGetCellText(2, 2, out string? adStateHeader));
                    Assert.True(sheet.TryGetCellText(2, 3, out string? endpointHeader));
                    Assert.True(sheet.TryGetCellText(2, 4, out string? missingHeader));
                    Assert.True(sheet.TryGetCellText(3, 1, out string? enabled));
                    Assert.True(sheet.TryGetCellText(3, 2, out string? adState));
                    Assert.True(sheet.TryGetCellText(3, 3, out string? endpoint));
                    Assert.True(sheet.TryGetCellText(3, 4, out string? missing));
                    Assert.Equal("Enabled", enabledHeader);
                    Assert.Equal("AD State", adStateHeader);
                    Assert.Equal("Endpoint", endpointHeader);
                    Assert.Equal("Missing", missingHeader);
                    Assert.Equal("1", enabled);
                    Assert.Equal("Disabled", adState);
                    Assert.Equal("PC-01", endpoint);
                    Assert.True(string.IsNullOrEmpty(missing));
                }

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                    WorksheetPart membersPart = spreadsheet.WorkbookPart!.WorksheetParts.ElementAt(1);
                    TableDefinitionPart tablePart = Assert.Single(membersPart.TableDefinitionParts);
                    TableStyleInfo style = Assert.IsType<TableStyleInfo>(tablePart.Table!.TableStyleInfo);
                    Assert.Equal("TableStyleMedium2", style.Name?.Value);
                    Assert.False(style.ShowRowStripes?.Value);
                }
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Fact]
        public void SheetComposer_DataTableUsesExcelRowBoundaryInsteadOfGenericCellLimit() {
            DataTable members = CreateMembersTable();
            using ExcelDocument document = ExcelDocument.Create();

            document.Compose("Members", composer => {
                string range = composer.TableFrom(
                    members,
                    configure: options => options.MaxCells = 1,
                    freezeHeaderRow: false);
                Assert.Equal("A1:D3", range);
            });
        }

        [Fact]
        public void SheetComposer_DataTableCanonicalizesUniqueExplicitColumnCasing() {
            DataTable members = CreateMembersTable();
            using ExcelDocument document = ExcelDocument.Create();

            document.Compose("Members", composer => {
                string range = composer.TableFrom(
                    members,
                    configure: options => options.Columns = new[] { "adstate" },
                    freezeHeaderRow: false);
                Assert.Equal("A1:A3", range);
            });

            ExcelSheet sheet = document["Members"];
            Assert.True(sheet.TryGetCellText(1, 1, out string? header));
            Assert.True(sheet.TryGetCellText(2, 1, out string? value));
            Assert.Equal("AD State", header);
            Assert.Equal("Disabled", value);
        }

        [Fact]
        public void SheetComposer_DataTableExplicitColumnsDoNotMatchLastPathSegments() {
            var table = new DataTable("Members");
            table.Columns.Add("Customer.Name", typeof(string));
            table.Rows.Add("Alice");
            using ExcelDocument document = ExcelDocument.Create();

            document.Compose("Members", composer => {
                string range = composer.TableFrom(
                    table,
                    configure: options => options.Columns = new[] { "Name" },
                    freezeHeaderRow: false);
                Assert.Equal("A1:A2", range);
            });

            ExcelSheet sheet = document["Members"];
            Assert.True(sheet.TryGetCellText(1, 1, out string? header));
            Assert.True(sheet.TryGetCellText(2, 1, out string? value));
            Assert.Equal("Name", header);
            Assert.True(string.IsNullOrEmpty(value));
        }

        [Fact]
        public void SheetComposer_DataTableTreatsDottedColumnNamesAsLiteralSchema() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            try {
                var table = new DataTable("Members");
                table.Columns.Add("Customer.Name", typeof(string));
                table.Rows.Add("Alice");
                using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                    document.Compose("Members", composer =>
                        composer.TableFrom(table, freezeHeaderRow: false));
                    document.Save();
                }

                using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
                Worksheet worksheet = spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet;
                Assert.Empty(worksheet.Descendants<ConditionalFormatting>());
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Fact]
        public void SheetComposer_EmptyObjectTableSkipsExemptProjectionLimits() {
            using ExcelDocument document = ExcelDocument.Create();

            document.Compose("Data", composer => {
                string range = composer.TableFrom(
                    Array.Empty<object>(),
                    configure: options => {
                        options.Columns = new[] { "First", "Second" };
                        options.MaxColumns = 1;
                    },
                    freezeHeaderRow: false);
                Assert.Equal("A1:A1", range);
            });

            Assert.True(document["Data"].TryGetCellText(1, 1, out string? text));
            Assert.Equal("(no data)", text);
        }

        [Fact]
        public void SheetComposer_ObjectTableExplainsHardExcelRowBoundary() {
            using ExcelDocument document = ExcelDocument.Create();

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                document.Compose("Data", composer => composer.TableFrom(
                    new CountOnlyRows(A1.MaxRows),
                    configure: options => options.MaxRows = int.MaxValue,
                    freezeHeaderRow: false)));

            Assert.Contains("room for at most 1048575 data rows", exception.Message, StringComparison.Ordinal);
            Assert.Contains("split the data across multiple worksheets", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("cannot be overridden", exception.Message, StringComparison.Ordinal);
            Assert.DoesNotContain("options.MaxRows = 1048576", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void SheetComposer_DataTableFallsBackWhenLaterWorksheetContentExists() {
            DataTable members = CreateMembersTable();
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            try {
                using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                    document.Compose("Members", composer => {
                        composer.Sheet.Cell(100, 1, "Footer");
                        string range = composer.TableFrom(members, freezeHeaderRow: false);
                        Assert.Equal("A1:D3", range);
                    });
                    document.Save();
                }

                using ExcelDocument reloaded = ExcelDocument.Load(
                    filePath,
                    new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
                ExcelSheet sheet = reloaded["Members"];
                Assert.True(sheet.TryGetCellText(2, 1, out string? endpoint));
                Assert.True(sheet.TryGetCellText(100, 1, out string? footer));
                Assert.Equal("PC-01", endpoint);
                Assert.Equal("Footer", footer);
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Fact]
        public void SheetComposer_DataTablePreservesCaseDistinctSchemaColumns() {
            var table = new DataTable("CaseDistinct");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("name", typeof(string));
            table.Rows.Add("Upper", "Lower");
            using ExcelDocument document = ExcelDocument.Create();

            document.Compose("Data", composer => {
                string range = composer.TableFrom(table, freezeHeaderRow: false);
                Assert.Equal("A1:B2", range);
            });

            ExcelSheet sheet = document["Data"];
            Assert.True(sheet.TryGetCellText(1, 1, out string? firstHeader));
            Assert.True(sheet.TryGetCellText(1, 2, out string? secondHeader));
            Assert.True(sheet.TryGetCellText(2, 1, out string? firstValue));
            Assert.True(sheet.TryGetCellText(2, 2, out string? secondValue));
            Assert.Equal("Name", firstHeader);
            Assert.Equal("Name (2)", secondHeader);
            Assert.Equal("Upper", firstValue);
            Assert.Equal("Lower", secondValue);
        }

        [Fact]
        public void SheetComposer_DataTablePreservesExplicitCaseDistinctColumnsAndRejectsAmbiguousFallback() {
            var table = new DataTable("CaseDistinct");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("name", typeof(string));
            table.Rows.Add("Upper", "Lower");

            using (var document = ExcelDocument.Create()) {
                document.Compose("Data", composer => {
                    string range = composer.TableFrom(
                        table,
                        configure: options => options.Columns = new[] { "Name", "name" },
                        freezeHeaderRow: false);
                    Assert.Equal("A1:B2", range);
                });

                ExcelSheet sheet = document["Data"];
                Assert.True(sheet.TryGetCellText(2, 1, out string? firstValue));
                Assert.True(sheet.TryGetCellText(2, 2, out string? secondValue));
                Assert.Equal("Upper", firstValue);
                Assert.Equal("Lower", secondValue);
            }

            using var ambiguousDocument = ExcelDocument.Create();
            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                ambiguousDocument.Compose("Data", composer =>
                    composer.TableFrom(
                        table,
                        configure: options => options.Columns = new[] { "NAME" },
                        freezeHeaderRow: false)));
            Assert.Contains("ambiguous", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("exact column casing", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void SheetComposer_DataTableProjectionRulesPreserveAndOrderCaseDistinctColumns() {
            var table = new DataTable("CaseDistinct");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("name", typeof(string));
            table.Columns.Add("Value", typeof(int));
            table.Rows.Add("Upper", "Lower", 42);
            using ExcelDocument document = ExcelDocument.Create();

            document.Compose("Data", composer => {
                string range = composer.TableFrom(
                    table,
                    configure: options => {
                        options.Ignore = new[] { "Missing" };
                        options.PinnedFirst = new[] { "name" };
                        options.PropertyPriority["Value"] = -1;
                    },
                    freezeHeaderRow: false);
                Assert.Equal("A1:C2", range);
            });

            ExcelSheet sheet = document["Data"];
            Assert.True(sheet.TryGetCellText(1, 1, out string? firstHeader));
            Assert.True(sheet.TryGetCellText(1, 2, out string? secondHeader));
            Assert.True(sheet.TryGetCellText(1, 3, out string? thirdHeader));
            Assert.True(sheet.TryGetCellText(2, 1, out string? firstValue));
            Assert.True(sheet.TryGetCellText(2, 2, out string? secondValue));
            Assert.True(sheet.TryGetCellText(2, 3, out string? thirdValue));
            Assert.Equal("Name", firstHeader);
            Assert.Equal("Value", secondHeader);
            Assert.Equal("Name (2)", thirdHeader);
            Assert.Equal("Lower", firstValue);
            Assert.Equal("42", secondValue);
            Assert.Equal("Upper", thirdValue);
        }

        [Fact]
        public void SheetComposer_DataTableProjectionRulesRejectAmbiguousCaseInsensitiveTargets() {
            var table = new DataTable("CaseDistinct");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("name", typeof(string));
            table.Rows.Add("Upper", "Lower");
            using ExcelDocument document = ExcelDocument.Create();

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                document.Compose("Data", composer =>
                    composer.TableFrom(
                        table,
                        configure: options => options.PinnedFirst = new[] { "NAME" },
                        freezeHeaderRow: false)));

            Assert.Contains("ambiguous", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("exact full column name and casing", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void SheetComposer_DataTableFiltersCaseDistinctColumnsByExactSchemaIdentity() {
            var table = new DataTable("CaseDistinct");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("name", typeof(string));
            table.Columns.Add("Value", typeof(int));
            table.Rows.Add("Upper", "Lower", 42);

            using (var includeDocument = ExcelDocument.Create()) {
                includeDocument.Compose("Data", composer => composer.TableFrom(
                    table,
                    configure: options => options.IncludeProperties = new[] { "Name" },
                    freezeHeaderRow: false));
                ExcelSheet sheet = includeDocument["Data"];
                Assert.True(sheet.TryGetCellText(2, 1, out string? value));
                Assert.Equal("Upper", value);
                Assert.False(sheet.TryGetCellText(1, 2, out _));
            }

            foreach (bool useIgnore in new[] { false, true }) {
                using var excludeDocument = ExcelDocument.Create();
                excludeDocument.Compose("Data", composer => composer.TableFrom(
                    table,
                    configure: options => {
                        if (useIgnore) options.Ignore = new[] { "Name" };
                        else options.ExcludeProperties = new[] { "Name" };
                    },
                    freezeHeaderRow: false));
                ExcelSheet sheet = excludeDocument["Data"];
                Assert.True(sheet.TryGetCellText(2, 1, out string? lower));
                Assert.True(sheet.TryGetCellText(2, 2, out string? number));
                Assert.Equal("Lower", lower);
                Assert.Equal("42", number);
            }

            using var ambiguousDocument = ExcelDocument.Create();
            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                ambiguousDocument.Compose("Data", composer => composer.TableFrom(
                    table,
                    configure: options => options.ExcludeProperties = new[] { "NAME" },
                    freezeHeaderRow: false)));
            Assert.Contains("ambiguous", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("ExcludeProperties", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void SheetComposer_TitledDataTableAppliesVisualsToItsActualHeaderRow() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            try {
                DataTable members = CreateMembersTable();
                using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                    document.Compose("Members", composer => {
                        string range = composer.TableFrom(
                            members,
                            title: "Members",
                            freezeHeaderRow: false,
                            visuals: options => options.NumericColumnFormats["Enabled"] = "0.0000");
                        Assert.Equal("A2:D4", range);
                    });
                    document.Save();
                }

                using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
                WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                Cell dataCell = worksheetPart.Worksheet.Descendants<Cell>()
                    .Single(cell => cell.CellReference?.Value == "B3");
                Assert.NotNull(dataCell.StyleIndex);

                Stylesheet stylesheet = workbookPart.WorkbookStylesPart!.Stylesheet!;
                CellFormat cellFormat = stylesheet.CellFormats!.Elements<CellFormat>()
                    .ElementAt((int)dataCell.StyleIndex!.Value);
                Assert.True(cellFormat.ApplyNumberFormat?.Value);
                NumberingFormat numberFormat = stylesheet.NumberingFormats!.Elements<NumberingFormat>()
                    .Single(format => format.NumberFormatId?.Value == cellFormat.NumberFormatId?.Value);
                Assert.Equal("0.0000", numberFormat.FormatCode?.Value);
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Fact]
        public void SheetComposer_DataTableExplainsHardExcelColumnBoundary() {
            var table = new DataTable("Wide");
            for (int index = 0; index <= A1.MaxColumns; index++) {
                table.Columns.Add("Column" + index, typeof(string));
            }
            table.Rows.Add(table.NewRow());
            using ExcelDocument document = ExcelDocument.Create();

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                document.Compose("Data", composer => composer.TableFrom(table, freezeHeaderRow: false)));

            Assert.Contains("requires at least 16385 columns", exception.Message, StringComparison.Ordinal);
            Assert.Contains("Select fewer columns or split the data across multiple worksheets", exception.Message, StringComparison.Ordinal);
            Assert.Contains("cannot be overridden", exception.Message, StringComparison.Ordinal);
            Assert.DoesNotContain("options.MaxColumns = 16385", exception.Message, StringComparison.Ordinal);
        }

        private static DataTable CreateMembersTable() {
            var table = new DataTable("Members");
            table.Columns.Add("Endpoint", typeof(string));
            table.Columns.Add("Enabled", typeof(bool));
            table.Columns.Add("ADState", typeof(string));
            table.Columns.Add("Ignored", typeof(string));
            table.Rows.Add("PC-01", true, "Disabled", "one");
            table.Rows.Add("PC-02", false, "Gone", "two");
            return table;
        }

        private sealed class CountOnlyRows : IReadOnlyCollection<int> {
            internal CountOnlyRows(int count) {
                Count = count;
            }

            public int Count { get; }

            public IEnumerator<int> GetEnumerator() =>
                throw new InvalidOperationException("The known row-count check must run before enumeration.");

            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
        }
    }
}
