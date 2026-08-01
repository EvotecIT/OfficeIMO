using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using System.Globalization;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void PdfTables_SaveTablesAsExcel_ImportsDetectedTablesAsWorkbookTables() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "A-100", "Alpha", "2" },
                new[] { "B-200", "Beta", "14" }
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 70, 170, 60 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .ToBytes();

        using var workbook = new MemoryStream();
        PdfExcelTableImportReport report = PdfExcelTableConverterExtensions.SaveTablesAsExcel(
            LoadTables(pdf),
            workbook,
            new PdfExcelTableImportOptions {
                AutoFitColumns = false
            });

        PdfExcelTableImportEntry result = Assert.Single(report.Entries);
        Assert.Equal(1, result.PageNumber);
        Assert.Equal(0, result.TableIndex);
        Assert.Equal(3, result.ColumnCount);
        Assert.Equal(2, result.RowCount);
        Assert.False(result.Truncated);
        Assert.Equal("A1:C3", result.Range);

        byte[] workbookBytes = workbook.ToArray();
        using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(workbookBytes), false)) {
            WorksheetPart worksheet = Assert.Single(spreadsheet.WorkbookPart!.WorksheetParts);
            TableDefinitionPart tablePart = Assert.Single(worksheet.TableDefinitionParts);
            Table tableDefinition = tablePart.Table!;
            Assert.Equal(result.TableName, tableDefinition.Name?.Value);
            Assert.Equal("A1:C3", tableDefinition.Reference?.Value);
            Assert.NotNull(tableDefinition.GetFirstChild<AutoFilter>());
        }

        using ExcelDocumentReader reader = ExcelDocumentReader.Open(workbookBytes);
        ExcelTableInfo table = Assert.Single(reader.GetTables());
        Assert.Equal(result.TableName, table.Name);
        Assert.Equal(result.SheetName, table.SheetName);
        Assert.Equal(new[] { "Code", "Name", "Qty" }, table.Columns.Select(column => column.Name).ToArray());

        object?[,] values = reader.GetSheet(result.SheetName).ReadRange(result.Range);
        Assert.Equal("Code", values[0, 0]);
        Assert.Equal("A-100", values[1, 0]);
        Assert.Equal(14d, Convert.ToDouble(values[2, 2], CultureInfo.InvariantCulture));
    }

    [Fact]
    public void PdfTables_SaveTablesAsExcel_SupportsNonSeekableDestinationStreams() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Table(new[] {
                new[] { "Code", "Qty" },
                new[] { "A-100", "2" },
                new[] { "B-200", "3" }
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 180, 80 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .ToBytes();
        using var workbook = new NonSeekableReadWriteBuffer(Array.Empty<byte>());

        PdfExcelTableImportReport report = PdfExcelTableConverterExtensions.SaveTablesAsExcel(
            LoadTables(pdf),
            workbook,
            new PdfExcelTableImportOptions {
                AutoFitColumns = false
            });

        PdfExcelTableImportEntry result = Assert.Single(report.Entries);
        using ExcelDocumentReader reader = ExcelDocumentReader.Open(workbook.ToArray());
        object?[,] values = reader.GetSheet(result.SheetName).ReadRange(result.Range);
        Assert.Equal("A-100", values[1, 0]);
        Assert.Equal(2d, Convert.ToDouble(values[1, 1], CultureInfo.InvariantCulture));
    }

    [Fact]
    public void PdfTables_SaveTablesAsExcel_WritesDetectedNumericColumnsAsNumberCells() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "A-100", "Alpha", "2" },
                new[] { "B-200", "Beta", "14" }
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 70, 170, 60 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .ToBytes();

        using var workbook = new MemoryStream();
        PdfExcelTableImportReport report = PdfExcelTableConverterExtensions.SaveTablesAsExcel(
            LoadTables(pdf),
            workbook,
            new PdfExcelTableImportOptions {
                AutoFitColumns = false
            });

        PdfExcelTableImportEntry result = Assert.Single(report.Entries);
        byte[] workbookBytes = workbook.ToArray();
        using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(workbookBytes), false)) {
            SheetData sheetData = GetOnlySheetData(spreadsheet);
            Cell codeCell = GetCell(sheetData, "A2");
            Cell nameCell = GetCell(sheetData, "B2");
            Cell quantityCell = GetCell(sheetData, "C2");
            Cell secondQuantityCell = GetCell(sheetData, "C3");

            Assert.True(IsTextCell(codeCell));
            Assert.True(IsTextCell(nameCell));
            Assert.True(quantityCell.DataType == null || quantityCell.DataType.Value == CellValues.Number);
            Assert.True(secondQuantityCell.DataType == null || secondQuantityCell.DataType.Value == CellValues.Number);
            Assert.Equal("2", quantityCell.CellValue?.Text);
            Assert.Equal("14", secondQuantityCell.CellValue?.Text);
        }

        using ExcelDocumentReader reader = ExcelDocumentReader.Open(workbookBytes);
        object?[,] values = reader.GetSheet(result.SheetName).ReadRange(result.Range);
        Assert.Equal("A-100", values[1, 0]);
        Assert.Equal("Alpha", values[1, 1]);
        Assert.Equal(2d, Convert.ToDouble(values[1, 2], CultureInfo.InvariantCulture));

        using var textWorkbook = new MemoryStream();
        PdfExcelTableConverterExtensions.SaveTablesAsExcel(
            LoadTables(pdf),
            textWorkbook,
            new PdfExcelTableImportOptions {
                AutoFitColumns = false,
                ConvertNumericColumns = false
            });

        using SpreadsheetDocument textSpreadsheet = SpreadsheetDocument.Open(new MemoryStream(textWorkbook.ToArray()), false);
        SheetData textSheetData = GetOnlySheetData(textSpreadsheet);
        Assert.True(IsTextCell(GetCell(textSheetData, "C2")));
    }

    [Fact]
    public void PdfTables_SaveTablesAsExcel_MergesPageContinuationsAndCollapsesRepeatedHeaders() {
        var rows = new List<string[]> {
            new[] { "Group", "State" },
            new[] { "Metric", "Owner" }
        };
        for (int index = 1; index <= 30; index++) {
            rows.Add(new[] {
                "Check " + index.ToString(CultureInfo.InvariantCulture),
                "Team " + index.ToString(CultureInfo.InvariantCulture)
            });
        }

        var style = new PdfCore.PdfTableStyle {
            HeaderRowCount = 2,
            RepeatHeaderRowCount = 2,
            ColumnWidthPoints = new List<double?> { 120, 120 },
            CellPaddingX = 5,
            CellPaddingY = 3
        };
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 9
            })
            .Table(rows, style: style)
            .ToBytes();

        PdfCore.PdfLogicalDocument logical = LoadTables(pdf);
        PdfCore.PdfLogicalTableContinuationGroup bounded = Assert.Single(
            PdfCore.PdfLogicalTableContinuations.Group(logical, 2, true, true, 64, 4D));
        Assert.All(bounded.Segments, segment => Assert.InRange(segment.Data.Rows.Count, 0, 6));
        Assert.Equal(30, bounded.TotalRowCount);
        Assert.Equal(2, bounded.Rows.Count);
        Assert.True(bounded.Truncated);
        using var workbook = new MemoryStream();
        PdfExcelTableImportReport report = logical.SaveTablesAsExcel(
            workbook,
            new PdfExcelTableImportOptions {
                AutoFitColumns = false,
                SuppressRepeatedBodyHeaderRows = true
            });

        string tableDetails = string.Join("; ", logical.Pages.SelectMany((page, pageIndex) => page.Tables.Select((table, tableIndex) =>
            $"P{page.PageNumber}/T{tableIndex}: {table.YTop:0.##}-{table.YBottom:0.##}, {table.DetectionKind}, rows={table.Rows.Count}, headers={string.Join("|", PdfCore.PdfLogicalTableAnalysis.Extract(table).Columns)}")));
        Assert.True(report.Entries.Count == 1, tableDetails);
        PdfExcelTableImportEntry entry = report.Entries[0];
        Assert.True(entry.SourceTableCount > 1);
        Assert.Equal(entry.SourceTableCount, entry.SourcePageNumbers.Count);
        Assert.Equal(Enumerable.Range(1, entry.SourceTableCount), entry.SourcePageNumbers);
        Assert.Equal(30, entry.RowCount);
        Assert.Equal(30, entry.TotalRowCount);
        Assert.False(entry.Truncated);
        Assert.Equal(1, entry.AdditionalHeaderRowCount);
        Assert.Equal(entry.SourceTableCount - 1, entry.SuppressedRepeatedHeaderRows);

        using ExcelDocumentReader reader = ExcelDocumentReader.Open(workbook.ToArray());
        ExcelTableInfo table = Assert.Single(reader.GetTables());
        Assert.Equal(new[] { "Group / Metric", "State / Owner" }, table.Columns.Select(column => column.Name).ToArray());
        object?[,] values = reader.GetSheet(entry.SheetName).ReadRange(entry.Range);
        Assert.Equal("Check 1", values[1, 0]);
        Assert.Equal("Check 30", values[30, 0]);
    }

    [Fact]
    public void PdfTables_SaveTablesAsExcel_MergesHeaderlessPageContinuationsUsingPrimaryColumns() {
        var rows = new List<string[]> { new[] { "Item", "Quantity" } };
        for (int index = 1; index <= 30; index++) {
            rows.Add(new[] {
                "Entry " + index.ToString(CultureInfo.InvariantCulture),
                index.ToString(CultureInfo.InvariantCulture)
            });
        }

        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 9
            })
            .Table(rows, style: new PdfCore.PdfTableStyle {
                HeaderRowCount = 1,
                RepeatHeaderRowCount = 0,
                ColumnWidthPoints = new List<double?> { 160, 80 },
                CellPaddingX = 5,
                CellPaddingY = 3
            })
            .ToBytes();

        PdfCore.PdfLogicalDocument logical = LoadTables(pdf);
        using var workbook = new MemoryStream();
        PdfExcelTableImportReport report = logical.SaveTablesAsExcel(
            workbook,
            new PdfExcelTableImportOptions {
                AutoFitColumns = false,
                ContinuationGeometryTolerancePoints = 8D
            });

        string details = string.Join("; ", logical.Pages.SelectMany((page, pageIndex) => page.Tables.Select((table, tableIndex) => {
            PdfCore.PdfLogicalTableData data = PdfCore.PdfLogicalTableAnalysis.Extract(table);
            return $"P{page.PageNumber}/T{tableIndex}: {table.YTop:0.##}-{table.YBottom:0.##}, {table.DetectionKind}, header={data.Structure.HasHeaderRow}, columns={string.Join("|", data.Columns)}, geometry={string.Join("|", table.Columns.Select(column => $"{column.From:0.##}-{column.To:0.##}"))}, rows={data.Rows.Count}";
        })));
        Assert.True(report.Entries.Count == 1, details);
        PdfExcelTableImportEntry entry = report.Entries[0];
        Assert.True(entry.SourceTableCount > 1);
        Assert.Equal(30, entry.RowCount);
        Assert.Equal(0, entry.SuppressedRepeatedHeaderRows);
        using ExcelDocumentReader reader = ExcelDocumentReader.Open(workbook.ToArray());
        ExcelTableInfo table = Assert.Single(reader.GetTables());
        Assert.Equal(new[] { "Item", "Quantity" }, table.Columns.Select(column => column.Name).ToArray());
        object?[,] values = reader.GetSheet(entry.SheetName).ReadRange(entry.Range);
        Assert.Equal("Entry 1", values[1, 0]);
        Assert.Equal("Entry 30", values[30, 0]);
    }

    [Fact]
    public void PdfTables_SaveTablesAsExcel_DoesNotSuppressRepeatedOrdinaryRowsOnMergedContinuations() {
        var rows = new List<string[]> { new[] { "Status", "Owner" } };
        for (int index = 0; index < 30; index++) rows.Add(new[] { "Pending", "Team A" });

        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 9
            })
            .Table(rows, style: new PdfCore.PdfTableStyle {
                HeaderRowCount = 1,
                RepeatHeaderRowCount = 1,
                ColumnWidthPoints = new List<double?> { 160, 80 },
                CellPaddingX = 5,
                CellPaddingY = 3
            })
            .ToBytes();

        using var workbook = new MemoryStream();
        PdfExcelTableImportEntry entry = Assert.Single(LoadTables(pdf).SaveTablesAsExcel(
            workbook,
            new PdfExcelTableImportOptions {
                AutoFitColumns = false,
                ContinuationGeometryTolerancePoints = 8D
            }).Entries);

        Assert.True(entry.SourceTableCount > 1);
        Assert.Equal(30, entry.RowCount);
        Assert.Equal(30, entry.TotalRowCount);
        Assert.Equal(0, entry.AdditionalHeaderRowCount);
        Assert.Equal(0, entry.SuppressedRepeatedHeaderRows);
        using ExcelDocumentReader reader = ExcelDocumentReader.Open(workbook.ToArray());
        object?[,] values = reader.GetSheet(entry.SheetName).ReadRange(entry.Range);
        Assert.Equal("Pending", values[30, 0]);
        Assert.Equal("TeamA", values[30, 1]);
    }

    [Fact]
    public void PdfTables_SaveTablesAsExcel_WritesBooleanPercentageDateAndNumericColumnsAsTypedCells() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 9
            })
            .Table(new[] {
                new[] { "Active", "Completion" },
                new[] { "Yes", "12.5%" },
                new[] { "No", "100.0%" },
                new[] { "Yes", "50.0%" },
                new[] { "No", "00.0%" }
            }, style: new PdfCore.PdfTableStyle {
                HeaderRowCount = 1,
                ColumnWidthPoints = new List<double?> { 120, 120 }
            })
            .PageBreak()
            .Table(new[] {
                new[] { "Due Date", "Quantity" },
                new[] { "2026-07-01", "2" },
                new[] { "2026-07-31", "14" },
                new[] { "2026-08-15", "8" },
                new[] { "2026-09-01", "21" }
            }, style: new PdfCore.PdfTableStyle {
                HeaderRowCount = 1,
                ColumnWidthPoints = new List<double?> { 160, 80 }
            })
            .ToBytes();

        PdfCore.PdfLogicalDocument logical = LoadTables(pdf);
        using var workbook = new MemoryStream();
        PdfExcelTableImportReport report = logical.SaveTablesAsExcel(
            workbook,
            new PdfExcelTableImportOptions { AutoFitColumns = false });

        Assert.Equal(2, report.Entries.Count);
        Assert.Equal(new[] {
            PdfExcelTableColumnKind.Boolean,
            PdfExcelTableColumnKind.Percentage,
            PdfExcelTableColumnKind.DateTime,
            PdfExcelTableColumnKind.Number
        }, report.Entries.SelectMany(entry => entry.ColumnKinds));

        byte[] workbookBytes = workbook.ToArray();
        using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(workbookBytes), false)) {
            WorksheetPart percentageWorksheet = spreadsheet.WorkbookPart!.WorksheetParts.Single(worksheet =>
                worksheet.TableDefinitionParts.Any(tablePart =>
                    string.Equals(tablePart.Table?.Name?.Value, report.Entries[0].TableName, StringComparison.Ordinal)));
            SheetData sheetData = percentageWorksheet.Worksheet.GetFirstChild<SheetData>()!;
            Cell percentageCell = GetCell(sheetData, "B2");
            Stylesheet stylesheet = spreadsheet.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            CellFormat percentageFormat = stylesheet.CellFormats!
                .Elements<CellFormat>()
                .ElementAt((int)(percentageCell.StyleIndex?.Value ?? 0U));
            uint percentageFormatId = percentageFormat.NumberFormatId?.Value ?? 0U;
            string? percentageFormatCode = stylesheet.NumberingFormats?
                .Elements<NumberingFormat>()
                .SingleOrDefault(format => format.NumberFormatId?.Value == percentageFormatId)?
                .FormatCode?.Value;
            Assert.Equal("0.00%", percentageFormatCode);
        }

        using ExcelDocumentReader reader = ExcelDocumentReader.Open(workbookBytes);
        object?[,] statusValues = reader.GetSheet(report.Entries[0].SheetName).ReadRange(report.Entries[0].Range);
        Assert.Equal(true, statusValues[1, 0]);
        Assert.Equal(0.125d, Convert.ToDouble(statusValues[1, 1], CultureInfo.InvariantCulture), 8);
        object?[,] dueValues = reader.GetSheet(report.Entries[1].SheetName).ReadRange(report.Entries[1].Range);
        Assert.Equal(new DateTime(2026, 7, 1), Convert.ToDateTime(dueValues[1, 0], CultureInfo.InvariantCulture));
        Assert.Equal(14d, Convert.ToDouble(dueValues[2, 1], CultureInfo.InvariantCulture));
    }

    [Fact]
    public void PdfTables_SaveTablesAsExcel_DoesNotInferRatiosAsDates() {
        byte[] pdf = PdfCore.PdfDocument.Create()
            .Table(new[] {
                new[] { "Ratio", "Label" },
                new[] { "1/2", "Half" },
                new[] { "3/4", "Three quarters" }
            }, style: new PdfCore.PdfTableStyle {
                HeaderRowCount = 1,
                ColumnWidthPoints = new List<double?> { 120, 160 }
            })
            .ToBytes();

        using var workbook = new MemoryStream();
        PdfExcelTableImportReport report = LoadTables(pdf).SaveTablesAsExcel(
            workbook,
            new PdfExcelTableImportOptions { AutoFitColumns = false });

        PdfExcelTableImportEntry entry = Assert.Single(report.Entries);
        Assert.Equal(PdfExcelTableColumnKind.Text, entry.ColumnKinds[0]);
        using ExcelDocumentReader reader = ExcelDocumentReader.Open(workbook.ToArray());
        object?[,] values = reader.GetSheet(entry.SheetName).ReadRange(entry.Range);
        Assert.Equal("1/2", values[1, 0]);
        Assert.Equal("3/4", values[2, 0]);
    }

    [Fact]
    public void PdfTables_SaveTablesAsExcel_ImportsTimeOnlyValuesWithoutInventingDates() {
        byte[] pdf = PdfCore.PdfDocument.Create()
            .Table(new[] {
                new[] { "Start", "End", "Label" },
                new[] { "09:30", "10:15", "Morning" },
                new[] { "14:45", "16:30", "Afternoon" }
            }, style: new PdfCore.PdfTableStyle {
                HeaderRowCount = 1,
                ColumnWidthPoints = new List<double?> { 90, 90, 160 }
            })
            .ToBytes();

        using var workbook = new MemoryStream();
        PdfExcelTableImportEntry entry = Assert.Single(LoadTables(pdf).SaveTablesAsExcel(
            workbook,
            new PdfExcelTableImportOptions { AutoFitColumns = false }).Entries);

        Assert.Equal(PdfExcelTableColumnKind.Time, entry.ColumnKinds[0]);
        Assert.Equal(PdfExcelTableColumnKind.Time, entry.ColumnKinds[1]);
        using ExcelDocumentReader reader = ExcelDocumentReader.Open(workbook.ToArray());
        object?[,] values = reader.GetSheet(entry.SheetName).ReadRange(entry.Range);
        DateTime first = Assert.IsType<DateTime>(values[1, 0]);
        DateTime second = Assert.IsType<DateTime>(values[2, 0]);
        Assert.Equal(new DateTime(1899, 12, 30), first.Date);
        Assert.Equal(new DateTime(1899, 12, 30), second.Date);
        Assert.Equal(TimeSpan.FromHours(9.5D), first.TimeOfDay);
        Assert.Equal(TimeSpan.FromHours(14.75D), second.TimeOfDay);
    }

    [Theory]
    [InlineData("Vendor")]
    [InlineData("Trend")]
    [InlineData("Spend")]
    public void PdfTables_SaveTablesAsExcel_DoesNotMatchDateHintsInsideHeaderWords(string header) {
        byte[] pdf = PdfCore.PdfDocument.Create()
            .Table(new[] {
                new[] { header, "Label" },
                new[] { "1/2", "First" },
                new[] { "3/4", "Second" }
            }, style: new PdfCore.PdfTableStyle {
                HeaderRowCount = 1,
                ColumnWidthPoints = new List<double?> { 120, 160 }
            })
            .ToBytes();

        using var workbook = new MemoryStream();
        PdfExcelTableImportEntry entry = Assert.Single(LoadTables(pdf).SaveTablesAsExcel(
            workbook,
            new PdfExcelTableImportOptions { AutoFitColumns = false }).Entries);

        Assert.Equal(PdfExcelTableColumnKind.Text, entry.ColumnKinds[0]);
        using ExcelDocumentReader reader = ExcelDocumentReader.Open(workbook.ToArray());
        object?[,] values = reader.GetSheet(entry.SheetName).ReadRange(entry.Range);
        Assert.Equal("1/2", values[1, 0]);
        Assert.Equal("3/4", values[2, 0]);
    }

    [Theory]
    [InlineData("01/02/2025", "03/04/2025")]
    [InlineData("01-02-2025", "03-04-2025")]
    [InlineData("01.02.2025", "03.04.2025")]
    public void PdfTables_SaveTablesAsExcel_KeepsAmbiguousNumericDatesAsText(string value, string secondValue) {
        byte[] pdf = PdfCore.PdfDocument.Create()
            .Table(new[] {
                new[] { "Value", "Label" },
                new[] { value, "First" },
                new[] { secondValue, "Second" }
            }, style: new PdfCore.PdfTableStyle {
                HeaderRowCount = 1,
                ColumnWidthPoints = new List<double?> { 120, 160 }
            })
            .ToBytes();

        using var workbook = new MemoryStream();
        PdfExcelTableImportEntry entry = Assert.Single(LoadTables(pdf).SaveTablesAsExcel(
            workbook,
            new PdfExcelTableImportOptions { AutoFitColumns = false }).Entries);

        Assert.Equal(PdfExcelTableColumnKind.Text, entry.ColumnKinds[0]);
        using ExcelDocumentReader reader = ExcelDocumentReader.Open(workbook.ToArray());
        object?[,] values = reader.GetSheet(entry.SheetName).ReadRange(entry.Range);
        Assert.Equal(value, values[1, 0]);
    }

    [Fact]
    public void PdfTables_SaveTablesAsExcel_KeepsPositionedCellRecoveryBoundedAndTableOnly() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 640,
                PageHeight = 280,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 9
            })
            .Table(new[] {
                new[] { "Active", "Completion Percentage", "Due Date", "Quantity" },
                new[] { "Yes", "12.5%", "2026-07-01", "2" },
                new[] { "No", "100%", "2026-07-31", "14" }
            }, style: new PdfCore.PdfTableStyle {
                HeaderRowCount = 1,
                ColumnWidthPoints = new List<double?> { 80, 160, 120, 80 }
            })
            .ToBytes();

        PdfCore.PdfLogicalDocument logical = LoadTables(pdf);
        using var workbook = new MemoryStream();
        PdfExcelTableImportReport report = logical.SaveTablesAsExcel(
            workbook,
            new PdfExcelTableImportOptions { AutoFitColumns = false });

        IReadOnlyList<PdfCore.PdfTextSpan> spans = logical.Pages[0].TextBlocks.SelectMany(static block => block.Spans).ToArray();
        List<PdfCore.TextLayoutEngine.TextLine> lines = PdfCore.TextLayoutEngine.BuildLines(spans);
        List<PdfCore.StructuredTable> recovered = PdfCore.TableDetector.DetectPositionedCellTables(lines);
        PdfCore.StructuredTable positionedTable = Assert.Single(recovered);
        Assert.Equal("positioned-cells-bounded", positionedTable.Kind);
        Assert.Equal(3, positionedTable.Rows.Count);
        PdfExcelTableImportEntry entry = Assert.Single(report.Entries);
        Assert.Equal(4, entry.ColumnCount);
        Assert.Equal(2, entry.RowCount);
        Assert.True(report.SourceScope.NonTableTextBlockCount == 0);
        using ExcelDocumentReader reader = ExcelDocumentReader.Open(workbook.ToArray());
        object?[,] values = reader.GetSheet(entry.SheetName).ReadRange(entry.Range);
        Assert.Equal("Due Date", values[0, 2]);
        Assert.Equal(new DateTime(2026, 7, 31), Convert.ToDateTime(values[2, 2], CultureInfo.InvariantCulture));
    }

    [Fact]
    public void PdfTables_PositionedCellRecoveryProcessesBandsNotCoveredByOtherTables() {
        static PdfCore.TextLayoutEngine.TextLine Row(double y, string left, double leftAdvance, string right) {
            var spans = new List<PdfCore.PdfTextSpan> {
                new(left, "F1", 10, 20, y, leftAdvance),
                new(right, "F1", 10, 220, y, 30)
            };
            return new PdfCore.TextLayoutEngine.TextLine(y, 20, 250, left + " " + right, spans);
        }

        var bands = new List<List<PdfCore.TextLayoutEngine.TextLine>> {
            new() { Row(700, "Code", 80, "Total") },
            new() { Row(680, "A-100", 80, "12") },
            new() { Row(660, "B-200", 80, "14") },
            new() { Row(300, "Name", 30, "Qty") },
            new() { Row(280, "Alpha", 90, "2") },
            new() { Row(260, "Beta", 150, "14") }
        };

        List<PdfCore.StructuredTable> tables = PdfCore.TableDetector.DetectTablesFromBands(bands);

        Assert.Contains(tables, table => table.Kind == "band-group" && table.Rows[0][0] == "Code");
        PdfCore.StructuredTable recovered = Assert.Single(tables, table => table.Kind == "positioned-cells-bounded");
        Assert.Equal(new[] { "Name", "Qty" }, recovered.Rows[0]);
        Assert.Equal(new[] { "Beta", "14" }, recovered.Rows[2]);
    }

    [Fact]
    public void PdfTables_PositionedCellRecoverySplitsAlignedTablesAcrossLargeVerticalGaps() {
        static PdfCore.TextLayoutEngine.TextLine Row(double y, string left, string right) {
            var spans = new List<PdfCore.PdfTextSpan> {
                new(left, "F1", 10, 20, y, 60),
                new(right, "F1", 10, 180, y, 30)
            };
            return new PdfCore.TextLayoutEngine.TextLine(y, 20, 210, left + " " + right, spans);
        }

        var lines = new[] {
            Row(700, "Code", "Qty"),
            Row(680, "A-100", "12"),
            Row(660, "B-200", "14"),
            Row(500, "Name", "Qty"),
            Row(480, "Alpha", "12"),
            Row(460, "Beta", "14")
        };

        List<PdfCore.StructuredTable> tables = PdfCore.TableDetector.DetectPositionedCellTables(lines);

        Assert.Equal(2, tables.Count);
        Assert.Equal(new[] { "Code", "Qty" }, tables[0].Rows[0]);
        Assert.Equal(new[] { "Name", "Qty" }, tables[1].Rows[0]);
    }

    [Fact]
    public void PdfTables_PositionedCellRecoveryKeepsHorizontallySeparateOverlappingLines() {
        static PdfCore.TextLayoutEngine.TextLine Row(
            double y,
            double x,
            string left,
            double leftAdvance,
            string right) {
            var spans = new List<PdfCore.PdfTextSpan> {
                new(left, "F1", 10, x, y, leftAdvance),
                new(right, "F1", 10, x + 200, y, 30)
            };
            return new PdfCore.TextLayoutEngine.TextLine(y, x, x + 230, left + " " + right, spans);
        }

        var bands = new List<List<PdfCore.TextLayoutEngine.TextLine>> {
            new() { Row(700, 20, "Code", 80, "Total") },
            new() { Row(680, 20, "A-100", 80, "12") },
            new() { Row(660, 20, "B-200", 80, "14") },
            new() { Row(700, 350, "Name", 30, "Qty") },
            new() { Row(680, 350, "Alpha", 90, "2") },
            new() { Row(660, 350, "Beta", 150, "14") }
        };

        List<PdfCore.StructuredTable> tables = PdfCore.TableDetector.DetectTablesFromBands(bands);

        Assert.Contains(tables, table => table.Kind == "band-group" && table.Rows[0][0] == "Code");
        PdfCore.StructuredTable recovered = Assert.Single(
            tables,
            table => table.Kind == "positioned-cells-bounded");
        Assert.Equal(new[] { "Name", "Qty" }, recovered.Rows[0]);
        Assert.Equal(new[] { "Beta", "14" }, recovered.Rows[2]);
    }

    [Fact]
    public void PdfTables_PositionedContinuationsMatchStableColumnStartsInsteadOfTextWidths() {
        static PdfCore.PdfLogicalTable Table(double firstWidth, double secondWidth) {
            var table = new PdfCore.StructuredTable {
                Kind = "positioned-cells-bounded",
                YTop = 700,
                YBottom = 660
            };
            table.Columns.Add(new PdfCore.StructuredTableColumn { From = 20, To = 20 + firstWidth });
            table.Columns.Add(new PdfCore.StructuredTableColumn { From = 220, To = 220 + secondWidth });
            table.Rows.Add(new[] { "Code", "Total" });
            table.Rows.Add(new[] { "A-100", "12" });
            return PdfCore.PdfLogicalTable.From(1, table);
        }

        Assert.True(PdfCore.PdfLogicalTableContinuations.HasCompatibleColumns(
            Table(40, 30),
            Table(160, 90),
            tolerance: 4D));
    }

    [Fact]
    public void PdfTables_NumericParserHandlesInvoiceNumberText() {
        Assert.True(PdfCore.PdfLogicalTableAnalysis.TryParseNumericValue("$1,234.50", CultureInfo.InvariantCulture, out decimal currency));
        Assert.Equal(1234.50m, currency);

        Assert.True(PdfCore.PdfLogicalTableAnalysis.TryParseNumericValue("(99.95)", CultureInfo.InvariantCulture, out decimal parenthesizedNegative));
        Assert.Equal(-99.95m, parenthesizedNegative);

        Assert.True(PdfCore.PdfLogicalTableAnalysis.TryParseNumericValue("1 234,50", CultureInfo.GetCultureInfo("pl-PL"), out decimal localized));
        Assert.Equal(1234.50m, localized);

        Assert.False(PdfCore.PdfLogicalTableAnalysis.TryParseNumericValue("12%", CultureInfo.InvariantCulture, out _));
    }

    [Fact]
    public void PdfTables_SaveTablesAsExcel_AppliesRowCapsAndKeepsWorkbookValidWhenEmpty() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .KeyValueTable(new[] {
                PdfCore.PdfKeyValueRow.Text("InvoiceId", "INV-001"),
                PdfCore.PdfKeyValueRow.Text("Customer", "Evotec"),
                PdfCore.PdfKeyValueRow.Text("Due", "2026-06-30")
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 120, 170 },
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .PageBreak()
            .Paragraph(p => p.Text("No table on this page."))
            .ToBytes();

        using var workbook = new MemoryStream();
        PdfExcelTableImportReport report = PdfExcelTableConverterExtensions.SaveTablesAsExcel(
            LoadTables(pdf, PdfCore.PdfPageRange.From(1, 1)),
            workbook,
            new PdfExcelTableImportOptions {
                MaxRows = 2,
                AutoFitColumns = false
            });

        PdfExcelTableImportEntry result = Assert.Single(report.Entries);
        Assert.Equal(1, result.PageNumber);
        Assert.Equal(2, result.RowCount);
        Assert.Equal(3, result.TotalRowCount);
        Assert.True(result.Truncated);
        Assert.True(report.HasLoss);
        Assert.Throws<InvalidOperationException>(() => report.RequireNoLoss());

        using ExcelDocumentReader reader = ExcelDocumentReader.Open(workbook.ToArray());
        object?[,] values = reader.GetSheet(result.SheetName).ReadRange(result.Range);
        Assert.Equal("Key", values[0, 0]);
        Assert.Equal("InvoiceId", values[1, 0]);
        Assert.Equal("Customer", values[2, 0]);

        using var emptyWorkbook = new MemoryStream();
        PdfExcelTableImportReport emptyReport = PdfExcelTableConverterExtensions.SaveTablesAsExcel(
            LoadTables(pdf, PdfCore.PdfPageRange.From(2, 2)),
            emptyWorkbook,
            new PdfExcelTableImportOptions {
                AutoFitColumns = false
            });

        Assert.Empty(emptyReport.Entries);
        using ExcelDocumentReader emptyReader = ExcelDocumentReader.Open(emptyWorkbook.ToArray());
        object?[,] emptyValues = emptyReader.GetSheet("PDF Tables").ReadRange("A1:A1");
        Assert.Equal("No PDF tables detected.", emptyValues[0, 0]);
    }

    private static PdfCore.PdfLogicalDocument LoadTables(byte[] pdf, params PdfCore.PdfPageRange[] ranges) {
        var layout = new PdfCore.PdfTextLayoutOptions { ForceSingleColumn = true };
        return ranges.Length == 0
            ? PdfCore.PdfLogicalDocument.Load(pdf, layout)
            : PdfCore.PdfLogicalDocument.LoadPageRanges(pdf, layout, ranges);
    }

    private static Cell GetCell(SheetData sheetData, string reference) {
        return sheetData.Descendants<Cell>()
            .Single(cell => string.Equals(cell.CellReference?.Value, reference, StringComparison.OrdinalIgnoreCase));
    }

    private static SheetData GetOnlySheetData(SpreadsheetDocument spreadsheet) {
        WorkbookPart workbookPart = spreadsheet.WorkbookPart ?? throw new InvalidOperationException("Workbook part is missing.");
        WorksheetPart worksheetPart = Assert.Single(workbookPart.WorksheetParts);
        Worksheet worksheet = worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet is missing.");
        return worksheet.GetFirstChild<SheetData>() ?? throw new InvalidOperationException("SheetData is missing.");
    }

    private static bool IsTextCell(Cell cell) {
        CellValues? dataType = cell.DataType?.Value;
        return dataType == CellValues.SharedString || dataType == CellValues.String || dataType == CellValues.InlineString;
    }
}
