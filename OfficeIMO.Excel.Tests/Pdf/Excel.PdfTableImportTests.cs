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
    public void PdfTables_SaveTablesAsExcel_PreservesHeaderWhenNarrativePrecedesAutoSizedTable() {
        byte[] pdf = PdfCore.PdfDocument.Create()
            .H1("Quarterly results")
            .Paragraph(paragraph => paragraph.Text("Revenue improved in the current quarter."))
            .Table(new[] {
                new[] { "Region", "Revenue", "Active" },
                new[] { "North", "1250", "True" },
                new[] { "South", "980", "False" },
                new[] { "West", "1430", "True" }
            })
            .ToBytes();

        PdfCore.PdfDocumentReadResult logical = PdfCore.PdfDocument.Load(pdf).Read();
        PdfCore.PdfLogicalTable table = Assert.Single(logical.Pages.SelectMany(static page => page.Tables));
        Assert.Equal(4, table.Rows.Count);

        using var workbook = new MemoryStream();
        PdfExcelTableImportEntry result = Assert.Single(logical.SaveTablesAsExcel(
            workbook,
            new PdfExcelTableImportOptions {
                AutoFitColumns = false
            }).Entries);

        Assert.Equal(3, result.RowCount);
        using ExcelDocumentReader reader = ExcelDocumentReader.Open(workbook.ToArray());
        ExcelTableInfo excelTable = Assert.Single(reader.GetTables());
        Assert.Equal(new[] { "Region", "Revenue", "Active" }, excelTable.Columns.Select(static column => column.Name).ToArray());
        object?[,] values = reader.GetSheet(result.SheetName).ReadRange(result.Range);
        Assert.Equal("North", values[1, 0]);
        Assert.Equal(1250d, Convert.ToDouble(values[1, 1], CultureInfo.InvariantCulture));
        Assert.Equal("West", values[3, 0]);
    }

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

        PdfCore.PdfDocumentReadResult logical = LoadTables(pdf);
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
    public void PdfTables_ContinuationGroupingUsesTheVisibleCropOrigin() {
        byte[] pdf = BuildContinuationTablePdf(120D, 120D);
        byte[] cropped = PdfCore.PdfDocument.Load(pdf)
            .Pages.SetCropBox(10, 10, 310, 210)
            .ToBytes();
        PdfCore.PdfDocumentReadResult logical = LoadTables(cropped);

        PdfCore.PdfLogicalTableContinuationGroup group = Assert.Single(
            PdfCore.PdfLogicalTableContinuations.Group(logical, 0, true, true, 64, 4D));

        Assert.True(group.Segments.Count > 1);
        Assert.Equal(30, group.TotalRowCount);
    }

    [Fact]
    public void PdfTables_ContinuationGroupingDoesNotMergeSidewaysRotatedTables() {
        byte[] pdf = BuildContinuationTablePdf(45D, 45D);
        byte[] rotated = PdfCore.PdfDocument.Load(pdf)
            .Pages.Rotate(90)
            .ToBytes();
        PdfCore.PdfDocumentReadResult logical = LoadTables(rotated);
        Assert.True(logical.Pages.Count > 1);

        IReadOnlyList<PdfCore.PdfLogicalTableContinuationGroup> groups =
            PdfCore.PdfLogicalTableContinuations.Group(logical, 0, true, true, 64, 4D);

        Assert.NotEmpty(groups);
        Assert.All(groups, static group => Assert.Single(group.Segments));
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

        PdfCore.PdfDocumentReadResult logical = LoadTables(pdf);
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

        PdfCore.PdfDocumentReadResult logical = LoadTables(pdf);
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
    [InlineData("March 5", "April 7")]
    [InlineData("01/02", "03/04")]
    public void PdfTables_SaveTablesAsExcel_KeepsYearlessDateValuesAsText(string firstValue, string secondValue) {
        byte[] pdf = PdfCore.PdfDocument.Create()
            .Table(new[] {
                new[] { "Date", "Label" },
                new[] { firstValue, "First" },
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
        Assert.Equal(firstValue, values[1, 0]);
        Assert.Equal(secondValue, values[2, 0]);
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

        PdfCore.PdfDocumentReadResult logical = LoadTables(pdf);
        using var workbook = new MemoryStream();
        PdfExcelTableImportReport report = logical.SaveTablesAsExcel(
            workbook,
            new PdfExcelTableImportOptions { AutoFitColumns = false });

        IReadOnlyList<PdfCore.PdfTextSpan> spans = logical.Pages[0].TextBlocks.SelectMany(static block => block.Spans).ToArray();
        List<PdfCore.TextLayoutEngine.TextLine> lines = PdfCore.TextLayoutEngine.BuildLines(spans);
        List<PdfCore.StructuredTable> recovered = PdfCore.TableDetector.DetectPositionedCellTables(
            lines,
            logical.Pages[0].Height);
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
        Assert.DoesNotContain(tables, table =>
            table.Kind == "band-group" &&
            table.Rows.SelectMany(static row => row).Contains("Name", StringComparer.Ordinal));
    }

    [Theory]
    [InlineData("Name", "Total")]
    [InlineData("Metric", "2025")]
    [InlineData("Metric", "FY2025")]
    [InlineData("Metric", "Q1")]
    [InlineData("Employee", "Salary")]
    public void PdfTables_BandGroupingStopsAtANewEmphasizedHeader(string headerLeft, string headerRight) {
        static PdfCore.TextLayoutEngine.TextLine Row(double y, bool emphasized, string left, string right) {
            string? baseFont = emphasized ? "Helvetica-Bold" : "Helvetica";
            var spans = new List<PdfCore.PdfTextSpan> {
                new(left, "F1", 10, 20, y, 60, baseFont: baseFont),
                new(right, "F1", 10, 180, y, 40, baseFont: baseFont)
            };
            return new PdfCore.TextLayoutEngine.TextLine(y, 20, 220, left + " " + right, spans);
        }

        var bands = new List<List<PdfCore.TextLayoutEngine.TextLine>> {
            new() { Row(700, true, "Code", "Qty") },
            new() { Row(680, false, "A-100", "12") },
            new() { Row(660, false, "B-200", "14") },
            new() { Row(640, true, headerLeft, headerRight) },
            new() { Row(620, false, "Alpha", "22") },
            new() { Row(600, false, "Beta", "24") }
        };

        List<PdfCore.StructuredTable> tables = PdfCore.TableDetector.DetectTablesFromBands(bands);
        PdfCore.StructuredTable[] grouped = tables.Where(static table => table.Kind == "band-group").ToArray();

        Assert.Equal(2, grouped.Length);
        Assert.Equal(new[] { "Code", "Qty" }, grouped[0].Rows[0]);
        Assert.Equal(new[] { headerLeft, headerRight }, grouped[1].Rows[0]);
    }

    [Fact]
    public void PdfTables_BandGroupingStopsAtASplitlessInterveningHeader() {
        static PdfCore.TextLayoutEngine.TextLine WideRow(double y, bool emphasized, string left, string right) {
            string? baseFont = emphasized ? "Helvetica-Bold" : "Helvetica";
            var spans = new List<PdfCore.PdfTextSpan> {
                new(left, "F1", 10, 20, y, 60, baseFont: baseFont),
                new(right, "F1", 10, 400, y, 40, baseFont: baseFont)
            };
            return new PdfCore.TextLayoutEngine.TextLine(y, 20, 440, left + " " + right, spans);
        }

        static PdfCore.TextLayoutEngine.TextLine NarrowRow(double y, string left, string right) {
            var spans = new List<PdfCore.PdfTextSpan> {
                new(left, "F1", 10, 20, y, 60),
                new(right, "F1", 10, 180, y, 40)
            };
            return new PdfCore.TextLayoutEngine.TextLine(y, 20, 220, left + " " + right, spans);
        }

        static PdfCore.TextLayoutEngine.TextLine SplitlessHeader(double y) {
            var spans = new List<PdfCore.PdfTextSpan> {
                new("Name", "F1", 10, 20, y, 150, baseFont: "Helvetica-Bold"),
                new("Total", "F1", 10, 180, y, 100, baseFont: "Helvetica-Bold")
            };
            return new PdfCore.TextLayoutEngine.TextLine(y, 20, 280, "Name Total", spans);
        }

        var bands = new List<List<PdfCore.TextLayoutEngine.TextLine>> {
            new() { WideRow(700, true, "Code", "Qty") },
            new() { WideRow(680, false, "A-100", "12") },
            new() { WideRow(660, false, "B-200", "14") },
            new() { SplitlessHeader(640) },
            new() { NarrowRow(620, "Alpha", "22") },
            new() { NarrowRow(600, "Beta", "24") }
        };

        PdfCore.StructuredTable[] grouped = PdfCore.TableDetector.DetectTablesFromBands(bands)
            .Where(static table => table.Kind == "band-group")
            .ToArray();

        Assert.Equal(2, grouped.Length);
        Assert.Equal(new[] { "Code", "Qty" }, grouped[0].Rows[0]);
        Assert.Equal(new[] { "Name", "Total" }, grouped[1].Rows[0]);
    }

    [Theory]
    [InlineData("Total", "0")]
    [InlineData("Total", "7")]
    [InlineData("Total", "N/A")]
    [InlineData("Total", "—")]
    [InlineData("Grand Total", "TBD")]
    [InlineData("Total", "26")]
    [InlineData("Grand Total", "26")]
    [InlineData("Sub Total", "26")]
    [InlineData("Net Subtotal", "26")]
    [InlineData("Overall Totals", "26")]
    public void PdfTables_BandGroupingKeepsEmphasizedSummaryRows(string summaryLabel, string summaryValue) {
        static PdfCore.TextLayoutEngine.TextLine Row(double y, bool emphasized, string left, string right) {
            string? baseFont = emphasized ? "Helvetica-Bold" : "Helvetica";
            var spans = new List<PdfCore.PdfTextSpan> {
                new(left, "F1", 10, 20, y, 60, baseFont: baseFont),
                new(right, "F1", 10, 180, y, 40, baseFont: baseFont)
            };
            return new PdfCore.TextLayoutEngine.TextLine(y, 20, 220, left + " " + right, spans);
        }

        var bands = new List<List<PdfCore.TextLayoutEngine.TextLine>> {
            new() { Row(700, true, "Code", "Qty") },
            new() { Row(680, false, "A-100", "12") },
            new() { Row(660, false, "B-200", "14") },
            new() { Row(640, true, summaryLabel, summaryValue) }
        };

        PdfCore.StructuredTable table = Assert.Single(
            PdfCore.TableDetector.DetectTablesFromBands(bands),
            static candidate => candidate.Kind == "band-group");

        Assert.Equal(4, table.Rows.Count);
        Assert.Equal(new[] { summaryLabel, summaryValue }, table.Rows[3]);
    }

    [Theory]
    [InlineData("North", "1250")]
    [InlineData("North", "7")]
    [InlineData("North", "2025")]
    [InlineData("North", "FY2025")]
    [InlineData("North", "Q1")]
    [InlineData("North Region", "Q1")]
    [InlineData("C-300", "Closed")]
    [InlineData("Enabled", "Yes")]
    [InlineData("Central", "N/A")]
    public void PdfTables_BandGroupingKeepsEmphasizedDataRows(string label, string value) {
        static PdfCore.TextLayoutEngine.TextLine Row(double y, bool emphasized, string left, string right) {
            string? baseFont = emphasized ? "Helvetica-Bold" : "Helvetica";
            var spans = new List<PdfCore.PdfTextSpan> {
                new(left, "F1", 10, 20, y, 60, baseFont: baseFont),
                new(right, "F1", 10, 180, y, 40, baseFont: baseFont)
            };
            return new PdfCore.TextLayoutEngine.TextLine(y, 20, 220, left + " " + right, spans);
        }

        var bands = new List<List<PdfCore.TextLayoutEngine.TextLine>> {
            new() { Row(700, true, "Code", "Qty") },
            new() { Row(680, false, "A-100", "12") },
            new() { Row(660, false, "B-200", "14") },
            new() { Row(640, true, label, value) },
            new() { Row(620, false, "South", "980") }
        };

        PdfCore.StructuredTable table = Assert.Single(
            PdfCore.TableDetector.DetectTablesFromBands(bands),
            static candidate => candidate.Kind == "band-group");

        Assert.Equal(5, table.Rows.Count);
        Assert.Equal(new[] { label, value }, table.Rows[3]);
        Assert.Equal(new[] { "South", "980" }, table.Rows[4]);
    }

    [Fact]
    public void PdfTables_BandGroupingUsesStableLargeRowRhythmWhileExtending() {
        static PdfCore.TextLayoutEngine.TextLine Row(
            double y,
            bool emphasized,
            string left,
            double leftAdvance,
            string right) {
            string? baseFont = emphasized ? "Helvetica-Bold" : "Helvetica";
            var spans = new List<PdfCore.PdfTextSpan> {
                new(left, "F1", 10, 20, y, leftAdvance, baseFont: baseFont),
                new(right, "F1", 10, 200, y, 40, baseFont: baseFont)
            };
            return new PdfCore.TextLayoutEngine.TextLine(y, 20, 240, left + " " + right, spans);
        }

        var bands = new List<List<PdfCore.TextLayoutEngine.TextLine>> {
            new() { Row(700, true, "Region", 100, "Total") },
            new() { Row(640, false, "North", 110, "1250") },
            new() { Row(580, false, "South East", 130, "980") },
            new() { Row(520, false, "Western Europe", 145, "1430") }
        };

        PdfCore.StructuredTable table = Assert.Single(
            PdfCore.TableDetector.DetectTablesFromBands(bands),
            static candidate => candidate.Kind == "band-group");

        Assert.Equal(4, table.Rows.Count);
        Assert.Equal(new[] { "Western Europe", "1430" }, table.Rows[3]);
    }

    [Fact]
    public void PdfTables_BandGroupingKeepsNaturalSpanningSectionRowsAcrossSplitDrift() {
        static PdfCore.TextLayoutEngine.TextLine WideRow(double y, bool emphasized, string left, string right) {
            string? baseFont = emphasized ? "Helvetica-Bold" : "Helvetica";
            var spans = new List<PdfCore.PdfTextSpan> {
                new(left, "F1", 10, 20, y, 100, baseFont: baseFont),
                new(right, "F1", 10, 200, y, 40, baseFont: baseFont)
            };
            return new PdfCore.TextLayoutEngine.TextLine(y, 20, 240, left + " " + right, spans);
        }

        static PdfCore.TextLayoutEngine.TextLine NarrowRow(double y, string left, string right) {
            var spans = new List<PdfCore.PdfTextSpan> {
                new(left, "F1", 10, 20, y, 50),
                new(right, "F1", 10, 110, y, 130)
            };
            return new PdfCore.TextLayoutEngine.TextLine(y, 20, 240, left + " " + right, spans);
        }

        var sectionSpans = new List<PdfCore.PdfTextSpan> {
            new("North", "F1", 10, 20, 640, 60, baseFont: "Helvetica-Bold"),
            new("America", "F1", 10, 90, 640, 80, baseFont: "Helvetica-Bold")
        };
        var bands = new List<List<PdfCore.TextLayoutEngine.TextLine>> {
            new() { WideRow(700, true, "Region", "Total") },
            new() { WideRow(680, false, "Global", "2230") },
            new() { WideRow(660, false, "Europe", "980") },
            new() { new PdfCore.TextLayoutEngine.TextLine(640, 20, 180, "North America", sectionSpans) },
            new() { NarrowRow(620, "United States", "900") },
            new() { NarrowRow(600, "Canada", "350") }
        };

        PdfCore.StructuredTable table = Assert.Single(
            PdfCore.TableDetector.DetectTablesFromBands(bands),
            static candidate => candidate.Kind == "band-group");

        Assert.Equal(6, table.Rows.Count);
        Assert.Equal("North America", table.Rows[3][0]);
        Assert.Equal(string.Empty, table.Rows[3][1]);
        Assert.Equal(new[] { "Canada", "350" }, table.Rows[5]);
    }

    [Fact]
    public void PdfTables_BandGroupingUsesPhysicalRhythmAcrossSpanningSection() {
        static PdfCore.TextLayoutEngine.TextLine Row(
            double y,
            bool emphasized,
            string left,
            double leftAdvance,
            string right) {
            string? baseFont = emphasized ? "Helvetica-Bold" : "Helvetica";
            var spans = new List<PdfCore.PdfTextSpan> {
                new(left, "F1", 10, 20, y, leftAdvance, baseFont: baseFont),
                new(right, "F1", 10, 200, y, 40, baseFont: baseFont)
            };
            return new PdfCore.TextLayoutEngine.TextLine(y, 20, 240, left + " " + right, spans);
        }

        var sectionSpans = new List<PdfCore.PdfTextSpan> {
            new("North America", "F1", 10, 20, 520, 170, baseFont: "Helvetica-Bold")
        };
        var bands = new List<List<PdfCore.TextLayoutEngine.TextLine>> {
            new() { Row(700, true, "Region", 100, "Total") },
            new() { Row(640, false, "Global", 100, "2230") },
            new() { Row(580, false, "Europe", 100, "980") },
            new() { new PdfCore.TextLayoutEngine.TextLine(520, 20, 190, "North America", sectionSpans) },
            new() { Row(460, false, "United States", 155, "900") },
            new() { Row(400, false, "Canada", 150, "350") }
        };

        PdfCore.StructuredTable table = Assert.Single(
            PdfCore.TableDetector.DetectTablesFromBands(bands),
            static candidate => candidate.Kind == "band-group");

        Assert.Equal(6, table.Rows.Count);
        Assert.Equal("North America", table.Rows[3][0]);
        Assert.Equal(string.Empty, table.Rows[3][1]);
        Assert.Equal(new[] { "Canada", "350" }, table.Rows[5]);
    }

    [Fact]
    public void PdfTables_BandGroupingStopsAtNewHeaderAfterAttachedTwoRowTable() {
        static PdfCore.TextLayoutEngine.TextLine Row(double y, bool emphasized, string left, string right) {
            string? baseFont = emphasized ? "Helvetica-Bold" : "Helvetica";
            var spans = new List<PdfCore.PdfTextSpan> {
                new(left, "F1", 10, 20, y, 60, baseFont: baseFont),
                new(right, "F1", 10, 180, y, 40, baseFont: baseFont)
            };
            return new PdfCore.TextLayoutEngine.TextLine(y, 20, 220, left + " " + right, spans);
        }

        var attachedHeaderSpans = new List<PdfCore.PdfTextSpan> {
            new("Code", "F1", 10, 20, 700, 150, baseFont: "Helvetica-Bold"),
            new("Qty", "F1", 10, 180, 700, 40, baseFont: "Helvetica-Bold")
        };
        var bands = new List<List<PdfCore.TextLayoutEngine.TextLine>> {
            new() { new PdfCore.TextLayoutEngine.TextLine(700, 20, 220, "Code Qty", attachedHeaderSpans) },
            new() { Row(680, false, "A-100", "12") },
            new() { Row(660, true, "Name", "Total") },
            new() { Row(640, false, "Alpha", "22") },
            new() { Row(620, false, "Beta", "24") }
        };

        PdfCore.StructuredTable[] grouped = PdfCore.TableDetector.DetectTablesFromBands(bands)
            .Where(static table => table.Kind == "band-group")
            .ToArray();

        Assert.Equal(2, grouped.Length);
        Assert.Equal(2, grouped[0].Rows.Count);
        Assert.Equal(new[] { "Code", "Qty" }, grouped[0].Rows[0]);
        Assert.Equal(new[] { "Name", "Total" }, grouped[1].Rows[0]);
    }

    [Fact]
    public void PdfTables_BandGroupingDoesNotBridgeDistantCaptionByProportionAlone() {
        static PdfCore.TextLayoutEngine.TextLine Row(double y, bool emphasized, string left, string right) {
            string? baseFont = emphasized ? "Helvetica-Bold" : "Helvetica";
            var spans = new List<PdfCore.PdfTextSpan> {
                new(left, "F1", 10, 20, y, 60, baseFont: baseFont),
                new(right, "F1", 10, 180, y, 40, baseFont: baseFont)
            };
            return new PdfCore.TextLayoutEngine.TextLine(y, 20, 220, left + " " + right, spans);
        }

        var captionSpans = new List<PdfCore.PdfTextSpan> {
            new("INTERIM RESULTS", "F1", 10, 20, 400, 200, baseFont: "Helvetica-Bold")
        };
        var bands = new List<List<PdfCore.TextLayoutEngine.TextLine>> {
            new() { Row(700, true, "Code", "Qty") },
            new() { Row(680, false, "A-100", "12") },
            new() { Row(660, false, "B-200", "14") },
            new() { new PdfCore.TextLayoutEngine.TextLine(400, 20, 220, "INTERIM RESULTS", captionSpans) },
            new() { Row(140, false, "Alpha", "22") },
            new() { Row(120, false, "Beta", "24") }
        };

        PdfCore.StructuredTable first = Assert.Single(
            PdfCore.TableDetector.DetectTablesFromBands(bands),
            static table => table.Kind == "band-group" && table.Rows[0][0] == "Code");

        Assert.Equal(3, first.Rows.Count);
        Assert.DoesNotContain(first.Rows, static row => row.Contains("INTERIM RESULTS", StringComparer.Ordinal));
        Assert.DoesNotContain(first.Rows, static row => row.Contains("Alpha", StringComparer.Ordinal));
    }

    [Fact]
    public void PdfTables_BandGroupingDoesNotUseDistantAttachedHeaderAsRowRhythm() {
        static PdfCore.TextLayoutEngine.TextLine Row(double y, string left, string right) {
            var spans = new List<PdfCore.PdfTextSpan> {
                new(left, "F1", 10, 20, y, 60),
                new(right, "F1", 10, 180, y, 40)
            };
            return new PdfCore.TextLayoutEngine.TextLine(y, 20, 220, left + " " + right, spans);
        }

        var headerSpans = new List<PdfCore.PdfTextSpan> {
            new("Account", "F1", 10, 20, 900, 150, baseFont: "Helvetica-Bold"),
            new("Total", "F1", 10, 180, 900, 40, baseFont: "Helvetica-Bold")
        };
        var captionSpans = new List<PdfCore.PdfTextSpan> {
            new("INTERIM RESULTS", "F1", 10, 20, 450, 200, baseFont: "Helvetica-Bold")
        };
        var bands = new List<List<PdfCore.TextLayoutEngine.TextLine>> {
            new() { new PdfCore.TextLayoutEngine.TextLine(900, 20, 220, "Account Total", headerSpans) },
            new() { Row(700, "North", "1250") },
            new() { new PdfCore.TextLayoutEngine.TextLine(450, 20, 220, "INTERIM RESULTS", captionSpans) },
            new() { Row(200, "Alpha", "22") },
            new() { Row(180, "Beta", "24") }
        };

        PdfCore.StructuredTable first = Assert.Single(
            PdfCore.TableDetector.DetectTablesFromBands(bands),
            static table => table.Kind == "band-group" && table.Rows[0][0] == "Account");

        Assert.Equal(2, first.Rows.Count);
        Assert.DoesNotContain(first.Rows, static row => row.Contains("INTERIM RESULTS", StringComparer.Ordinal));
        Assert.DoesNotContain(first.Rows, static row => row.Contains("Alpha", StringComparer.Ordinal));
    }

    [Fact]
    public void PdfTables_BandGroupingUsesRowRhythmForLargeHeaderGaps() {
        static PdfCore.TextLayoutEngine.TextLine Header(double y) {
            var spans = new List<PdfCore.PdfTextSpan> {
                new("Account Name", "F1", 10, 20, y, 150, baseFont: "Helvetica-Bold"),
                new("Annual Total", "F1", 10, 180, y, 70, baseFont: "Helvetica-Bold")
            };
            return new PdfCore.TextLayoutEngine.TextLine(y, 20, 250, "Account Name Annual Total", spans);
        }

        static PdfCore.TextLayoutEngine.TextLine Body(double y, string left, string right) {
            var spans = new List<PdfCore.PdfTextSpan> {
                new(left, "F1", 10, 20, y, 60),
                new(right, "F1", 10, 180, y, 40)
            };
            return new PdfCore.TextLayoutEngine.TextLine(y, 20, 220, left + " " + right, spans);
        }

        var bands = new List<List<PdfCore.TextLayoutEngine.TextLine>> {
            new() { Header(700) },
            new() { Body(640, "North", "1250") },
            new() { Body(580, "South", "980") }
        };

        PdfCore.StructuredTable table = Assert.Single(
            PdfCore.TableDetector.DetectTablesFromBands(bands),
            static candidate => candidate.Kind == "band-group");

        Assert.Equal(3, table.Rows.Count);
        Assert.Equal(new[] { "Account Name", "Annual Total" }, table.Rows[0]);
        Assert.Equal(new[] { "South", "980" }, table.Rows[2]);
    }

    [Fact]
    public void PdfTables_BandGroupingAllowsStrongTwoRowTablesAcrossLargeGaps() {
        var headerSpans = new List<PdfCore.PdfTextSpan> {
            new("Account", "F1", 10, 20, 700, 145, baseFont: "Helvetica-Bold"),
            new("Total", "F1", 10, 180, 700, 70, baseFont: "Helvetica-Bold")
        };
        var bodySpans = new List<PdfCore.PdfTextSpan> {
            new("North", "F1", 10, 20, 600, 60),
            new("1250", "F1", 10, 180, 600, 40)
        };
        var bands = new List<List<PdfCore.TextLayoutEngine.TextLine>> {
            new() { new PdfCore.TextLayoutEngine.TextLine(700, 20, 250, "Account Total", headerSpans) },
            new() { new PdfCore.TextLayoutEngine.TextLine(600, 20, 220, "North 1250", bodySpans) }
        };

        PdfCore.StructuredTable table = Assert.Single(
            PdfCore.TableDetector.DetectTablesFromBands(bands),
            static candidate => candidate.Kind == "band-group");

        Assert.Equal(2, table.Rows.Count);
        Assert.Equal(new[] { "Account", "Total" }, table.Rows[0]);
        Assert.Equal(new[] { "North", "1250" }, table.Rows[1]);
    }

    [Fact]
    public void PdfTables_BandGroupingKeepsSplitlessHeaderAboveMultiLineBodyBand() {
        var headerSpans = new List<PdfCore.PdfTextSpan> {
            new("Account", "F1", 10, 20, 800, 145, baseFont: "Helvetica-Bold"),
            new("Total", "F1", 10, 180, 800, 70, baseFont: "Helvetica-Bold")
        };

        static PdfCore.TextLayoutEngine.TextLine BodyLine(double y, string left, string right) {
            var spans = new List<PdfCore.PdfTextSpan> {
                new(left, "F1", 10, 20, y, 60),
                new(right, "F1", 10, 180, y, 40)
            };
            return new PdfCore.TextLayoutEngine.TextLine(y, 20, 220, left + " " + right, spans);
        }

        var bands = new List<List<PdfCore.TextLayoutEngine.TextLine>> {
            new() { new PdfCore.TextLayoutEngine.TextLine(800, 20, 250, "Account Total", headerSpans) },
            new() { BodyLine(680, "North", "1250"), BodyLine(668, "Region", "100") },
            new() { BodyLine(640, "South", "980") }
        };

        PdfCore.StructuredTable table = Assert.Single(
            PdfCore.TableDetector.DetectTablesFromBands(bands),
            static candidate => candidate.Kind == "band-group");

        Assert.Equal(4, table.Rows.Count);
        Assert.Equal(new[] { "Account", "Total" }, table.Rows[0]);
        Assert.Equal(new[] { "North", "1250" }, table.Rows[1]);
        Assert.Equal(new[] { "South", "980" }, table.Rows[3]);
    }

    [Fact]
    public void PdfTables_BandGroupingKeepsBaseSplitForAttachedSplitlessHeader() {
        var headerSpans = new List<PdfCore.PdfTextSpan> {
            new("Account", "F1", 10, 20, 700, 140, baseFont: "Helvetica-Bold"),
            new("Total", "F1", 10, 165, 700, 80, baseFont: "Helvetica-Bold")
        };

        static PdfCore.TextLayoutEngine.TextLine Body(
            double y,
            string left,
            double leftAdvance,
            double rightX,
            double rightAdvance,
            string right) {
            var spans = new List<PdfCore.PdfTextSpan> {
                new(left, "F1", 10, 20, y, leftAdvance),
                new(right, "F1", 10, rightX, y, rightAdvance)
            };
            return new PdfCore.TextLayoutEngine.TextLine(y, 20, rightX + rightAdvance, left + " " + right, spans);
        }

        var bands = new List<List<PdfCore.TextLayoutEngine.TextLine>> {
            new() { new PdfCore.TextLayoutEngine.TextLine(700, 20, 245, "Account Total", headerSpans) },
            new() { Body(680, "North", 80, 220, 40, "1250") },
            new() { Body(660, "Western Europe", 155, 201, 59, "980") }
        };

        PdfCore.StructuredTable table = Assert.Single(
            PdfCore.TableDetector.DetectTablesFromBands(bands),
            static candidate => candidate.Kind == "band-group");

        Assert.Equal(3, table.Rows.Count);
        Assert.True(table.Columns[0].To > 165);
        Assert.Equal(new[] { "Account", "Total" }, table.Rows[0]);
        Assert.Equal(new[] { "Western Europe", "980" }, table.Rows[2]);
    }

    [Fact]
    public void PdfTables_BandGroupingStopsAtMultiLineEmphasizedHeaderBand() {
        static PdfCore.TextLayoutEngine.TextLine Row(double y, bool emphasized, string left, string right) {
            string? baseFont = emphasized ? "Helvetica-Bold" : "Helvetica";
            var spans = new List<PdfCore.PdfTextSpan> {
                new(left, "F1", 10, 20, y, 60, baseFont: baseFont),
                new(right, "F1", 10, 180, y, 40, baseFont: baseFont)
            };
            return new PdfCore.TextLayoutEngine.TextLine(y, 20, 220, left + " " + right, spans);
        }

        var bands = new List<List<PdfCore.TextLayoutEngine.TextLine>> {
            new() { Row(700, true, "Code", "Qty") },
            new() { Row(680, false, "A-100", "12") },
            new() { Row(660, false, "B-200", "14") },
            new() { Row(640, true, "Account", "Annual"), Row(632, true, "Name", "Total") },
            new() { Row(610, false, "North", "1250") },
            new() { Row(590, false, "South", "980") }
        };

        PdfCore.StructuredTable[] grouped = PdfCore.TableDetector.DetectTablesFromBands(bands)
            .Where(static table => table.Kind == "band-group")
            .ToArray();

        Assert.Equal(2, grouped.Length);
        Assert.Equal(new[] { "Code", "Qty" }, grouped[0].Rows[0]);
        Assert.Equal(new[] { "Account", "Annual" }, grouped[1].Rows[0]);
        Assert.Equal(new[] { "Name", "Total" }, grouped[1].Rows[1]);
    }

    [Fact]
    public void PdfTables_BandGroupingAdjustsSimilarSplitThatCrossesAColumnAnchor() {
        static PdfCore.TextLayoutEngine.TextLine Header(double y) {
            var spans = new List<PdfCore.PdfTextSpan> {
                new("Region", "F1", 10, 20, y, 60, baseFont: "Helvetica-Bold"),
                new("Total", "F1", 10, 120, y, 40, baseFont: "Helvetica-Bold")
            };
            return new PdfCore.TextLayoutEngine.TextLine(y, 20, 160, "Region Total", spans);
        }

        static PdfCore.TextLayoutEngine.TextLine Body(double y, string left, string right) {
            var spans = new List<PdfCore.PdfTextSpan> {
                new(left, "F1", 10, 20, y, 50),
                new(right, "F1", 10, 99, y, 61)
            };
            return new PdfCore.TextLayoutEngine.TextLine(y, 20, 160, left + " " + right, spans);
        }

        var bands = new List<List<PdfCore.TextLayoutEngine.TextLine>> {
            new() { Header(700) },
            new() { Body(680, "North", "7") },
            new() { Body(660, "South", "9") }
        };

        PdfCore.StructuredTable table = Assert.Single(
            PdfCore.TableDetector.DetectTablesFromBands(bands),
            static candidate => candidate.Kind == "band-group");

        Assert.Equal(new[] { "North", "7" }, table.Rows[1]);
        Assert.Equal(new[] { "South", "9" }, table.Rows[2]);
    }

    [Fact]
    public void PdfTables_BandGroupingRejectsDriftWithoutACommonSplit() {
        static PdfCore.TextLayoutEngine.TextLine Row(double y, double x, bool emphasized, string left, string right) {
            string? baseFont = emphasized ? "Helvetica-Bold" : "Helvetica";
            var spans = new List<PdfCore.PdfTextSpan> {
                new(left, "F1", 10, x, y, 115, baseFont: baseFont),
                new(right, "F1", 10, x + 140, y, 245 - (x + 140), baseFont: baseFont)
            };
            return new PdfCore.TextLayoutEngine.TextLine(y, x, 245, left + " " + right, spans);
        }

        static PdfCore.TextLayoutEngine.TextLine DriftedRow(double y) {
            var spans = new List<PdfCore.PdfTextSpan> {
                new("West", "F1", 10, 50, y, 105),
                new("region", "F1", 10, 160, y, 20),
                new("1430", "F1", 10, 205, y, 40)
            };
            return new PdfCore.TextLayoutEngine.TextLine(y, 50, 245, "West region 1430", spans);
        }

        var bands = new List<List<PdfCore.TextLayoutEngine.TextLine>> {
            new() { Row(700, 20, true, "Region", "Total") },
            new() { Row(680, 35, false, "North", "1250") },
            new() { DriftedRow(660) }
        };

        List<PdfCore.StructuredTable> grouped = PdfCore.TableDetector.DetectTablesFromBands(bands)
            .Where(static table => table.Kind == "band-group")
            .ToList();

        Assert.DoesNotContain(grouped, static table =>
            table.Rows.SelectMany(static row => row).Contains("West", StringComparer.Ordinal));
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
        Assert.DoesNotContain(tables, table =>
            table.Kind == "band-group" &&
            table.Rows.SelectMany(static row => row).Contains("Name", StringComparer.Ordinal));
    }

    [Fact]
    public void PdfTables_PositionedCellRecoveryPartitionsSideBySideTablesOnSharedBaselines() {
        static PdfCore.TextLayoutEngine.TextLine Row(
            double y,
            string leftName,
            string leftValue,
            string rightName,
            string rightValue) {
            var spans = new List<PdfCore.PdfTextSpan> {
                new(leftName, "F1", 10, 20, y, 50),
                new(leftValue, "F1", 10, 160, y, 30),
                new(rightName, "F1", 10, 460, y, 50),
                new(rightValue, "F1", 10, 600, y, 30)
            };
            return new PdfCore.TextLayoutEngine.TextLine(
                y,
                20,
                630,
                string.Join(" ", leftName, leftValue, rightName, rightValue),
                spans);
        }

        var lines = new[] {
            Row(700, "Code", "Qty", "Name", "Total"),
            Row(680, "A-100", "2", "Alpha", "12"),
            Row(660, "B-200", "14", "Beta", "24")
        };

        List<PdfCore.StructuredTable> tables = PdfCore.TableDetector.DetectPositionedCellTables(lines);

        Assert.Equal(2, tables.Count);
        Assert.Equal(new[] { "Code", "Qty" }, tables[0].Rows[0]);
        Assert.Equal(new[] { "Name", "Total" }, tables[1].Rows[0]);
        Assert.All(tables, table => Assert.Equal("positioned-cells-bounded", table.Kind));
    }

    [Fact]
    public void PdfTables_PositionedContinuationsMatchStableColumnStartsInsteadOfTextWidths() {
        PdfCore.PdfLogicalPage page = Assert.Single(PdfCore.PdfDocumentReadResult.Load(
            PdfCore.PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Page geometry")).ToBytes()).Pages);

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
            page,
            Table(160, 90),
            page,
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

    private static PdfCore.PdfDocumentReadResult LoadTables(byte[] pdf, params PdfCore.PdfPageRange[] ranges) {
        var layout = new PdfCore.PdfTextLayoutOptions { ForceSingleColumn = true };
        return ranges.Length == 0
            ? PdfCore.PdfDocumentReadResult.Load(pdf, layout)
            : PdfCore.PdfDocumentReadResult.LoadPageRanges(pdf, layout, ranges);
    }

    private static byte[] BuildContinuationTablePdf(double firstColumnWidth, double secondColumnWidth) {
        var rows = new List<string[]> { new[] { "Metric", "Owner" } };
        for (int index = 1; index <= 30; index++) {
            rows.Add(new[] {
                "C" + index.ToString("D2", CultureInfo.InvariantCulture),
                "T" + index.ToString("D2", CultureInfo.InvariantCulture)
            });
        }
        return PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
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
                ColumnWidthPoints = new List<double?> { firstColumnWidth, secondColumnWidth },
                CellPaddingX = 5,
                CellPaddingY = 3
            })
            .ToBytes();
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
