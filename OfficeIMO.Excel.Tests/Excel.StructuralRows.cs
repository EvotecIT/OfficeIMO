using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Threaded = DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_InsertRows_UpdatesWorkbookReferencesAndRowBoundMetadata() {
            string filePath = Path.Combine(_directoryWithFiles, "Excel.StructuralRows.Insert.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet data = document.AddWorksheet("Data");
                ExcelSheet summary = document.AddWorksheet("Summary");

                data.CellAt(1, 1).SetValue("Value");
                data.CellAt(1, 2).SetValue("Link");
                data.CellAt(2, 1).SetValue(10);
                data.CellAt(2, 2).SetValue("First");
                data.CellAt(3, 1).SetValue(20);
                data.CellAt(3, 2).SetValue("Second");
                data.CellFormula(
                    1,
                    5,
                    "SUM(A2:A3)+$A$3+SUM(2:3)+SUM(A3:A2)+SUM(3:2)+SUM(Data!A2:Data!A3)+SUM(Data!2:Data!3)");
                data.SetArrayFormula("C2:C3", "A2:A3*2");
                data.AddTable("A1:B3", hasHeader: true, name: "DataTable", OfficeIMO.Excel.TableStyle.TableStyleMedium2);
                data.SetNamedRange("LocalData", "A2:A3", save: false);
                summary.SetNamedRange("SummaryLocal", "A2:A3", save: false);
                document.SetNamedRange("GlobalData", "'Data'!A2:A3", save: false);
                document.SetPrintTitles(data, firstRow: 2, lastRow: 3, firstCol: null, lastCol: null, save: false);
                summary.CellFormula(1, 1, "'Data'!A3+A3");
                summary.SetInternalLink(2, 1, data, "A3", display: "Second value");
                data.Range("A2:A3").Validation.CustomFormula("A2>0");
                data.AddConditionalFormulaRule("B2:B3", "A2>0");
                data.SetComment(3, 1, "Second value");
                data.AddThreadedComment("A3", "Threaded second value");
                data.SetHyperlink(3, 2, "https://example.org", display: "Second");
                data.MergeRange("D2:D3");
                data.AutoFilterAdd("F1:G3");
                data.AddManualRowPageBreak(3);
                data.AddSparklines("A2:B2", "H2");

                data.InsertRows(2);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart dataPart = GetStructuralWorksheetPart(spreadsheet, "Data");
                WorksheetPart summaryPart = GetStructuralWorksheetPart(spreadsheet, "Summary");
                Cell[] dataCells = dataPart.Worksheet.Descendants<Cell>().ToArray();
                Cell arrayAnchor = dataCells.Single(cell => cell.CellReference?.Value == "C3");
                DataValidation validation = Assert.Single(dataPart.Worksheet.Descendants<DataValidation>());
                ConditionalFormattingRule conditionalRule = Assert.Single(dataPart.Worksheet.Descendants<ConditionalFormattingRule>());
                Table table = Assert.Single(dataPart.TableDefinitionParts).Table;
                Comment comment = Assert.Single(dataPart.WorksheetCommentsPart!.Comments!.CommentList!.Elements<Comment>());
                Threaded.ThreadedComment threadedComment = Assert.Single(
                    dataPart.WorksheetThreadedCommentsParts.Single().ThreadedComments!.Elements<Threaded.ThreadedComment>());
                Hyperlink hyperlink = Assert.Single(dataPart.Worksheet.Descendants<Hyperlink>());
                MergeCell merge = Assert.Single(dataPart.Worksheet.Descendants<MergeCell>());
                AutoFilter worksheetFilter = dataPart.Worksheet.GetFirstChild<AutoFilter>()!;
                Break pageBreak = Assert.Single(dataPart.Worksheet.GetFirstChild<RowBreaks>()!.Elements<Break>());
                DocumentFormat.OpenXml.Office2010.Excel.Sparkline sparkline = Assert.Single(
                    dataPart.Worksheet.Descendants<DocumentFormat.OpenXml.Office2010.Excel.Sparkline>());
                DefinedName[] names = spreadsheet.WorkbookPart!.Workbook.DefinedNames!.Elements<DefinedName>().ToArray();

                Assert.Contains(dataCells, cell => cell.CellReference?.Value == "A3" && cell.CellValue?.Text == "10");
                Assert.Contains(dataCells, cell => cell.CellReference?.Value == "A4" && cell.CellValue?.Text == "20");
                Assert.Equal(
                    "SUM(A3:A4)+$A$4+SUM(3:4)+SUM(A4:A3)+SUM(4:3)+SUM(Data!A3:Data!A4)+SUM(Data!3:Data!4)",
                    dataCells.Single(cell => cell.CellReference?.Value == "E1").CellFormula!.Text);
                Assert.Equal("A3:A4*2", arrayAnchor.CellFormula!.Text);
                Assert.Equal("C3:C4", arrayAnchor.CellFormula.Reference!.Value);
                Assert.Equal("'Data'!A4+A3", summaryPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A1").CellFormula!.Text);
                Assert.Equal("'Data'!A4", Assert.Single(summaryPart.Worksheet.Descendants<Hyperlink>()).Location!.Value);
                Assert.Equal("A1:B4", table.Reference!.Value);
                Assert.Equal("A1:B4", table.GetFirstChild<AutoFilter>()!.Reference!.Value);
                Assert.Equal("A3:A4", validation.SequenceOfReferences!.InnerText);
                Assert.Equal("A3>0", validation.Formula1!.Text);
                Assert.Equal("B3:B4", dataPart.Worksheet.Descendants<ConditionalFormatting>().Single().SequenceOfReferences!.InnerText);
                Assert.Equal("A3>0", Assert.Single(conditionalRule.Elements<Formula>()).Text);
                Assert.Equal("A4", comment.Reference!.Value);
                Assert.Equal("A4", threadedComment.Ref!.Value);
                Assert.Equal("B4", hyperlink.Reference!.Value);
                Assert.Equal("D3:D4", merge.Reference!.Value);
                Assert.Equal("F1:G4", worksheetFilter.Reference!.Value);
                Assert.Equal(4U, pageBreak.Id!.Value);
                Assert.Equal("H3", sparkline.ReferenceSequence!.Text);
                Assert.Equal("A3:B3", sparkline.Formula!.Text);
                Assert.Equal("'Data'!$A$3:$A$4", names.Single(name => name.Name?.Value == "GlobalData").Text);
                Assert.Equal("'Data'!$A$3:$A$4", names.Single(name => name.Name?.Value == "LocalData").Text);
                Assert.Equal("'Summary'!$A$2:$A$3", names.Single(name => name.Name?.Value == "SummaryLocal").Text);
                Assert.Equal("'Data'!$3:$4", names.Single(name => name.Name?.Value == "_xlnm.Print_Titles").Text);
                CalculationProperties calculation = spreadsheet.WorkbookPart.Workbook.GetFirstChild<CalculationProperties>()!;
                Assert.True(calculation.FullCalculationOnLoad!.Value);
                Assert.True(calculation.ForceFullCalculation!.Value);
                Assert.All(
                    spreadsheet.WorkbookPart.WorksheetParts.SelectMany(part => part.Worksheet.Descendants<CellFormula>()),
                    formula => Assert.True(formula.CalculateCell?.Value));
                Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
            }
        }

        [Fact]
        public void Test_DeleteRows_RewritesDeletedAndSurvivingFormulaReferences() {
            string filePath = Path.Combine(_directoryWithFiles, "Excel.StructuralRows.Delete.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet data = document.AddWorksheet("Data");
                ExcelSheet summary = document.AddWorksheet("Summary");
                for (int row = 1; row <= 5; row++) {
                    data.CellAt(row, 1).SetValue(row);
                }

                data.CellFormula(
                    1,
                    5,
                    "SUM(A2:A5)+A3+A5+SUM(2:5)+SUM(3:4)+SUM(A5:A2)+SUM(5:2)+SUM(Data!A2:Data!A5)+SUM(Data!2:Data!5)");
                data.SetArrayFormula("C2:C5", "A2:A5*2");
                summary.CellFormula(1, 1, "'Data'!A2:A5+'Data'!A3+A3");

                data.DeleteRows(3, 2);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart dataPart = GetStructuralWorksheetPart(spreadsheet, "Data");
                WorksheetPart summaryPart = GetStructuralWorksheetPart(spreadsheet, "Summary");
                Cell[] dataCells = dataPart.Worksheet.Descendants<Cell>().ToArray();
                Cell arrayAnchor = dataCells.Single(cell => cell.CellReference?.Value == "C2");

                Assert.Contains(dataCells, cell => cell.CellReference?.Value == "A3" && cell.CellValue?.Text == "5");
                Assert.Equal(
                    "SUM(A2:A3)+#REF!+A3+SUM(2:3)+SUM(#REF!)+SUM(A3:A2)+SUM(3:2)+SUM(Data!A2:Data!A3)+SUM(Data!2:Data!3)",
                    dataCells.Single(cell => cell.CellReference?.Value == "E1").CellFormula!.Text);
                Assert.Equal("A2:A3*2", arrayAnchor.CellFormula!.Text);
                Assert.Equal("C2:C3", arrayAnchor.CellFormula.Reference!.Value);
                Assert.Equal(
                    "'Data'!A2:A3+#REF!+A3",
                    summaryPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A1").CellFormula!.Text);
                Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
            }
        }

        [Fact]
        public void Test_InsertRows_UpdatesPivotSourceAndLocation() {
            string filePath = Path.Combine(_directoryWithFiles, "Excel.StructuralRows.Pivot.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet data = document.AddWorksheet("Data");
                data.CellAt(1, 1).SetValue("Region");
                data.CellAt(1, 2).SetValue("Sales");
                data.CellAt(2, 1).SetValue("East");
                data.CellAt(2, 2).SetValue(10);
                data.CellAt(3, 1).SetValue("West");
                data.CellAt(3, 2).SetValue(20);
                data.CellAt(4, 1).SetValue("East");
                data.CellAt(4, 2).SetValue(30);
                data.AddPivotTable(
                    sourceRange: "A1:B4",
                    destinationCell: "E2",
                    name: "SalesPivot",
                    rowFields: new[] { "Region" },
                    dataFields: new[] { new ExcelPivotDataField("Sales", DataConsolidateFunctionValues.Sum) });

                data.InsertRows(2);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                PivotTablePart pivotPart = spreadsheet.WorkbookPart!.WorksheetParts
                    .SelectMany(part => part.PivotTableParts)
                    .Single();
                WorksheetSource source = pivotPart.PivotTableCacheDefinitionPart!
                    .PivotCacheDefinition!
                    .CacheSource!
                    .WorksheetSource!;

                Assert.Equal("A1:B5", source.Reference!.Value);
                Assert.Equal("E3:F4", pivotPart.PivotTableDefinition!.Location!.Reference!.Value);
                Assert.True(pivotPart.PivotTableCacheDefinitionPart.PivotCacheDefinition.RefreshOnLoad!.Value);
                Assert.False(pivotPart.PivotTableCacheDefinitionPart.PivotCacheDefinition.SaveData!.Value);
                Assert.Equal(4U, pivotPart.PivotTableCacheDefinitionPart.PivotCacheDefinition.RecordCount!.Value);
                Assert.Equal(0U, pivotPart.PivotTableCacheDefinitionPart.PivotTableCacheRecordsPart!.PivotCacheRecords!.Count!.Value);
                Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
            }
        }

        [Fact]
        public void Test_InsertRows_RewritesChartReferencesAndInvalidatesChangedCaches() {
            string filePath = Path.Combine(_directoryWithFiles, "Excel.StructuralRows.Chart.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet data = document.AddWorksheet("Data");
                data.CellAt(1, 1).SetValue("Category");
                data.CellAt(1, 2).SetValue("Value");
                data.CellAt(2, 1).SetValue("First");
                data.CellAt(2, 2).SetValue(10);
                data.CellAt(3, 1).SetValue("Second");
                data.CellAt(3, 2).SetValue(20);
                data.AddChartFromRange(
                    "A1:B3",
                    row: 1,
                    column: 4,
                    type: ExcelChartType.ColumnClustered,
                    title: "Values");

                ChartPart chartPart = data.WorksheetPart.DrawingsPart!.ChartParts.Single();
                Assert.Contains(
                    chartPart.ChartSpace.Descendants<C.Formula>(),
                    formula => formula.Parent!.ChildElements.Any(element =>
                        element.LocalName.EndsWith("Cache", System.StringComparison.OrdinalIgnoreCase)));

                data.InsertRows(2);
                document.Save();
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            ChartPart savedChart = GetStructuralWorksheetPart(spreadsheet, "Data")
                .DrawingsPart!.ChartParts.Single();
            C.Formula[] formulas = savedChart.ChartSpace.Descendants<C.Formula>().ToArray();
            C.Formula[] shiftedFormulas = formulas
                .Where(formula => formula.Text.EndsWith("$A$3:$A$4", System.StringComparison.Ordinal)
                    || formula.Text.EndsWith("$B$3:$B$4", System.StringComparison.Ordinal))
                .ToArray();

            Assert.Equal(2, shiftedFormulas.Length);
            Assert.DoesNotContain(formulas, formula =>
                formula.Text.EndsWith("$A$2:$A$3", System.StringComparison.Ordinal)
                    || formula.Text.EndsWith("$B$2:$B$3", System.StringComparison.Ordinal));
            Assert.All(shiftedFormulas, formula =>
                Assert.DoesNotContain(
                    formula.Parent!.ChildElements,
                    element => element.LocalName.EndsWith("Cache", System.StringComparison.OrdinalIgnoreCase)));
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
        }

        [Fact]
        public void Test_DeleteRows_MaterializesSharedFormulaGroupWhenItsAnchorIsDeleted() {
            string filePath = Path.Combine(_directoryWithFiles, "Excel.StructuralRows.SharedFormula.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet data = document.AddWorksheet("Data");
                for (int row = 2; row <= 4; row++) {
                    data.CellAt(row, 1).SetValue(row);
                }

                SheetData sheetData = data.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;
                for (int row = 2; row <= 4; row++) {
                    Row rowElement = sheetData.Elements<Row>().Single(item => item.RowIndex?.Value == (uint)row);
                    rowElement.Append(new Cell {
                        CellReference = $"B{row}",
                        CellFormula = row == 2
                            ? new CellFormula("A2*2") {
                                FormulaType = CellFormulaValues.Shared,
                                SharedIndex = 7U,
                                Reference = "B2:B4"
                            }
                            : new CellFormula {
                                FormulaType = CellFormulaValues.Shared,
                                SharedIndex = 7U
                            },
                        CellValue = new CellValue((row * 2).ToString())
                    });
                }

                data.DeleteRows(2);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart dataPart = GetStructuralWorksheetPart(spreadsheet, "Data");
                Cell[] formulaCells = dataPart.Worksheet.Descendants<Cell>()
                    .Where(cell => cell.CellFormula != null)
                    .OrderBy(cell => cell.CellReference?.Value)
                    .ToArray();

                Assert.Equal(new[] { "B2", "B3" }, formulaCells.Select(cell => cell.CellReference!.Value).ToArray());
                Assert.Equal(new[] { "A2*2", "A3*2" }, formulaCells.Select(cell => cell.CellFormula!.Text).ToArray());
                Assert.All(formulaCells, cell => {
                    Assert.Null(cell.CellFormula!.FormulaType);
                    Assert.Null(cell.CellFormula.SharedIndex);
                    Assert.Null(cell.CellFormula.Reference);
                    Assert.True(cell.CellFormula.CalculateCell?.Value);
                });
                Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
            }
        }

        [Fact]
        public void Test_InsertRows_MaterializesCrossSheetSharedFormulasAndPreservesExternalReferences() {
            string filePath = Path.Combine(_directoryWithFiles, "Excel.StructuralRows.CrossSheetShared.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet data = document.AddWorksheet("Data");
                ExcelSheet summary = document.AddWorksheet("Summary");
                for (int row = 2; row <= 4; row++) {
                    data.CellAt(row, 1).SetValue(row);
                    summary.CellAt(row, 1).SetValue(row);
                }

                SheetData sheetData = summary.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;
                for (int row = 2; row <= 4; row++) {
                    Row rowElement = sheetData.Elements<Row>().Single(item => item.RowIndex?.Value == (uint)row);
                    rowElement.Append(new Cell {
                        CellReference = $"B{row}",
                        CellFormula = row == 2
                            ? new CellFormula("'Data'!A2") {
                                FormulaType = CellFormulaValues.Shared,
                                SharedIndex = 12U,
                                Reference = "B2:B4"
                            }
                            : new CellFormula {
                                FormulaType = CellFormulaValues.Shared,
                                SharedIndex = 12U
                            }
                    });
                }
                summary.CellFormula(
                    1,
                    4,
                    "[Other.xlsx]Data!A3+'[Other.xlsx]Data'!A3+Data:Other!A3+Other:Data!A3+'Data'!A3");

                data.InsertRows(3);
                document.Save();
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            WorksheetPart summaryPart = GetStructuralWorksheetPart(spreadsheet, "Summary");
            Cell[] sharedCells = summaryPart.Worksheet.Descendants<Cell>()
                .Where(cell => cell.CellReference?.Value is "B2" or "B3" or "B4")
                .OrderBy(cell => cell.CellReference?.Value)
                .ToArray();
            Assert.Equal(new[] { "'Data'!A2", "'Data'!A4", "'Data'!A5" }, sharedCells.Select(cell => cell.CellFormula!.Text));
            Assert.All(sharedCells, cell => Assert.Null(cell.CellFormula!.FormulaType));
            Assert.Equal(
                "[Other.xlsx]Data!A3+'[Other.xlsx]Data'!A3+Data:Other!A3+Other:Data!A3+'Data'!A4",
                summaryPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "D1").CellFormula!.Text);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
        }

        [Fact]
        public void Test_InsertRows_RejectsSharedFormulaFollowerOverflowBeforeMaterialization() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetValue(1);
            sheet.CellAt(2, 1).SetValue(2);
            SheetData sheetData = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;
            Row firstRow = sheetData.Elements<Row>().Single(row => row.RowIndex?.Value == 1U);
            Row secondRow = sheetData.Elements<Row>().Single(row => row.RowIndex?.Value == 2U);
            firstRow.Append(new Cell {
                CellReference = "B1",
                CellFormula = new CellFormula($"A{A1.MaxRows - 1}") {
                    FormulaType = CellFormulaValues.Shared,
                    SharedIndex = 21U,
                    Reference = "B1:B2"
                }
            });
            secondRow.Append(new Cell {
                CellReference = "B2",
                CellFormula = new CellFormula {
                    FormulaType = CellFormulaValues.Shared,
                    SharedIndex = 21U
                }
            });

            Assert.Throws<InvalidOperationException>(() => sheet.InsertRows(1));
            CellFormula[] formulas = sheet.WorksheetPart.Worksheet.Descendants<CellFormula>().ToArray();
            Assert.All(formulas, formula => Assert.Equal(CellFormulaValues.Shared, formula.FormulaType!.Value));
            Assert.Equal($"A{A1.MaxRows - 1}", formulas[0].Text);
            Assert.True(string.IsNullOrEmpty(formulas[1].Text));
        }

        [Fact]
        public void Test_InsertRows_RejectsReferenceOverflowAtomically() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetValue("Keep");
            sheet.CellAt(1, 2).SetValue("Value");
            sheet.CellAt(2, 1).SetValue("Item");
            sheet.CellAt(2, 2).SetValue(1);
            sheet.CellFormula(1, 10, $"A{A1.MaxRows}");
            document.SetNamedRange("Bottom", $"'Data'!A{A1.MaxRows}", save: false);
            sheet.Range("C1").Validation.CustomFormula($"A{A1.MaxRows}>0");
            DataValidation validation = Assert.Single(sheet.WorksheetPart.Worksheet.Descendants<DataValidation>());
            sheet.AddChartFromRange("A1:B2", row: 4, column: 4);
            C.Formula chartFormula = sheet.WorksheetPart.DrawingsPart!.ChartParts.Single()
                .ChartSpace.Descendants<C.Formula>().First();
            chartFormula.Text = $"'Data'!$A${A1.MaxRows}";
            sheet.AddPivotTable(
                "A1:B2",
                "E5",
                rowFields: new[] { "Keep" },
                dataFields: new[] { new ExcelPivotDataField("Value", DataConsolidateFunctionValues.Sum) });
            PivotTablePart pivotPart = sheet.WorksheetPart.PivotTableParts.Single();
            WorksheetSource pivotSource = pivotPart.PivotTableCacheDefinitionPart!
                .PivotCacheDefinition!.CacheSource!.WorksheetSource!;
            Location pivotLocation = pivotPart.PivotTableDefinition!.Location!;
            sheet.AddSparklines("A1:B1", "D1");
            DocumentFormat.OpenXml.Office2010.Excel.Sparkline sparkline = Assert.Single(
                sheet.WorksheetPart.Worksheet.Descendants<DocumentFormat.OpenXml.Office2010.Excel.Sparkline>());

            Assert.Throws<InvalidOperationException>(() => sheet.InsertRows(1));
            sheet.CellFormula(1, 10, "A1");
            Assert.Throws<InvalidOperationException>(() => sheet.InsertRows(1));
            document.SetNamedRange("Bottom", "'Data'!A1", save: false);
            Assert.Throws<InvalidOperationException>(() => sheet.InsertRows(1));
            validation.Formula1!.Text = "A1>0";
            Assert.Throws<InvalidOperationException>(() => sheet.InsertRows(1));
            chartFormula.Text = "'Data'!$A$1";
            pivotSource.Reference = $"A{A1.MaxRows - 1}:B{A1.MaxRows}";
            Assert.Throws<InvalidOperationException>(() => sheet.InsertRows(1));
            pivotSource.Reference = "A1:B2";
            pivotLocation.Reference = $"E{A1.MaxRows - 1}:F{A1.MaxRows}";
            Assert.Throws<InvalidOperationException>(() => sheet.InsertRows(1));
            pivotLocation.Reference = "E5:F6";
            sparkline.ReferenceSequence!.Text = $"D{A1.MaxRows}";
            Assert.Throws<InvalidOperationException>(() => sheet.InsertRows(1));
            sparkline.ReferenceSequence.Text = "D1";
            sheet.AddManualRowPageBreak(A1.MaxRows);
            Assert.Throws<InvalidOperationException>(() => sheet.InsertRows(1));

            Assert.Equal("Keep", sheet.CellAt(1, 1).GetValue<string>());
            Assert.Equal((uint)A1.MaxRows, Assert.Single(sheet.WorksheetPart.Worksheet.GetFirstChild<RowBreaks>()!.Elements<Break>()).Id!.Value);
        }

        [Fact]
        public void Test_StructuralRows_RejectsEditsThroughPivotOutput() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetValue("Region");
            sheet.CellAt(1, 2).SetValue("Sales");
            sheet.CellAt(2, 1).SetValue("East");
            sheet.CellAt(2, 2).SetValue(10);
            sheet.CellAt(3, 1).SetValue("West");
            sheet.CellAt(3, 2).SetValue(20);
            sheet.AddPivotTable(
                "A1:B3",
                "E5",
                rowFields: new[] { "Region" },
                dataFields: new[] { new ExcelPivotDataField("Sales", DataConsolidateFunctionValues.Sum) });

            Assert.Throws<InvalidOperationException>(() => sheet.InsertRows(6));
            Assert.Throws<InvalidOperationException>(() => sheet.DeleteRows(5));
            Assert.Equal("E5:F6", sheet.WorksheetPart.PivotTableParts.Single().PivotTableDefinition!.Location!.Reference!.Value);
        }

        [Fact]
        public void Test_StructuralRows_RejectsOwnerBoundaryChangesBeforeMutation() {
            string filePath = Path.Combine(_directoryWithFiles, "Excel.StructuralRows.OwnedRanges.xlsx");

            using var document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetValue("Value");
            sheet.CellAt(1, 2).SetValue("Result");
            for (int row = 2; row <= 4; row++) {
                sheet.CellAt(row, 1).SetValue(row);
                sheet.CellAt(row, 2).SetValue(row * 2);
            }

            sheet.SetArrayFormula("C2:C4", "A2:A4*2");
            sheet.AddTable("A1:B4", hasHeader: true, name: "OwnedData", OfficeIMO.Excel.TableStyle.TableStyleMedium2);

            InvalidOperationException insertArray = Assert.Throws<InvalidOperationException>(() => sheet.InsertRows(3));
            InvalidOperationException deleteArrayOwner = Assert.Throws<InvalidOperationException>(() => sheet.DeleteRows(2));
            InvalidOperationException deleteTableHeader = Assert.Throws<InvalidOperationException>(() => sheet.DeleteRows(1));

            Assert.Contains("array formula", insertArray.Message, System.StringComparison.OrdinalIgnoreCase);
            Assert.Contains("owner row", deleteArrayOwner.Message, System.StringComparison.OrdinalIgnoreCase);
            Assert.Contains("header row", deleteTableHeader.Message, System.StringComparison.OrdinalIgnoreCase);
            Assert.Equal(2, sheet.CellAt(2, 1).GetValue<int>());
            Assert.Equal("A2:A4*2", sheet.GetFormulaText(2, 3));
            Assert.Equal("A1:B4", sheet.GetTableRange("OwnedData"));
        }

        [Fact]
        public void Test_StructuralRows_RejectsInvalidOrOverflowingOperationsBeforeMutation() {
            string filePath = Path.Combine(_directoryWithFiles, "Excel.StructuralRows.Bounds.xlsx");

            using var document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(A1.MaxRows, 1).SetValue("Last");

            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.InsertRows(0));
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.InsertRows(1, 0));
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.DeleteRows(A1.MaxRows, 2));
            Assert.Throws<InvalidOperationException>(() => sheet.InsertRows(A1.MaxRows));
            Assert.Equal("Last", sheet.CellAt(A1.MaxRows, 1).GetValue<string>());
        }

        private static WorksheetPart GetStructuralWorksheetPart(SpreadsheetDocument spreadsheet, string sheetName) {
            Sheet sheet = spreadsheet.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>()
                .Single(candidate => candidate.Name?.Value == sheetName);
            return (WorksheetPart)spreadsheet.WorkbookPart.GetPartById(sheet.Id!.Value!);
        }
    }
}
