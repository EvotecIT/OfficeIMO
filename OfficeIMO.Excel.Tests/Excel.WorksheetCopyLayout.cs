using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        public static IEnumerable<object[]> WorksheetLayoutCopyRoutes() {
            yield return new object[] { "ValuesIntoExistingWorkbook" };
            yield return new object[] { "ValuesIntoEmptyWorkbook" };
            yield return new object[] { "ValuesFromEditedSourceIntoEmptyWorkbook" };
            yield return new object[] { "PackageIntoExistingWorkbook" };
            yield return new object[] { "PackageIntoEmptyWorkbook" };
            yield return new object[] { "WithinWorkbook" };
            yield return new object[] { "MergeValuesIntoExistingWorkbook" };
            yield return new object[] { "MergePackageIntoExistingWorkbook" };
        }

        [Theory]
        [MemberData(nameof(WorksheetLayoutCopyRoutes))]
        public void Test_WorksheetCopy_KeepsColumnsAndSheetLayout(string route) {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyLayoutSource-" + route + ".xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyLayoutTarget-" + route + ".xlsx");
            CreateLayoutSource(sourcePath);

            string copiedPath = sourcePath;
            if (route == "WithinWorkbook") {
                using var document = ExcelDocument.Load(sourcePath);
                document.CopyWorksheet("Data", "Copy", ExcelSheetNameValidationMode.Strict);
                document.Save();
            } else {
                copiedPath = targetPath;
                bool emptyTarget = route.EndsWith("IntoEmptyWorkbook", StringComparison.Ordinal);
                using (var createdTarget = ExcelDocument.Create(targetPath)) {
                    if (!emptyTarget) {
                        ExcelSheet existing = createdTarget.AddSheet("Existing");
                        existing.Cell(1, 1).SetValue("Existing");
                    }

                    createdTarget.Save();
                }

                if (!emptyTarget) {
                    AddUnrelatedTargetCellFormats(targetPath);
                }

                var mode = route.StartsWith("Package", StringComparison.Ordinal) ? ExcelWorksheetCopyMode.Package : ExcelWorksheetCopyMode.Values;
                bool editedSource = route.Contains("FromEditedSource");
                using var sourceDocument = ExcelDocument.Load(sourcePath, new ExcelLoadOptions {
                    AccessMode = editedSource ? OfficeIMO.DocumentAccessMode.ReadWrite : OfficeIMO.DocumentAccessMode.ReadOnly
                });
                if (editedSource) {
                    // An unsaved edit means the copy reads the source worksheet from memory rather than from its saved part.
                    sourceDocument["Data"].Cell(1, 1).SetValue("Header");
                }

                using var targetDocument = ExcelDocument.Load(targetPath);
                if (route.StartsWith("Merge", StringComparison.Ordinal)) {
                    targetDocument.MergeWorkbookFrom(sourceDocument, new ExcelWorkbookMergeOptions {
                        CopyMode = route.StartsWith("MergePackage", StringComparison.Ordinal) ? ExcelWorksheetCopyMode.Package : ExcelWorksheetCopyMode.Values,
                        TargetSheetNames = new Dictionary<string, string> { ["Data"] = "Copy" },
                        SheetNameValidationMode = ExcelSheetNameValidationMode.Strict
                    });
                } else {
                    targetDocument.CopyWorksheetFrom(sourceDocument, "Data", "Copy", ExcelSheetNameValidationMode.Strict, new ExcelWorksheetCopyOptions { CopyMode = mode });
                }
                targetDocument.Save();
            }

            using (SpreadsheetDocument source = SpreadsheetDocument.Open(sourcePath, false))
            using (SpreadsheetDocument copied = SpreadsheetDocument.Open(copiedPath, false)) {
                Worksheet sourceSheet = GetWorksheetPartByNameForOperations(source, "Data").Worksheet!;
                Worksheet copiedSheet = GetWorksheetPartByNameForOperations(copied, "Copy").Worksheet!;

                Column[] sourceColumns = sourceSheet.GetFirstChild<Columns>()!.Elements<Column>().ToArray();
                Column[] copiedColumns = copiedSheet.GetFirstChild<Columns>()?.Elements<Column>().ToArray() ?? Array.Empty<Column>();
                Assert.Equal(sourceColumns.Length, copiedColumns.Length);
                for (int index = 0; index < sourceColumns.Length; index++) {
                    Column expected = sourceColumns[index];
                    Column actual = copiedColumns[index];
                    Assert.Equal(expected.Min!.Value, actual.Min!.Value);
                    Assert.Equal(expected.Max!.Value, actual.Max!.Value);
                    Assert.Equal(expected.Width?.Value, actual.Width?.Value);
                    Assert.Equal(expected.CustomWidth?.Value, actual.CustomWidth?.Value);
                    Assert.Equal(expected.Hidden?.Value, actual.Hidden?.Value);
                    Assert.Equal(expected.OutlineLevel?.Value, actual.OutlineLevel?.Value);
                    Assert.Equal(expected.Collapsed?.Value, actual.Collapsed?.Value);
                    Assert.Equal(expected.Phonetic?.Value, actual.Phonetic?.Value);
                    Assert.Equal(expected.BestFit?.Value, actual.BestFit?.Value);
                    AssertSameCellFormat(source, expected.Style?.Value, copied, actual.Style?.Value);
                }

                SheetFormatProperties sourceFormat = sourceSheet.GetFirstChild<SheetFormatProperties>()!;
                SheetFormatProperties copiedFormat = copiedSheet.GetFirstChild<SheetFormatProperties>()!;
                Assert.NotNull(copiedFormat);
                Assert.Equal(sourceFormat.DefaultColumnWidth?.Value, copiedFormat.DefaultColumnWidth?.Value);
                Assert.Equal(sourceFormat.BaseColumnWidth?.Value, copiedFormat.BaseColumnWidth?.Value);
                Assert.Equal(sourceFormat.DefaultRowHeight?.Value, copiedFormat.DefaultRowHeight?.Value);
                Assert.Equal(sourceFormat.OutlineLevelColumn?.Value, copiedFormat.OutlineLevelColumn?.Value);

                Assert.Equal(
                    sourceSheet.Descendants<MergeCell>().Select(cell => cell.Reference!.Value),
                    copiedSheet.Descendants<MergeCell>().Select(cell => cell.Reference!.Value));

                Assert.Equal("A2:A10", Assert.Single(copiedSheet.Descendants<DataValidation>()).SequenceOfReferences!.InnerText);
                Assert.Equal("B2:B10", Assert.Single(copiedSheet.Descendants<ConditionalFormatting>()).SequenceOfReferences!.InnerText);

                PageMargins copiedMargins = copiedSheet.GetFirstChild<PageMargins>()!;
                Assert.Equal(1.25D, copiedMargins.Left!.Value);
                PageSetup copiedPageSetup = copiedSheet.GetFirstChild<PageSetup>()!;
                Assert.Equal(OrientationValues.Landscape, copiedPageSetup.Orientation!.Value);
                Assert.Equal(9U, copiedPageSetup.PaperSize!.Value);
                Assert.True(copiedSheet.GetFirstChild<PrintOptions>()!.GridLines!.Value);
                Assert.Equal("&CLayout", copiedSheet.GetFirstChild<HeaderFooter>()!.OddHeader!.Text);

                Assert.Equal("Copy!$B2>1", copiedSheet.Descendants<ConditionalFormattingRule>().Single().GetFirstChild<Formula>()!.Text.Replace("'", string.Empty));
                AssertSameDifferentialFormat(
                    ReadDifferentialFormats(source),
                    sourceSheet.Descendants<ConditionalFormattingRule>().Single().FormatId!.Value,
                    ReadDifferentialFormats(copied),
                    copiedSheet.Descendants<ConditionalFormattingRule>().Single().FormatId!.Value,
                    "cfRule");

                Assert.Equal("FF00B050", copiedSheet.GetFirstChild<SheetProperties>()!.TabColor!.Rgb!.Value);
                int targetWorkbookViews = copied.WorkbookPart!.Workbook!.BookViews?.Elements<WorkbookView>().Count() ?? 1;
                SheetView[] copiedViews = copiedSheet.GetFirstChild<SheetViews>()!.Elements<SheetView>().ToArray();
                Assert.All(copiedViews, view => Assert.True(view.WorkbookViewId!.Value < targetWorkbookViews));
                Assert.Equal(Math.Min(2, targetWorkbookViews), copiedViews.Length);
                SheetView copiedView = copiedViews.Single(view => view.WorkbookViewId!.Value == 0U);
                Assert.False(copiedView.ShowGridLines!.Value);
                Assert.Equal(85U, copiedView.ZoomScale!.Value);
                Assert.Equal(PaneStateValues.Frozen, copiedView.Pane!.State!.Value);
                Assert.Equal("B2", copiedView.Pane.TopLeftCell!.Value);

                Row[] copiedRows = copiedSheet.GetFirstChild<SheetData>()!.Elements<Row>().ToArray();
                Row copiedRow2 = copiedRows.Single(row => row.RowIndex!.Value == 2U);
                Assert.Equal(30D, copiedRow2.Height!.Value);
                Assert.True(copiedRow2.CustomHeight!.Value);
                Row copiedRow3 = copiedRows.Single(row => row.RowIndex!.Value == 3U);
                Assert.True(copiedRow3.CustomFormat!.Value);
                AssertSameCellFormat(source, sourceSheet.Descendants<Row>().Single(row => row.RowIndex!.Value == 3U).StyleIndex!.Value, copied, copiedRow3.StyleIndex?.Value);
                Row copiedRow4 = copiedRows.Single(row => row.RowIndex!.Value == 4U);
                Assert.True(copiedRow4.Hidden!.Value);
                Assert.Equal(1, copiedRow4.OutlineLevel!.Value);

                Sheet copySheetElement = copied.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().Single(sheet => sheet.Name == "Copy");
                uint copyPosition = (uint)copied.WorkbookPart.Workbook.Sheets.Elements<Sheet>().ToList().IndexOf(copySheetElement);
                DefinedName[] copyNames = copied.WorkbookPart.Workbook.DefinedNames?.Elements<DefinedName>()
                    .Where(name => name.LocalSheetId?.Value == copyPosition)
                    .ToArray() ?? Array.Empty<DefinedName>();
                Assert.Equal("OFFSET(Copy!$A$1,0,0,PrintRows,5)", Assert.Single(copyNames, name => name.Name == "_xlnm.Print_Area").Text.Replace("'", string.Empty));
                Assert.Equal("20", Assert.Single(copyNames, name => name.Name == "PrintRows").Text);
                Assert.Equal("Copy!$1:$1", Assert.Single(copyNames, name => name.Name == "_xlnm.Print_Titles").Text.Replace("'", string.Empty));
                Assert.Equal("AllowedValues", copiedSheet.Descendants<DataValidation>().Single().GetFirstChild<Formula1>()!.Text);
                Assert.Equal("Copy!$H$1:$H$2", Assert.Single(copyNames, name => name.Name == "AllowedValues").Text.Replace("'", string.Empty));

                Assert.Empty(new OpenXmlValidator().Validate(copied));
            }

            using (var reloaded = ExcelDocument.Load(copiedPath, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                ExcelSheet copySheet = reloaded["Copy"];
                ExcelColumnInfo columnE = Assert.Single(copySheet.GetColumnDefinitions(), column => column.StartIndex == 5);
                Assert.Equal(30D, columnE.Width!.Value.CharacterCount);
                Assert.True(copySheet.TryGetCellText(1, 1, out string header));
                Assert.Equal("Header", header);
            }

            File.Delete(sourcePath);
            if (File.Exists(targetPath)) File.Delete(targetPath);
        }

        [Theory]
        [InlineData(ExcelWorksheetCopyMode.Values)]
        [InlineData(ExcelWorksheetCopyMode.Package)]
        public void Test_WorkbookMerge_RulesOnOtherSheetsFollowRenamedTables(ExcelWorksheetCopyMode mode) {
            string sourcePath = Path.Combine(_directoryWithFiles, "MergeRenamedTableSource-" + mode + ".xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "MergeRenamedTableTarget-" + mode + ".xlsx");
            CreateTableSource(sourcePath, "Data", "Sales");
            using (SpreadsheetDocument source = SpreadsheetDocument.Open(sourcePath, true)) {
                WorkbookPart workbookPart = source.WorkbookPart!;
                WorksheetPart checkPart = workbookPart.AddNewPart<WorksheetPart>();
                checkPart.Worksheet = new Worksheet(
                    new SheetData(new Row(new Cell { CellReference = "A1", DataType = CellValues.InlineString, InlineString = new InlineString(new Text("Total")) }) { RowIndex = 1U }),
                    new ConditionalFormatting(new ConditionalFormattingRule(new Formula("SUM(Sales[Revenue])>0")) { Type = ConditionalFormatValues.Expression, Priority = 1 }) {
                        SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1" }
                    });
                Sheets sheets = workbookPart.Workbook!.Sheets!;
                uint sheetId = sheets.Elements<Sheet>().Max(sheet => sheet.SheetId!.Value) + 1U;
                sheets.Append(new Sheet { Name = "Check", SheetId = sheetId, Id = workbookPart.GetIdOfPart(checkPart) });
                workbookPart.Workbook.Save();
            }

            CreateTableSource(targetPath, "Existing", "Sales");
            using (var sourceDocument = ExcelDocument.Load(sourcePath, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Load(targetPath)) {
                targetDocument.MergeWorkbookFrom(sourceDocument, new ExcelWorkbookMergeOptions { CopyMode = mode });
                targetDocument.Save();
            }

            using (SpreadsheetDocument target = SpreadsheetDocument.Open(targetPath, false)) {
                string importedTable = GetWorksheetPartByNameForOperations(target, "Data").TableDefinitionParts.Single().Table!.Name!.Value!;
                Assert.NotEqual("Sales", importedTable);
                Worksheet check = GetWorksheetPartByNameForOperations(target, "Check").Worksheet!;
                Assert.Equal("SUM(" + importedTable + "[Revenue])>0", check.Descendants<ConditionalFormattingRule>().Single().GetFirstChild<Formula>()!.Text);
                Assert.Empty(new OpenXmlValidator().Validate(target));
            }

            using (var reloaded = ExcelDocument.Load(targetPath, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.True(reloaded["Check"].TryGetCellText(1, 1, out string total));
                Assert.Equal("Total", total);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        private static void CreateLayoutSource(string path) {
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddSheet("Data");
                sheet.Cell(1, 1).SetValue("Header");
                sheet.Cell(2, 1).SetValue("Value");
                sheet.Cell(2, 2).SetValue(5);
                sheet.Cell(1, 8).SetValue("Yes");
                sheet.Cell(2, 8).SetValue("No");
                document.Save();
            }

            using SpreadsheetDocument package = SpreadsheetDocument.Open(path, true);
            WorkbookPart workbookPart = package.WorkbookPart!;
            Stylesheet stylesheet = workbookPart.WorkbookStylesPart!.Stylesheet!;
            stylesheet.Fills ??= new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }), new Fill(new PatternFill { PatternType = PatternValues.Gray125 }));
            stylesheet.Fills.Append(new Fill(new PatternFill(new ForegroundColor { Rgb = "FFFFC000" }) { PatternType = PatternValues.Solid }));
            stylesheet.Fills.Count = (uint)stylesheet.Fills.ChildElements.Count;
            uint fillId = stylesheet.Fills.Count.Value - 1U;
            stylesheet.CellFormats!.Append(new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = fillId, BorderId = 0U, FormatId = 0U, ApplyFill = true });
            stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.ChildElements.Count;
            uint columnStyle = stylesheet.CellFormats.Count.Value - 1U;
            stylesheet.Save();

            Worksheet worksheet = GetWorksheetPartByNameForOperations(package, "Data").Worksheet!;
            worksheet.GetFirstChild<Columns>()?.Remove();
            worksheet.GetFirstChild<SheetFormatProperties>()?.Remove();
            var format = new SheetFormatProperties { BaseColumnWidth = 10U, DefaultColumnWidth = 12.5D, DefaultRowHeight = 16.5D, CustomHeight = true, OutlineLevelColumn = 1 };
            var columns = new Columns(
                new Column { Min = 3U, Max = 4U, Width = 8D, CustomWidth = true, Style = columnStyle, Phonetic = true },
                new Column { Min = 5U, Max = 5U, Width = 30D, CustomWidth = true, BestFit = true },
                new Column { Min = 7U, Max = 8U, Width = 10D, CustomWidth = true, Hidden = true, OutlineLevel = 1, Collapsed = true });
            SheetData sheetData = worksheet.GetFirstChild<SheetData>()!;
            worksheet.InsertBefore(format, sheetData);
            worksheet.InsertBefore(columns, sheetData);

            worksheet.GetFirstChild<MergeCells>()?.Remove();
            worksheet.GetFirstChild<ConditionalFormatting>()?.Remove();
            worksheet.GetFirstChild<DataValidations>()?.Remove();
            worksheet.GetFirstChild<PageMargins>()?.Remove();
            worksheet.GetFirstChild<PageSetup>()?.Remove();
            worksheet.GetFirstChild<PrintOptions>()?.Remove();
            worksheet.GetFirstChild<HeaderFooter>()?.Remove();
            OpenXmlElement anchor = worksheet.Elements<SheetData>().Single();
            anchor = worksheet.InsertAfter(new MergeCells(new MergeCell { Reference = "C1:D1" }) { Count = 1U }, anchor);
            anchor = worksheet.InsertAfter(new ConditionalFormatting(new ConditionalFormattingRule(new Formula("B2>1")) { Type = ConditionalFormatValues.Expression, Priority = 1 }) {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "B2:B10" }
            }, anchor);
            anchor = worksheet.InsertAfter(new DataValidations(new DataValidation(new Formula1("AllowedValues")) {
                Type = DataValidationValues.List,
                AllowBlank = true,
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "A2:A10" }
            }) { Count = 1U }, anchor);
            anchor = worksheet.InsertAfter(new PrintOptions { GridLines = true }, anchor);
            anchor = worksheet.InsertAfter(new PageMargins { Left = 1.25D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D }, anchor);
            anchor = worksheet.InsertAfter(new PageSetup { Orientation = OrientationValues.Landscape, PaperSize = 9U }, anchor);
            worksheet.InsertAfter(new HeaderFooter(new OddHeader("&CLayout")), anchor);

            worksheet.GetFirstChild<SheetProperties>()?.Remove();
            worksheet.GetFirstChild<SheetViews>()?.Remove();
            worksheet.InsertAt(new SheetProperties(new TabColor { Rgb = "FF00B050" }), 0);
            worksheet.InsertBefore(new SheetViews(new SheetView(
                new Pane { HorizontalSplit = 1D, VerticalSplit = 1D, TopLeftCell = "B2", ActivePane = PaneValues.BottomRight, State = PaneStateValues.Frozen },
                new Selection { Pane = PaneValues.BottomRight, ActiveCell = "B2", SequenceOfReferences = new ListValue<StringValue> { InnerText = "B2" } }) {
                WorkbookViewId = 0U,
                ShowGridLines = false,
                ZoomScale = 85U
            }, new SheetView { WorkbookViewId = 1U, ZoomScale = 120U }), worksheet.GetFirstChild<SheetFormatProperties>()!);

            Row row2 = sheetData.Elements<Row>().Single(row => row.RowIndex!.Value == 2U);
            row2.Height = 30D;
            row2.CustomHeight = true;
            sheetData.Append(new Row { RowIndex = 3U, StyleIndex = columnStyle, CustomFormat = true });
            sheetData.Append(new Row { RowIndex = 4U, Hidden = true, OutlineLevel = 1 });

            stylesheet.DifferentialFormats ??= new DifferentialFormats();
            stylesheet.DifferentialFormats.Append(new DifferentialFormat(new Font(new Bold())));
            stylesheet.DifferentialFormats.Count = (uint)stylesheet.DifferentialFormats.ChildElements.Count;
            stylesheet.Save();
            ConditionalFormattingRule rule = worksheet.Descendants<ConditionalFormattingRule>().Single();
            rule.FormatId = stylesheet.DifferentialFormats.Count.Value - 1U;
            rule.GetFirstChild<Formula>()!.Text = "Data!$B2>1";
            worksheet.Save();

            workbookPart.Workbook!.DefinedNames ??= new DefinedNames();
            // A second workbook window; the sheet has a view for each.
            BookViews bookViews = workbookPart.Workbook.BookViews ??= new BookViews(new WorkbookView());
            bookViews.Append(new WorkbookView { XWindow = 200, YWindow = 200 });
            workbookPart.Workbook.DefinedNames.Append(
                new DefinedName("Data!$H$1:$H$2") { Name = "AllowedValues" },
                new DefinedName("OFFSET(Data!$A$1,0,0,PrintRows,5)") { Name = "_xlnm.Print_Area", LocalSheetId = 0U },
                new DefinedName("20") { Name = "PrintRows" },
                new DefinedName("Data!$1:$1") { Name = "_xlnm.Print_Titles", LocalSheetId = 0U });
            workbookPart.Workbook.Save();
            Assert.Empty(new OpenXmlValidator().Validate(package).Select(error => error.Path?.XPath + " " + error.Description));
        }

        private static void AddUnrelatedTargetCellFormats(string path) {
            using SpreadsheetDocument package = SpreadsheetDocument.Open(path, true);
            Stylesheet stylesheet = package.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            for (int index = 0; index < 3; index++) {
                stylesheet.Fills!.Append(new Fill(new PatternFill(new ForegroundColor { Rgb = "FF00000" + index }) { PatternType = PatternValues.Solid }));
                stylesheet.CellFormats!.Append(new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = (uint)stylesheet.Fills.ChildElements.Count - 1U, BorderId = 0U, FormatId = 0U, ApplyFill = true });
            }

            stylesheet.Fills!.Count = (uint)stylesheet.Fills.ChildElements.Count;
            stylesheet.CellFormats!.Count = (uint)stylesheet.CellFormats.ChildElements.Count;
            stylesheet.Save();
        }

        private static void AssertSameCellFormat(SpreadsheetDocument source, uint? sourceIndex, SpreadsheetDocument target, uint? targetIndex) {
            if (sourceIndex == null) {
                Assert.True(targetIndex == null || targetIndex == 0U);
                return;
            }

            Assert.NotNull(targetIndex);
            Stylesheet sourceStyles = source.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            Stylesheet targetStyles = target.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            CellFormat sourceFormat = sourceStyles.CellFormats!.Elements<CellFormat>().ElementAt((int)sourceIndex.Value);
            CellFormat targetFormat = targetStyles.CellFormats!.Elements<CellFormat>().ElementAt((int)targetIndex!.Value);
            Fill sourceFill = sourceStyles.Fills!.Elements<Fill>().ElementAt((int)sourceFormat.FillId!.Value);
            Fill targetFill = targetStyles.Fills!.Elements<Fill>().ElementAt((int)targetFormat.FillId!.Value);
            Assert.Equal(sourceFill.OuterXml, targetFill.OuterXml);
        }
    }
}
