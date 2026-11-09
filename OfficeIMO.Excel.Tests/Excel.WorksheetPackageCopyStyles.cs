using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_CopyWorksheetFrom_PackageModeMapsTableAndConditionalFormatDxfs() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyTableDxfSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyTableDxfTarget.xlsx");
            CreateTableSource(sourcePath, "Data", "Sales");
            AddTableAndConditionalFormatDxfs(sourcePath, "Data");

            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                targetDocument.AddSheet("Existing").Cell(1, 1).SetValue("Existing");
                targetDocument.Save();
            }

            AddDummyDifferentialFormat(targetPath);

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Load(targetPath)) {
                targetDocument.CopyWorksheetFrom(sourceDocument, "Data", "Imported", ExcelSheetNameValidationMode.Strict, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package
                });
                targetDocument.Save();
            }

            using (SpreadsheetDocument source = SpreadsheetDocument.Open(sourcePath, false))
            using (SpreadsheetDocument target = SpreadsheetDocument.Open(targetPath, false)) {
                List<DifferentialFormat> sourceDxfs = ReadDifferentialFormats(source);
                List<DifferentialFormat> targetDxfs = ReadDifferentialFormats(target);
                Table sourceTable = GetWorksheetPartByNameForOperations(source, "Data").TableDefinitionParts.Single().Table!;
                WorksheetPart importedPart = GetWorksheetPartByNameForOperations(target, "Imported");
                Table importedTable = importedPart.TableDefinitionParts.Single().Table!;

                var sourceTableIds = ReadTableDifferentialFormatIds(sourceTable);
                var importedTableIds = ReadTableDifferentialFormatIds(importedTable);
                Assert.Equal(10, sourceTableIds.Count);
                Assert.Equal(sourceTableIds.Select(_ => _.Name), importedTableIds.Select(_ => _.Name));
                for (int index = 0; index < sourceTableIds.Count; index++) {
                    AssertSameDifferentialFormat(sourceDxfs, sourceTableIds[index].Id, targetDxfs, importedTableIds[index].Id, sourceTableIds[index].Name);
                }

                ConditionalFormattingRule[] sourceRules = GetWorksheetPartByNameForOperations(source, "Data").Worksheet.Descendants<ConditionalFormattingRule>()
                    .OrderBy(_ => _.Priority!.Value)
                    .ToArray();
                ConditionalFormattingRule[] importedRules = importedPart.Worksheet.Descendants<ConditionalFormattingRule>()
                    .OrderBy(_ => _.Priority!.Value)
                    .ToArray();
                Assert.Equal(2, sourceRules.Length);
                Assert.Equal(sourceRules.Length, importedRules.Length);
                for (int index = 0; index < sourceRules.Length; index++) {
                    AssertSameDifferentialFormat(sourceDxfs, sourceRules[index].FormatId!.Value, targetDxfs, importedRules[index].FormatId!.Value, "cfRule " + index);
                }

                Assert.Empty(new OpenXmlValidator().Validate(target));
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_CopyWorksheetFrom_PackageModeKeepsDxfNumberFormatIdsUniqueInTarget() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyDxfNumberFormatSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyDxfNumberFormatTarget.xlsx");
            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                sourceDocument.AddSheet("Data").Cell(1, 1).SetValue(0.25);
                sourceDocument.Save();
            }

            using (SpreadsheetDocument source = SpreadsheetDocument.Open(sourcePath, true)) {
                Stylesheet stylesheet = source.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
                stylesheet.DifferentialFormats ??= new DifferentialFormats();
                stylesheet.DifferentialFormats.Append(new DifferentialFormat(new NumberingFormat { NumberFormatId = 164U, FormatCode = "0.0%" }));
                stylesheet.DifferentialFormats.Count = (uint)stylesheet.DifferentialFormats.Elements<DifferentialFormat>().Count();
                stylesheet.Save();
                Worksheet worksheet = GetWorksheetPartByNameForOperations(source, "Data").Worksheet;
                worksheet.InsertAfter(CreateExpressionRule("A1", "A1>0", 0U, 1), worksheet.GetFirstChild<SheetData>());
                worksheet.Save();
            }

            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                targetDocument.AddSheet("Existing").Cell(1, 1).SetValue(12.5).SetNumberFormat("#,##0.00 \"kg\"");
                targetDocument.Save();
            }

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Load(targetPath)) {
                targetDocument.CopyWorksheetFrom(sourceDocument, "Data", "Imported", ExcelSheetNameValidationMode.Strict, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package
                });
                targetDocument.Save();
            }

            using (SpreadsheetDocument target = SpreadsheetDocument.Open(targetPath, false)) {
                Stylesheet stylesheet = target.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
                ConditionalFormattingRule rule = GetWorksheetPartByNameForOperations(target, "Imported").Worksheet.Descendants<ConditionalFormattingRule>().Single();
                DifferentialFormat dxf = stylesheet.DifferentialFormats!.Elements<DifferentialFormat>().ElementAt((int)rule.FormatId!.Value);
                NumberingFormat dxfFormat = dxf.Elements<NumberingFormat>().Single();
                Assert.Equal("0.0%", dxfFormat.FormatCode!.Value);

                // Excel keeps one number format per id across numFmts and dxfs, so the copied id must not name the target's own format.
                NumberingFormat[] sameId = stylesheet.NumberingFormats!.Elements<NumberingFormat>()
                    .Where(_ => _.NumberFormatId?.Value == dxfFormat.NumberFormatId!.Value)
                    .ToArray();
                Assert.All(sameId, _ => Assert.Equal("0.0%", _.FormatCode!.Value));
                Assert.Contains(stylesheet.NumberingFormats!.Elements<NumberingFormat>(), _ => _.FormatCode!.Value == "#,##0.00 \"kg\"");
                Assert.Empty(new OpenXmlValidator().Validate(target));
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_CopyWorksheetFrom_PackageModeAddsEachSourceStyleOnce() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyStyleGrowthSource.xlsx");
            string twinSourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyStyleGrowthTwin.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyStyleGrowthTarget.xlsx");
            using (var sourceDocument = ExcelDocument.Create(sourcePath)) {
                foreach (string name in new[] { "First", "Second" }) {
                    ExcelSheet sheet = sourceDocument.AddSheet(name);
                    sheet.Cell(1, 1).SetValue("Region").SetBold(true);
                    sheet.Cell(1, 2).SetValue("Weight").SetBold(true).SetFillColor(OfficeColor.Parse("#FCE4D6"));
                    sheet.Cell(2, 1).SetValue("North").SetItalic(true);
                    sheet.Cell(2, 2).SetValue(1234.5).SetNumberFormat("#,##0.00 \"kg\"");
                    sheet.Cell(3, 1).SetValue("South");
                    sheet.Cell(3, 2).SetValue(0.5).SetNumberFormat("0.0%").SetFillColor(OfficeColor.Parse("#DDEBF7"));
                    sheet.AddTable("A1:B3", hasHeader: true, name: name + "Table");
                }

                sourceDocument.Save();
            }

            AddTableAndConditionalFormatDxfs(sourcePath, "First");
            AddTableAndConditionalFormatDxfs(sourcePath, "Second");
            File.Copy(sourcePath, twinSourcePath, overwrite: true);

            using (var targetDocument = ExcelDocument.Create(targetPath)) {
                targetDocument.AddSheet("Existing").Cell(1, 1).SetValue("Existing").SetBold(true);
                targetDocument.Save();
            }

            var copyOptions = new ExcelWorksheetCopyOptions { CopyMode = ExcelWorksheetCopyMode.Package };
            using (var sourceDocument = ExcelDocument.Load(sourcePath, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Load(targetPath)) {
                targetDocument.CopyWorksheetFrom(sourceDocument, "First", "First", ExcelSheetNameValidationMode.Strict, copyOptions);
                targetDocument.Save();
            }

            string afterFirstCopy = ReadStyleCounts(targetPath);

            using (var sourceDocument = ExcelDocument.Load(sourcePath, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            using (var twinDocument = ExcelDocument.Load(twinSourcePath, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Load(targetPath)) {
                targetDocument.CopyWorksheetFrom(sourceDocument, "Second", "Second", ExcelSheetNameValidationMode.Strict, copyOptions);
                targetDocument.CopyWorksheetFrom(sourceDocument, "First", "First again", ExcelSheetNameValidationMode.Strict, copyOptions);
                targetDocument.CopyWorksheetFrom(twinDocument, "Second", "Twin", ExcelSheetNameValidationMode.Strict, copyOptions);
                targetDocument.Save();
            }

            Assert.Equal(afterFirstCopy, ReadStyleCounts(targetPath));

            using (SpreadsheetDocument source = SpreadsheetDocument.Open(sourcePath, false))
            using (SpreadsheetDocument target = SpreadsheetDocument.Open(targetPath, false)) {
                foreach ((string sourceSheet, string targetSheet) in new[] {
                    ("First", "First"),
                    ("Second", "Second"),
                    ("First", "First again"),
                    ("Second", "Twin")
                }) {
                    Cell[] sourceCells = GetWorksheetPartByNameForOperations(source, sourceSheet).Worksheet.Descendants<Cell>().ToArray();
                    Cell[] targetCells = GetWorksheetPartByNameForOperations(target, targetSheet).Worksheet.Descendants<Cell>().ToArray();
                    Assert.Equal(sourceCells.Select(_ => _.CellReference!.Value), targetCells.Select(_ => _.CellReference!.Value));
                    for (int index = 0; index < sourceCells.Length; index++) {
                        Assert.Equal(DescribeCellFormat(source, sourceCells[index]), DescribeCellFormat(target, targetCells[index]));
                    }
                }

                Assert.Empty(new OpenXmlValidator().Validate(target));
            }

            File.Delete(sourcePath);
            File.Delete(twinSourcePath);
            File.Delete(targetPath);
        }

        [Theory]
        [InlineData("font")]
        [InlineData("fill")]
        [InlineData("border")]
        [InlineData("dxf")]
        public void Test_CopyWorksheetFrom_PackageModeKeepsFixedColorsAfterTargetThemeChanges(string kind) {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyFixedColorSource-" + kind + ".xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyFixedColorTarget-" + kind + ".xlsx");
            using (var source = ExcelDocument.Create(sourcePath)) {
                source.AddSheet("Data").Cell(1, 1).SetValue("Source");
                source.Save();
            }
            using (var target = ExcelDocument.Create(targetPath)) {
                target.AddSheet("Existing").Cell(1, 1).SetValue("Target");
                SetCopyAccentColor(target, "FF0000");
                target.Save();
            }
            AddCopyColorBinding(sourcePath, "Data", kind, themeLinked: false);
            AddCopyColorBinding(targetPath, "Existing", kind, themeLinked: true);
            var options = new ExcelWorksheetCopyOptions { CopyMode = ExcelWorksheetCopyMode.Package };
            using (var source = ExcelDocument.Load(sourcePath, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            using (var target = ExcelDocument.Load(targetPath)) {
                target.CopyWorksheetFrom(source, "Data", "Imported", ExcelSheetNameValidationMode.Strict, options);
                target.Save();
            }
            string firstCopyCounts = ReadStyleCounts(targetPath);
            using (var source = ExcelDocument.Load(sourcePath, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            using (var target = ExcelDocument.Load(targetPath)) {
                target.CopyWorksheetFrom(source, "Data", "Imported again", ExcelSheetNameValidationMode.Strict, options);
                SetCopyAccentColor(target, "0000FF");
                target.Save();
            }
            using (SpreadsheetDocument target = SpreadsheetDocument.Open(targetPath, false)) {
                foreach (string sheet in new[] { "Imported", "Imported again" }) {
                    ColorType color = ReadCopyStyleColor(target, sheet, kind);
                    Assert.Equal("FFFF0000", color.Rgb?.Value);
                    Assert.Null(color.Theme);
                    Assert.Null(color.Indexed);
                }
                ColorType existing = ReadCopyStyleColor(target, "Existing", kind);
                Assert.Equal(4U, existing.Theme?.Value);
                Assert.Null(existing.Rgb);
                XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
                using (Stream stream = target.WorkbookPart!.ThemePart!.GetStream()) {
                    Assert.Equal("0000FF", XDocument.Load(stream).Descendants(drawing + "accent1").Elements(drawing + "srgbClr").Single().Attribute("val")?.Value);
                }
                Assert.Empty(new OpenXmlValidator().Validate(target));
            }
            Assert.Equal(firstCopyCounts, ReadStyleCounts(targetPath));
            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Theory]
        [InlineData("font", "theme", true)]
        [InlineData("font", "theme", false)]
        [InlineData("font", "tint", true)]
        [InlineData("font", "tint", false)]
        [InlineData("font", "indexed", true)]
        [InlineData("font", "indexed", false)]
        [InlineData("fill", "theme", true)]
        [InlineData("border", "theme", true)]
        [InlineData("dxf", "theme", true)]
        [InlineData("dxf", "theme", false)]
        public void Test_CopyWorksheetFrom_PackageModeReusesEntriesWhoseColorsResolveAlike(string kind, string colorKind, bool sameColor) {
            string name = kind + "-" + colorKind + "-" + sameColor;
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyResolvedColorSource-" + name + ".xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyResolvedColorTarget-" + name + ".xlsx");
            XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
            using (var source = ExcelDocument.Create(sourcePath)) {
                source.AddSheet("Data").Cell(1, 1).SetValue("Source");
                SetCopyThemeColor(source, "dk1", new XElement(drawing + "sysClr", new XAttribute("val", "windowText"), new XAttribute("lastClr", "000000")));
                SetCopyAccentColor(source, "4472C4");
                source.Save();
            }

            // The target's theme writes the same colors another way, or holds other colors.
            string dark = sameColor ? "000000" : "1F1F1F";
            string accent = sameColor ? "4472C4" : "ED7D31";
            using (var target = ExcelDocument.Create(targetPath)) {
                target.AddSheet("Existing").Cell(1, 1).SetValue("Target");
                SetCopyThemeColor(target, "dk1", new XElement(drawing + "srgbClr", new XAttribute("val", dark)));
                SetCopyAccentColor(target, accent);
                target.Save();
            }

            AddCopyColorBinding(sourcePath, "Data", kind, CreateCopyReferenceColor(colorKind));
            AddCopyColorBinding(targetPath, "Existing", kind, CreateCopyReferenceColor(colorKind));
            if (colorKind == "indexed" && !sameColor) {
                SetCopyIndexedColor(targetPath, 10, "FF00FF00");
            }

            string beforeCopy = ReadStyleCounts(targetPath);
            var options = new ExcelWorksheetCopyOptions { CopyMode = ExcelWorksheetCopyMode.Package };
            using (var source = ExcelDocument.Load(sourcePath, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            using (var target = ExcelDocument.Load(targetPath)) {
                target.CopyWorksheetFrom(source, "Data", "Imported", ExcelSheetNameValidationMode.Strict, options);
                target.CopyWorksheetFrom(source, "Data", "Imported again", ExcelSheetNameValidationMode.Strict, options);
                target.Save();
            }

            string afterCopy = ReadStyleCounts(targetPath);
            using (SpreadsheetDocument target = SpreadsheetDocument.Open(targetPath, false)) {
                ColorType existing = ReadCopyStyleColor(target, "Existing", kind);
                ColorType imported = ReadCopyStyleColor(target, "Imported", kind);
                Assert.Equal(CreateCopyReferenceColor(colorKind).GetAttributes().Select(_ => _.LocalName + "=" + _.Value), existing.GetAttributes().Select(_ => _.LocalName + "=" + _.Value));
                Assert.Equal(ReadCopyStyleReference(target, "Imported", kind), ReadCopyStyleReference(target, "Imported again", kind));
                if (sameColor) {
                    Assert.Equal(beforeCopy, afterCopy);
                    Assert.Equal(ReadCopyStyleReference(target, "Existing", kind), ReadCopyStyleReference(target, "Imported", kind));
                } else {
                    Assert.NotEqual(beforeCopy, afterCopy);
                    Assert.NotEqual(ReadCopyStyleReference(target, "Existing", kind), ReadCopyStyleReference(target, "Imported", kind));
                    Assert.NotNull(imported.Rgb);
                    Assert.Null(imported.Theme);
                    Assert.Null(imported.Indexed);
                    Assert.Null(imported.Tint);
                }

                Assert.Empty(new OpenXmlValidator().Validate(target));
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Theory]
        [InlineData("border")]
        [InlineData("fill")]
        public void Test_CopyWorksheetFrom_PackageModeReusesEntriesThatDrawNothingAlike(string kind) {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyEmptyEntrySource-" + kind + ".xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyEmptyEntryTarget-" + kind + ".xlsx");
            using (var source = ExcelDocument.Create(sourcePath)) {
                ExcelSheet sheet = source.AddSheet("Data");
                sheet.Cell(1, 1).SetValue("Source");
                sheet.Cell(2, 1).SetValue(12.5);
                source.Save();
            }

            using (var target = ExcelDocument.Create(targetPath)) {
                target.AddSheet("Existing").Cell(1, 1).SetValue("Target");
                target.Save();
            }

            using (SpreadsheetDocument source = SpreadsheetDocument.Open(sourcePath, true)) {
                Stylesheet styles = source.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
                if (kind == "border") {
                    // Excel writes each side of its default border; OfficeIMO's target writes none.
                    WriteExcelDefaultStyles(styles, font: false, border: true);
                } else {
                    // A fill with no pattern draws what the target's none fill draws.
                    styles.Fills!.Append(new Fill());
                    styles.Fills.Count = (uint)styles.Fills.Elements<Fill>().Count();
                    styles.CellFormats!.Append(new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = styles.Fills.Count - 1U, BorderId = 0U, FormatId = 0U });
                    styles.CellFormats.Count = (uint)styles.CellFormats.Elements<CellFormat>().Count();
                    WorksheetPart part = GetWorksheetPartByNameForOperations(source, "Data");
                    part.Worksheet.Descendants<Cell>().Single(_ => _.CellReference?.Value == "A1").StyleIndex = styles.CellFormats.Count - 1U;
                    part.Worksheet.Save();
                }

                styles.Save();
            }

            string beforeCopy = ReadStyleCounts(targetPath);
            using (var source = ExcelDocument.Load(sourcePath, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            using (var target = ExcelDocument.Load(targetPath)) {
                target.CopyWorksheetFrom(source, "Data", "Imported", ExcelSheetNameValidationMode.Strict, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package
                });
                target.Save();
            }

            Assert.Equal(beforeCopy, ReadStyleCounts(targetPath));
            using (SpreadsheetDocument target = SpreadsheetDocument.Open(targetPath, false)) {
                Cell[] cells = GetWorksheetPartByNameForOperations(target, "Imported").Worksheet.Descendants<Cell>().ToArray();
                Assert.Equal(new[] { "A1", "A2" }, cells.Select(_ => _.CellReference!.Value));
                Assert.All(cells, _ => Assert.Equal(0U, _.StyleIndex?.Value ?? 0U));
                Assert.Empty(new OpenXmlValidator().Validate(target));
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void Test_CopyWorksheetFrom_PackageModeKeepsExcelDefaultFormatOnCopiedCells() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WorksheetCopyExcelDefaultSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "WorksheetCopyExcelDefaultTarget.xlsx");
            XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
            using (var source = ExcelDocument.Create(sourcePath)) {
                ExcelSheet sheet = source.AddSheet("Data");
                sheet.Cell(1, 1).SetValue("Region");
                sheet.Cell(1, 2).SetValue("Weight");
                sheet.Cell(2, 1).SetValue("North");
                sheet.Cell(2, 2).SetValue(12.5);
                SetCopyThemeColor(source, "dk1", new XElement(drawing + "sysClr", new XAttribute("val", "windowText"), new XAttribute("lastClr", "000000")));
                source.Save();
            }

            using (var target = ExcelDocument.Create(targetPath)) {
                target.AddSheet("Existing").Cell(1, 1).SetValue("Existing");
                SetCopyThemeColor(target, "dk1", new XElement(drawing + "srgbClr", new XAttribute("val", "000000")));
                target.Save();
            }

            // The source is laid out as Excel writes a new workbook. The target has Excel's default font, but keeps
            // OfficeIMO's default border.
            using (SpreadsheetDocument source = SpreadsheetDocument.Open(sourcePath, true)) {
                Stylesheet styles = source.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
                WriteExcelDefaultStyles(styles, font: true, border: true);
                styles.Save();
            }

            using (SpreadsheetDocument target = SpreadsheetDocument.Open(targetPath, true)) {
                Stylesheet styles = target.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
                WriteExcelDefaultStyles(styles, font: true, border: false);
                styles.Save();
            }

            string beforeCopy = ReadStyleCounts(targetPath);
            using (var source = ExcelDocument.Load(sourcePath, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            using (var target = ExcelDocument.Load(targetPath)) {
                target.CopyWorksheetFrom(source, "Data", "Imported", ExcelSheetNameValidationMode.Strict, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package
                });
                target.Save();
            }

            Assert.Equal(beforeCopy, ReadStyleCounts(targetPath));
            using (SpreadsheetDocument source = SpreadsheetDocument.Open(sourcePath, false))
            using (SpreadsheetDocument target = SpreadsheetDocument.Open(targetPath, false)) {
                Cell[] sourceCells = GetWorksheetPartByNameForOperations(source, "Data").Worksheet.Descendants<Cell>().ToArray();
                Cell[] importedCells = GetWorksheetPartByNameForOperations(target, "Imported").Worksheet.Descendants<Cell>().ToArray();
                Assert.Equal(sourceCells.Select(_ => _.CellReference!.Value), importedCells.Select(_ => _.CellReference!.Value));
                Assert.All(sourceCells, _ => Assert.Null(_.StyleIndex));
                Assert.All(importedCells, _ => Assert.Null(_.StyleIndex));
                Assert.Empty(new OpenXmlValidator().Validate(target));
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        private static void SetCopyAccentColor(ExcelDocument document, string rgb) {
            XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
            SetCopyThemeColor(document, "accent1", new XElement(drawing + "srgbClr", new XAttribute("val", rgb)));
        }

        private static void SetCopyThemeColor(ExcelDocument document, string slot, XElement color) {
            document.EnsureWorkbookTheme();
            XDocument theme = XDocument.Parse(document.GetWorkbookThemeXml()!);
            XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
            theme.Descendants(drawing + slot).Single().ReplaceNodes(color);
            document.SetWorkbookThemeXml(theme.ToString(SaveOptions.DisableFormatting));
        }

        /// <summary>Gives the workbook a custom palette that changes one indexed color.</summary>
        private static void SetCopyIndexedColor(string path, int index, string argb) {
            using SpreadsheetDocument document = SpreadsheetDocument.Open(path, true);
            Stylesheet styles = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            styles.Colors = new Colors(new IndexedColors(Enumerable.Range(0, 64).Select(_ => new RgbColor { Rgb = _ == index ? argb : "FF000000" })));
            styles.Save();
        }

        private static Color CreateCopyReferenceColor(string colorKind) {
            if (colorKind == "theme") {
                return new Color { Theme = 1U };
            }

            if (colorKind == "tint") {
                return new Color { Theme = 4U, Tint = 0.39997558519241921D };
            }

            return new Color { Indexed = 10U };
        }

        /// <summary>Replaces the default font, the default border, or both, with the forms Excel writes.</summary>
        private static void WriteExcelDefaultStyles(Stylesheet styles, bool font, bool border) {
            if (font) {
                Font first = styles.Fonts!.Elements<Font>().First();
                first.RemoveAllChildren();
                first.Append(
                    new FontSize { Val = 11D },
                    new Color { Theme = 1U },
                    new FontName { Val = "Calibri" },
                    new FontFamilyNumbering { Val = 2 },
                    new FontScheme { Val = FontSchemeValues.Minor });
            }

            if (border) {
                Border first = styles.Borders!.Elements<Border>().First();
                first.RemoveAllChildren();
                first.Append(new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder());
            }
        }

        private static uint ReadCopyStyleReference(SpreadsheetDocument document, string sheetName, string kind) {
            Worksheet worksheet = GetWorksheetPartByNameForOperations(document, sheetName).Worksheet;
            if (kind == "dxf") {
                return worksheet.Descendants<ConditionalFormattingRule>().Single().FormatId!.Value;
            }

            return worksheet.Descendants<Cell>().Single(_ => _.CellReference?.Value == "A1").StyleIndex?.Value ?? 0U;
        }

        private static void AddCopyColorBinding(string path, string sheetName, string kind, bool themeLinked) {
            Color color = themeLinked ? new Color { Theme = 4U } : new Color { Rgb = "FFFF0000" };
            AddCopyColorBinding(path, sheetName, kind, color);
        }

        private static void AddCopyColorBinding(string path, string sheetName, string kind, Color color) {
            using SpreadsheetDocument document = SpreadsheetDocument.Open(path, true);
            Stylesheet styles = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            WorksheetPart part = GetWorksheetPartByNameForOperations(document, sheetName);
            if (kind == "dxf") {
                styles.DifferentialFormats ??= new DifferentialFormats();
                uint index = (uint)styles.DifferentialFormats.Elements<DifferentialFormat>().Count();
                styles.DifferentialFormats.Append(new DifferentialFormat(new Font(color)));
                styles.DifferentialFormats.Count = index + 1U;
                part.Worksheet.InsertAfter(CreateExpressionRule("A1", "1=1", index, 1), part.Worksheet.GetFirstChild<SheetData>());
            } else {
                var format = new CellFormat { FontId = 0U, FillId = 0U, BorderId = 0U, NumberFormatId = 0U, FormatId = 0U };
                if (kind == "font") {
                    format.FontId = (uint)styles.Fonts!.Elements<Font>().Count();
                    styles.Fonts.Append(new Font(color));
                    styles.Fonts.Count = format.FontId!.Value + 1U;
                    format.ApplyFont = true;
                } else if (kind == "fill") {
                    format.FillId = (uint)styles.Fills!.Elements<Fill>().Count();
                    var foreground = new ForegroundColor();
                    foreground.SetAttributes(color.GetAttributes());
                    styles.Fills.Append(new Fill(new PatternFill(foreground) { PatternType = PatternValues.Solid }));
                    styles.Fills.Count = format.FillId!.Value + 1U;
                    format.ApplyFill = true;
                } else {
                    format.BorderId = (uint)styles.Borders!.Elements<Border>().Count();
                    styles.Borders.Append(new Border(new LeftBorder(color) { Style = BorderStyleValues.Thin }));
                    styles.Borders.Count = format.BorderId!.Value + 1U;
                    format.ApplyBorder = true;
                }
                uint index = (uint)styles.CellFormats!.Elements<CellFormat>().Count();
                styles.CellFormats.Append(format);
                styles.CellFormats.Count = index + 1U;
                part.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A1").StyleIndex = index;
            }
            styles.Save();
            part.Worksheet.Save();
        }

        private static ColorType ReadCopyStyleColor(SpreadsheetDocument document, string sheetName, string kind) {
            Stylesheet styles = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            Worksheet worksheet = GetWorksheetPartByNameForOperations(document, sheetName).Worksheet;
            if (kind == "dxf") {
                uint index = worksheet.Descendants<ConditionalFormattingRule>().Single().FormatId!.Value;
                return styles.DifferentialFormats!.Elements<DifferentialFormat>().ElementAt((int)index).Descendants<ColorType>().Single();
            }
            Cell cell = worksheet.Descendants<Cell>().Single(item => item.CellReference?.Value == "A1");
            CellFormat format = styles.CellFormats!.Elements<CellFormat>().ElementAt((int)(cell.StyleIndex?.Value ?? 0U));
            OpenXmlElement entry = kind switch {
                "font" => styles.Fonts!.Elements<Font>().ElementAt((int)(format.FontId?.Value ?? 0U)),
                "fill" => styles.Fills!.Elements<Fill>().ElementAt((int)(format.FillId?.Value ?? 0U)),
                _ => styles.Borders!.Elements<Border>().ElementAt((int)(format.BorderId?.Value ?? 0U))
            };
            return entry.Descendants<ColorType>().Single();
        }

        private static void CreateTableSource(string path, string sheetName, string tableName) {
            using var document = ExcelDocument.Create(path);
            ExcelSheet sheet = document.AddSheet(sheetName);
            sheet.Cell(1, 1).SetValue("Region");
            sheet.Cell(1, 2).SetValue("Revenue");
            sheet.Cell(2, 1).SetValue("North");
            sheet.Cell(2, 2).SetValue(10);
            sheet.Cell(3, 1).SetValue("South");
            sheet.Cell(3, 2).SetValue(20);
            sheet.AddTable("A1:B3", hasHeader: true, name: tableName);
            document.Save();
        }

        /// <summary>
        /// Gives the sheet's table every kind of dxf reference and adds two conditional formats, all pointing at dxfs
        /// appended to the source stylesheet.
        /// </summary>
        private static void AddTableAndConditionalFormatDxfs(string path, string sheetName) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(path, true);
            Stylesheet stylesheet = spreadsheet.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            stylesheet.DifferentialFormats ??= new DifferentialFormats();
            DifferentialFormats dxfs = stylesheet.DifferentialFormats;
            uint first = (uint)dxfs.Elements<DifferentialFormat>().Count();
            dxfs.Append(new DifferentialFormat(new Font(new Bold(), new Color { Rgb = "FF110000" })));
            dxfs.Append(new DifferentialFormat(new Fill(new PatternFill(new BackgroundColor { Rgb = "FF00AA00" }) { PatternType = PatternValues.Solid })));
            dxfs.Append(new DifferentialFormat(new Border(new LeftBorder(new Color { Rgb = "FF0000CC" }) { Style = BorderStyleValues.Thin })));
            dxfs.Append(new DifferentialFormat(new Font(new Italic(), new Color { Rgb = "FF333333" })));
            dxfs.Count = (uint)dxfs.Elements<DifferentialFormat>().Count();
            stylesheet.Save();

            WorksheetPart worksheetPart = GetWorksheetPartByNameForOperations(spreadsheet, sheetName);
            Table table = worksheetPart.TableDefinitionParts.Single().Table!;
            table.HeaderRowFormatId = first;
            table.DataFormatId = first + 1U;
            table.TotalsRowFormatId = first + 3U;
            table.HeaderRowBorderFormatId = first + 2U;
            table.BorderFormatId = first + 2U;
            table.TotalsRowBorderFormatId = first + 2U;
            TableColumn column = table.TableColumns!.Elements<TableColumn>().First();
            column.HeaderRowDifferentialFormattingId = first + 3U;
            column.DataFormatId = first;
            column.TotalsRowDifferentialFormattingId = first + 1U;
            AutoFilter autoFilter = table.AutoFilter!;
            autoFilter.RemoveAllChildren<FilterColumn>();
            autoFilter.Append(new FilterColumn(new ColorFilter { FormatId = first + 1U }) { ColumnId = 1U });
            table.Save();

            Worksheet worksheet = worksheetPart.Worksheet;
            ConditionalFormatting conditionalFormatting = CreateExpressionRule("B2:B3", "B2>10", first, 1);
            conditionalFormatting.Append(new ConditionalFormattingRule(new Formula("B2>15")) {
                Type = ConditionalFormatValues.Expression,
                FormatId = first + 1U,
                Priority = 2
            });
            worksheet.InsertAfter(conditionalFormatting, worksheet.GetFirstChild<SheetData>());
            worksheet.Save();
        }

        private static ConditionalFormatting CreateExpressionRule(string range, string formula, uint dxfId, int priority) =>
            new ConditionalFormatting(new ConditionalFormattingRule(new Formula(formula)) {
                Type = ConditionalFormatValues.Expression,
                FormatId = dxfId,
                Priority = priority
            }) {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = range }
            };

        private static List<(string Name, uint Id)> ReadTableDifferentialFormatIds(Table table) {
            var ids = new List<(string Name, UInt32Value? Id)> {
                ("headerRowDxfId", table.HeaderRowFormatId),
                ("dataDxfId", table.DataFormatId),
                ("totalsRowDxfId", table.TotalsRowFormatId),
                ("headerRowBorderDxfId", table.HeaderRowBorderFormatId),
                ("tableBorderDxfId", table.BorderFormatId),
                ("totalsRowBorderDxfId", table.TotalsRowBorderFormatId)
            };
            foreach (TableColumn column in table.TableColumns!.Elements<TableColumn>()) {
                ids.Add((column.Name + ".headerRowDxfId", column.HeaderRowDifferentialFormattingId));
                ids.Add((column.Name + ".dataDxfId", column.DataFormatId));
                ids.Add((column.Name + ".totalsRowDxfId", column.TotalsRowDifferentialFormattingId));
            }

            foreach (ColorFilter filter in table.Descendants<ColorFilter>()) {
                ids.Add(("colorFilter", filter.FormatId));
            }

            return ids.Where(_ => _.Id != null).Select(_ => (_.Name, _.Id!.Value)).ToList();
        }

        private static List<DifferentialFormat> ReadDifferentialFormats(SpreadsheetDocument document) =>
            document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.DifferentialFormats!.Elements<DifferentialFormat>().ToList();

        private static void AssertSameDifferentialFormat(
            List<DifferentialFormat> sourceDxfs,
            uint sourceId,
            List<DifferentialFormat> targetDxfs,
            uint targetId,
            string reference) {
            Assert.True(targetId < targetDxfs.Count, $"{reference} points at dxf {targetId}, but the target has {targetDxfs.Count}.");
            Assert.Equal(CanonicalStyleXml(sourceDxfs[(int)sourceId]), CanonicalStyleXml(targetDxfs[(int)targetId]));
        }

        /// <summary>Describes how a cell renders: its font, fill, border and number format code, and its alignment.</summary>
        private static string DescribeCellFormat(SpreadsheetDocument document, Cell cell) {
            Stylesheet stylesheet = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            CellFormat format = stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)(cell.StyleIndex?.Value ?? 0U));
            uint numberFormatId = format.NumberFormatId?.Value ?? 0U;
            string numberFormat = stylesheet.NumberingFormats?.Elements<NumberingFormat>()
                .FirstOrDefault(_ => _.NumberFormatId?.Value == numberFormatId)?.FormatCode?.Value
                ?? "builtin:" + numberFormatId;
            return string.Join(" | ",
                CanonicalStyleXml(stylesheet.Fonts!.Elements<Font>().ElementAt((int)(format.FontId?.Value ?? 0U))),
                CanonicalStyleXml(stylesheet.Fills!.Elements<Fill>().ElementAt((int)(format.FillId?.Value ?? 0U))),
                CanonicalStyleXml(stylesheet.Borders!.Elements<Border>().ElementAt((int)(format.BorderId?.Value ?? 0U))),
                numberFormat,
                format.Alignment == null ? string.Empty : CanonicalStyleXml(format.Alignment));
        }

        private static string CanonicalStyleXml(OpenXmlElement element) => CanonicalStyleXml(XElement.Parse(element.OuterXml));

        private static string CanonicalStyleXml(XElement element) {
            IEnumerable<string> attributes = element.Attributes()
                .Where(_ => !_.IsNamespaceDeclaration)
                .OrderBy(_ => _.Name.ToString(), System.StringComparer.Ordinal)
                .Select(_ => " " + _.Name + "=\"" + _.Value + "\"");
            return "<" + element.Name + string.Concat(attributes) + ">" + string.Concat(element.Elements().Select(CanonicalStyleXml)) + "</>";
        }

        private static string ReadStyleCounts(string path) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(path, false);
            Stylesheet stylesheet = spreadsheet.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            return string.Join(", ",
                "numFmts " + (stylesheet.NumberingFormats?.Elements<NumberingFormat>().Count() ?? 0),
                "fonts " + stylesheet.Fonts!.Elements<Font>().Count(),
                "fills " + stylesheet.Fills!.Elements<Fill>().Count(),
                "borders " + stylesheet.Borders!.Elements<Border>().Count(),
                "cellStyleXfs " + stylesheet.CellStyleFormats!.Elements<CellFormat>().Count(),
                "cellXfs " + stylesheet.CellFormats!.Elements<CellFormat>().Count(),
                "dxfs " + stylesheet.DifferentialFormats!.Elements<DifferentialFormat>().Count());
        }
    }
}
