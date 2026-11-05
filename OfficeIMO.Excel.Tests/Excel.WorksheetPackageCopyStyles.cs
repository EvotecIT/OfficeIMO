using System.Collections.Generic;
using System.IO;
using System.Linq;
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

        private static string CanonicalStyleXml(OpenXmlElement element) => CanonicalStyleXml(XElement.Parse(element.OuterXml));

        private static string CanonicalStyleXml(XElement element) {
            IEnumerable<string> attributes = element.Attributes()
                .Where(_ => !_.IsNamespaceDeclaration)
                .OrderBy(_ => _.Name.ToString(), System.StringComparer.Ordinal)
                .Select(_ => " " + _.Name + "=\"" + _.Value + "\"");
            return "<" + element.Name + string.Concat(attributes) + ">" + string.Concat(element.Elements().Select(CanonicalStyleXml)) + "</>";
        }
    }
}
