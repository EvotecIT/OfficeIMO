using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Security;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void FormulaFunctionSearch_RemainsLinearForMalformedQuotedAndBracketedTokens() {
            string[] hostileFormulas = {
                new string('\'', 40_001) + "+SUM(A1)",
                new string('[', 40_000) + "SUM(A1)",
                string.Concat(Enumerable.Repeat("LET(", 10_000)) + new string('\'', 40_001) + "+SUM(A1)",
                string.Concat(Enumerable.Repeat("LAMBDA(", 10_000)) + new string('[', 40_000) + "SUM(A1)"
            };
            var stopwatch = Stopwatch.StartNew();
            foreach (string formula in hostileFormulas) {
                var cell = new ExcelFormulaCellInfo(
                    "Data",
                    "A1",
                    formula,
                    cachedValue: null,
                    isDirty: false,
                    isSupportedByOfficeIMO: false,
                    unsupportedReason: "Malformed security fixture");

                Assert.Empty(ExcelSheet.SearchFormulaCells(
                    new[] { cell },
                    new ExcelFormulaSearchOptions { Function = "SUM" },
                    Array.Empty<string>()));
            }
            stopwatch.Stop();

            Assert.True(
                stopwatch.Elapsed < TimeSpan.FromSeconds(3),
                $"Malformed formula scans exceeded the linear-time budget: {stopwatch.Elapsed}.");

            var legitimate = new ExcelFormulaCellInfo(
                "Data",
                "A1",
                "'Sales 2026'!SUM(A1)+Table1['#Data]",
                cachedValue: null,
                isDirty: false,
                isSupportedByOfficeIMO: false,
                unsupportedReason: "Search-only fixture");
            Assert.Single(ExcelSheet.SearchFormulaCells(
                new[] { legitimate },
                new ExcelFormulaSearchOptions { Function = "SUM" },
                Array.Empty<string>()));
        }

        [Fact]
        public async Task QueryBackedTable_TracksExactAuthoredPartAcrossConnectionIdCollision() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet importedSheet = document.AddWorksheet("Imported");
            importedSheet.CellValue(1, 1, "Value");
            importedSheet.AddTable(
                "A1:A1",
                hasHeader: true,
                name: "ImportedResults",
                style: OfficeIMO.Excel.ExcelTableStyle.TableStyleMedium2);
            TableDefinitionPart importedTablePart = Assert.Single(
                importedSheet.WorksheetPart.TableDefinitionParts);
            QueryTablePart importedQueryPart = importedTablePart.AddNewPart<QueryTablePart>();
            importedQueryPart.QueryTable = new QueryTable {
                Name = "ImportedResults",
                ConnectionId = 1U
            };
            importedQueryPart.QueryTable.Save();

            ExcelSheet authoredSheet = document.AddWorksheet("Authored");
            ExcelQueryBackedTableInfo authored = document.AddQueryBackedTable(
                new ExcelQueryBackedTableOptions {
                    ConnectionName = "TrustedQuery",
                    WorksheetName = authoredSheet.Name,
                    TableName = "TrustedResults",
                    ColumnNames = new[] { "Value" }
                });
            Assert.Equal(1U, authored.ConnectionId);

            ExcelQueryBackedTableInfo imported = document.GetQueryBackedTables()
                .Single(item => item.TableName == "ImportedResults");
            Assert.True(imported.IsImported);
            Assert.False(document.GetQueryBackedTables()
                .Single(item => item.TableName == authored.TableName).IsImported);

            var host = new StubQueryHost(new ExcelQueryExecutionResult(
                new[] { "Value" },
                new IReadOnlyList<object?>[] { new object?[] { "unsafe" } }));
            await Assert.ThrowsAsync<SecurityException>(() => document.RefreshQueryAsync(
                imported.TableName,
                host,
                new ExcelQueryExecutionPolicy { AllowExecution = true }));
            Assert.Equal(0, host.CallCount);
        }

        [Fact]
        public void PackageWorksheetCopy_EnforcesInCellImageCountAndByteBudgets() {
            using var sourceDocument = ExcelDocument.Create(new MemoryStream());
            ExcelSheet source = sourceDocument.AddWorksheet("Images");
            source.SetInCellImage(1, 1, TinyPng, altText: "Shared A");
            source.CellValue(1, 2, "placeholder");
            Cell first = source.WorksheetPart.Worksheet!.Descendants<Cell>()
                .Single(cell => cell.CellReference?.Value == "A1");
            Cell second = source.WorksheetPart.Worksheet.Descendants<Cell>()
                .Single(cell => cell.CellReference?.Value == "B1");
            second.CellValue = (CellValue?)first.CellValue?.CloneNode(true);
            second.DataType = first.DataType?.Value;
            second.ValueMetaIndex = first.ValueMetaIndex?.Value;
            second.InlineString = null;
            source.WorksheetPart.Worksheet.Save();

            AssertPackageImageCopyRejected(
                sourceDocument,
                new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package,
                    MaxInCellImages = 1
                },
                "in-cell image limit");
            AssertPackageImageCopyRejected(
                sourceDocument,
                new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package,
                    MaxTotalInCellImageBytes = TinyPng.LongLength
                },
                "aggregate in-cell image limit");
            AssertPackageImageCopyRejected(
                sourceDocument,
                new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package,
                    MaxInCellImageBytes = TinyPng.LongLength - 1L
                },
                "In-cell image payloads");

            var sameWorkbookOptions = new ExcelWorksheetCopyOptions {
                CopyMode = ExcelWorksheetCopyMode.Package,
                MaxInCellImages = 1
            };
            Assert.Throws<InvalidOperationException>(() => sourceDocument.CopyWorksheetFrom(
                sourceDocument,
                "Images",
                "RejectedCopy",
                ExcelSheetNameValidationMode.Sanitize,
                sameWorkbookOptions));
            Assert.Single(sourceDocument.Sheets);

            Assert.Throws<InvalidOperationException>(() => sourceDocument.MergeWorkbookFrom(
                sourceDocument,
                new ExcelWorkbookMergeOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package,
                    SheetNames = new[] { "Images" },
                    SheetNamePrefix = "Rejected ",
                    MaxInCellImages = 1
                }));
            Assert.Single(sourceDocument.Sheets);
        }

        [Fact]
        public void WorkbookMerge_ReusesSharedInCellImagePartsAcrossSheets() {
            using var sourceDocument = ExcelDocument.Create(new MemoryStream());
            ExcelSheet source = sourceDocument.AddWorksheet("First");
            source.SetInCellImage(1, 1, TinyPng, altText: "Shared");
            sourceDocument.CopyWorksheet("First", "Second");

            using var targetDocument = ExcelDocument.Create(new MemoryStream());
            targetDocument.AddWorksheet("Existing");
            targetDocument.MergeWorkbookFrom(sourceDocument, new ExcelWorkbookMergeOptions {
                CopyMode = ExcelWorksheetCopyMode.Package,
                SheetNames = new[] { "First", "Second" }
            });

            ExtendedPart relationships = targetDocument.WorkbookPartRoot.Parts
                .Select(pair => pair.OpenXmlPart)
                .OfType<ExtendedPart>()
                .Single(part => part.RelationshipType.EndsWith("/richValueRel", StringComparison.Ordinal));
            Assert.Single(relationships.Parts);
            Assert.Single(targetDocument.GetSheet("First").GetInCellImages());
            Assert.Single(targetDocument.GetSheet("Second").GetInCellImages());
        }

        [Fact]
        public void CellSmartTagMutation_DoesNotRemoveAnUnrelatedParent() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(2, 1, "delete");
            const string extensionNamespace = "urn:officeimo:untrusted-extension";
            var unrelatedParent = new OpenXmlUnknownElement(
                string.Empty,
                "cellSmartTags",
                extensionNamespace);
            var preservedChild = new OpenXmlUnknownElement(
                string.Empty,
                "preserved",
                extensionNamespace);
            var fakeSmartTag = new OpenXmlUnknownElement(
                string.Empty,
                "cellSmartTag",
                extensionNamespace);
            fakeSmartTag.SetAttribute(new OpenXmlAttribute(
                string.Empty,
                "r",
                string.Empty,
                "A2"));
            unrelatedParent.Append(preservedChild, fakeSmartTag);
            sheet.WorksheetPart.Worksheet.Append(unrelatedParent);

            sheet.DeleteCells("A2", ExcelCellShiftDirection.Up);

            Assert.Same(sheet.WorksheetPart.Worksheet, unrelatedParent.Parent);
            Assert.Contains(preservedChild, unrelatedParent.ChildElements);
            Assert.Contains(fakeSmartTag, unrelatedParent.ChildElements);
        }

        private static void AssertPackageImageCopyRejected(
            ExcelDocument sourceDocument,
            ExcelWorksheetCopyOptions options,
            string expectedMessage) {
            using var targetDocument = ExcelDocument.Create(new MemoryStream());
            targetDocument.AddWorksheet("Existing");

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                targetDocument.CopyWorksheetFrom(
                    sourceDocument,
                    "Images",
                    "Copied",
                    ExcelSheetNameValidationMode.Sanitize,
                    options));

            Assert.Contains(expectedMessage, exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Single(targetDocument.Sheets);
            Assert.Equal("Existing", targetDocument.Sheets[0].Name);
        }
    }
}
