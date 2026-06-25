using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ExcelWorkbookIntelligence_ReportsFormulasDoctorDiffAndCompliance() {
            string leftPath = Path.Combine(_directoryWithFiles, "ExcelWorkbookIntelligence.Left.xlsx");
            string rightPath = Path.Combine(_directoryWithFiles, "ExcelWorkbookIntelligence.Right.xlsx");

            using (var document = ExcelDocument.Create(leftPath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, 10);
                sheet.CellValue(1, 2, 20);
                sheet.CellFormula(1, 3, "SUM(A1:B1)+NOW()");
                sheet.MergeRange("A3:B3");
                document.SetNamedRange("TotalValue", "'Data'!C1", save: false);
                document.Save();
            }

            using (var document = ExcelDocument.Create(rightPath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, 11);
                sheet.CellValue(1, 2, 20);
                sheet.CellFormula(1, 3, "SUM(A1:B1)");
                document.Save();
            }

            using (var left = ExcelDocument.Load(leftPath, readOnly: false))
            using (var right = ExcelDocument.Load(rightPath, readOnly: true)) {
                ExcelFormulaAnalysisReport formulas = left.AnalyzeFormulas();
                Assert.Equal(1, formulas.FormulaCount);
                Assert.Equal(1, formulas.VolatileFormulaCount);
                Assert.Contains("SUM", formulas.Formulas[0].Functions);

                Assert.Contains(left.ListNamedRanges(), name => name.Name == "TotalValue");
                left.RenameNamedRange("TotalValue", "GrandTotal", save: false);
                Assert.Contains(left.ListNamedRanges(), name => name.Name == "GrandTotal");

                ExcelWorkbookDiagnosticReport doctor = left.RunWorkbookDoctor(new ExcelWorkbookDoctorOptions { ValidateOpenXml = false });
                Assert.False(doctor.HasErrors);
                Assert.DoesNotContain(doctor.Issues, issue => issue.Severity == ExcelFindingSeverity.Error);

                ExcelWorkbookDiffReport diff = left.CompareWorkbook(right);
                Assert.False(diff.AreEqual);
                Assert.Contains(diff.Differences, item => item.Category == "Cell" && item.Address == "A1");

                ExcelAccessibilityReport accessibility = left.AnalyzeAccessibility();
                Assert.Contains(accessibility.Findings, finding => finding.Category == "MergedCells");

                ExcelDataModelReport dataModel = left.InspectDataModel();
                Assert.False(dataModel.HasDataModelOrQueries);

                ExcelStreamingContractReport streaming = left.GetStreamingContract();
                Assert.Equal(1, streaming.WorksheetCount);
            }
        }

        [Fact]
        public void Test_ExcelDelimitedImport_NormalizesCsvIntoTable() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelDelimitedImport.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelDelimitedImportResult result = document.ImportDelimitedText(
                    "Name;Amount\r\nAlpha;10.5\r\nBeta;11.75",
                    new ExcelDelimitedImportOptions { Delimiter = ';', SheetName = "Import", TableName = "ImportData" });

                Assert.Equal(';', result.Delimiter);
                Assert.Equal("A1:B3", result.ImportResult.Range);
                Assert.Equal("ImportData", result.ImportResult.TableName);
                Assert.Equal(2, result.ImportResult.RowCount);
                Assert.Equal(2, result.ImportResult.ColumnCount);
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.True(document["Import"].TryGetCellText(2, 1, out string name));
                Assert.Equal("Alpha", name);
            }
        }

        [Fact]
        public void Test_ExcelDelimitedImport_StreamsDelimitedFileIntoTable() {
            string sourcePath = Path.Combine(_directoryWithFiles, "ExcelDelimitedImport.Source.csv");
            string filePath = Path.Combine(_directoryWithFiles, "ExcelDelimitedImport.File.xlsx");
            File.WriteAllText(sourcePath, "Name,Amount\r\nAlpha,10.5\r\nBeta,11.75");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelDelimitedImportResult result = document.ImportDelimitedFile(
                    sourcePath,
                    new ExcelDelimitedImportOptions { SheetName = "Import", TableName = "ImportData" });

                Assert.Equal(',', result.Delimiter);
                Assert.Equal("A1:B3", result.ImportResult.Range);
                Assert.Equal("ImportData", result.ImportResult.TableName);
                Assert.Equal(2, result.ImportResult.RowCount);
                Assert.Equal(2, result.ImportResult.ColumnCount);
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.True(document["Import"].TryGetCellText(3, 1, out string name));
                Assert.Equal("Beta", name);
            }
        }

        [Fact]
        public void Test_ExcelDelimitedImport_SkipsInitialRecordsBeforeHeader() {
            using var document = ExcelDocument.Create(new MemoryStream(), autoSave: false);

            ExcelDelimitedImportResult result = document.ImportDelimitedText(
                "generated by vendor\r\nName,Amount\r\nAlpha,10.5",
                new ExcelDelimitedImportOptions {
                    SheetName = "Import",
                    SkipInitialRecords = 1
                });

            Assert.Equal("A1:B2", result.ImportResult.Range);
            Assert.Equal(1, result.ImportResult.RowCount);
            ExcelSheet sheet = document["Import"];
            Assert.True(sheet.TryGetCellText(1, 1, out string header));
            Assert.Equal("Name", header);
            Assert.True(sheet.TryGetCellText(2, 1, out string name));
            Assert.Equal("Alpha", name);
        }

        [Fact]
        public void Test_ExcelDelimitedImport_PreservesFieldsBeyondHeader() {
            using var document = ExcelDocument.Create(new MemoryStream(), autoSave: false);

            ExcelDelimitedImportResult result = document.ImportDelimitedText(
                "Name\r\nAlpha,10.5\r\nBeta,11.75",
                new ExcelDelimitedImportOptions { SheetName = "Import" });

            Assert.Equal("A1:B3", result.ImportResult.Range);
            Assert.Equal(2, result.ImportResult.ColumnCount);
            ExcelSheet sheet = document["Import"];
            Assert.True(sheet.TryGetCellText(1, 2, out string header));
            Assert.Equal("Column2", header);
            Assert.True(sheet.TryGetCellText(2, 2, out string amount));
            Assert.Equal("10.5", amount);
        }

        [Fact]
        public void Test_ExcelWorkbookIntelligence_AdvancedReportsRepairAndDiffDepth() {
            string leftPath = Path.Combine(_directoryWithFiles, "ExcelWorkbookIntelligence.Advanced.Left.xlsx");
            string rightPath = Path.Combine(_directoryWithFiles, "ExcelWorkbookIntelligence.Advanced.Right.xlsx");

            using (var document = ExcelDocument.Create(leftPath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "{{CustomerName}}");
                sheet.CellValue(1, 2, 10);
                sheet.CellFormula(1, 3, "SUM(B1:B1)");
                sheet.SetComment("A1", "Needs customer binding", author: "Reviewer", initials: "RV");
                sheet.AddTable("A1:C1", hasHeader: true, name: "DataTable", TableStyle.TableStyleMedium2, includeAutoFilter: true);
                document.SetNamedRange("CustomerCell", "'Data'!A1", save: false);
                document.AddPowerQueryMetadata(new ExcelPowerQueryMetadataOptions {
                    Name = "CustomerQuery",
                    WorksheetName = "Data",
                    CommandText = "let Source = Excel.CurrentWorkbook() in Source",
                    RefreshOnOpen = true
                });
                document.AddWorkbookConnectionMetadata("<?xml version=\"1.0\" encoding=\"UTF-8\"?><connections xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><connection id=\"77\" name=\"Existing\" refreshOnLoad=\"0\" type=\"1\"/></connections>");
                ExcelPowerQueryMetadataResult isolatedRefresh = document.AddPowerQueryMetadata(new ExcelPowerQueryMetadataOptions {
                    Name = "IsolatedRefresh",
                    CommandText = "let Source = Excel.CurrentWorkbook() in Source",
                    RefreshOnOpen = true
                });
                Assert.Equal(78U, isolatedRefresh.ConnectionId);
                document.Save();
            }

            using (var document = ExcelDocument.Create(rightPath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Northwind");
                sheet.CellValue(1, 2, 10);
                sheet.CellFormula(1, 3, "SUM(B1:B1)");
                sheet.AddTable("A1:C1", hasHeader: true, name: "OtherTable", TableStyle.TableStyleMedium2, includeAutoFilter: true);
                document.Save();
            }

            using (var left = ExcelDocument.Load(leftPath, readOnly: false))
            using (var right = ExcelDocument.Load(rightPath, readOnly: true)) {
                ExcelWorkbookCommentReport comments = left.InspectComments();
                Assert.Equal(1, comments.CommentCount);
                Assert.Empty(comments.Issues);

                ExcelTemplateBindingReport missingTemplate = left.ValidateTemplateBindings(new Dictionary<string, object?>());
                Assert.False(missingTemplate.Passed);
                Assert.Contains("CustomerName", missingTemplate.MissingMarkerNames);
                Assert.Contains("Excel Template Markers", missingTemplate.Markdown);

                ExcelTemplateBindingReport boundTemplate = left.ValidateTemplateBindings(new Dictionary<string, object?> {
                    ["CustomerName"] = "Northwind"
                });
                Assert.True(boundTemplate.Passed);

                ExcelDataModelReport dataModel = left.InspectDataModel();
                Assert.True(dataModel.HasDataModelOrQueries);
                Assert.True(dataModel.ConnectionPartCount > 0);
                Assert.True(dataModel.QueryTablePartCount > 0);
                using (SpreadsheetDocument package = SpreadsheetDocument.Open(leftPath, false)) {
                    string existingConnection = EnumerateWorkbookConnectionXml(package.WorkbookPart!)
                        .First(text => text.Contains("Existing", System.StringComparison.Ordinal));
                    Assert.Contains("refreshOnLoad=\"0\"", existingConnection);
                }

                ExcelWorkbookDiffReport diff = left.CompareWorkbook(right, new ExcelWorkbookDiffOptions {
                    CompareComments = true,
                    CompareNamedRanges = true,
                    CompareTables = true,
                    CompareWorksheetMetadata = true,
                    MaxDifferences = 20
                });
                Assert.False(diff.AreEqual);
                Assert.Contains(diff.Differences, item => item.Category == "NamedRange");
                Assert.Contains(diff.Differences, item => item.Category == "Table");
                Assert.Contains(diff.Differences, item => item.Category == "Comment");

                ExcelWorkbookRepairReport repair = left.RepairWorkbook(new ExcelWorkbookRepairOptions { Save = false });
                Assert.True(repair.ActionCount > 0);
                Assert.NotNull(repair.Before);
                Assert.NotNull(repair.After);
            }
        }

        [Fact]
        public void Test_ExcelWorkbookIntelligence_NamedRangeSavePersistsPackage() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelWorkbookIntelligence.NamedRangeSave.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Value");
                document.SetNamedRange("Totals", "'Data'!A1:B1", scope: sheet, save: true);
            }

            using (var document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document["Data"];
                Assert.True(document.RenameNamedRange("Totals", "GrandTotal", sheet, save: true));
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                ExcelSheet sheet = document["Data"];
                Assert.Null(document.GetNamedRange("Totals", sheet));
                Assert.Equal("$A$1:$B$1", document.GetNamedRange("GrandTotal", sheet));
            }
        }

        [Fact]
        public void Test_ExcelWorkbookIntelligence_RepairSaveSkipsMissingDefaultDestination() {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream, autoSave: true)) {
                document.AddWorkSheet("Data").CellValue(1, 1, "Ready");
            }

            stream.Position = 0;
            using (var document = ExcelDocument.Load(stream, readOnly: false, autoSave: false)) {
                ExcelWorkbookRepairReport repair = document.RepairWorkbook(new ExcelWorkbookRepairOptions { Save = true });
                Assert.NotNull(repair.Before);
                Assert.NotNull(repair.After);
            }
        }

        [Fact]
        public void Test_ExcelWorkbookIntelligence_PowerQueryRejectsMissingWorksheetBeforeConnectionMutation() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelWorkbookIntelligence.QueryMissingSheet.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorkSheet("Data");

                Assert.Throws<ArgumentOutOfRangeException>(() => document.AddPowerQueryMetadata(new ExcelPowerQueryMetadataOptions {
                    Name = "MissingSheetQuery",
                    WorksheetName = "Missing",
                    CommandText = "let Source = Excel.CurrentWorkbook() in Source"
                }));

                Assert.False(document.InspectDataModel().HasDataModelOrQueries);
            }
        }


        [Fact]
        public void Test_ExcelWorkbookDiff_ReportsRightOnlyCellsWithinSameDimension() {
            string leftPath = Path.Combine(_directoryWithFiles, "ExcelWorkbookDiff.LeftSparse.xlsx");
            string rightPath = Path.Combine(_directoryWithFiles, "ExcelWorkbookDiff.RightSparse.xlsx");

            using (var document = ExcelDocument.Create(leftPath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "A");
                sheet.CellValue(1, 3, "C");
                document.Save();
            }

            using (var document = ExcelDocument.Create(rightPath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "A");
                sheet.CellValue(1, 2, "B");
                sheet.CellValue(1, 3, "C");
                document.Save();
            }

            using (var left = ExcelDocument.Load(leftPath, readOnly: true))
            using (var right = ExcelDocument.Load(rightPath, readOnly: true)) {
                ExcelWorkbookDiffReport diff = left.CompareWorkbook(right);
                Assert.Contains(diff.Differences, difference => difference.Category == "Cell" && difference.Address == "B1" && difference.RightValue == "B");
            }
        }

        [Fact]
        public void Test_ExcelWorkbookIntelligence_ThreadedCommentsAndQueryMetadataContracts() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelWorkbookIntelligence.ThreadedAndQuery.xlsx");

            string parentId;
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Review");
                sheet.CellValue(1, 1, "Variance");
                sheet.CellValue(1, 2, 123.45);

                ExcelThreadedCommentResult parent = sheet.AddThreadedComment(new ExcelThreadedCommentOptions {
                    Address = "B1",
                    Text = "Please confirm the variance.",
                    Author = "Finance Reviewer",
                    Id = "{00000000-0000-0000-0000-000000000101}",
                    Date = new DateTime(2026, 6, 22, 8, 0, 0, DateTimeKind.Utc)
                });
                parentId = parent.Id;

                ExcelThreadedCommentResult reply = sheet.AddThreadedComment(new ExcelThreadedCommentOptions {
                    Address = "B1",
                    Text = "Confirmed against source data.",
                    Author = "Report Owner",
                    ParentId = parentId,
                    Id = "{00000000-0000-0000-0000-000000000102}",
                    Done = true
                });

                ExcelPowerQueryMetadataResult first = document.AddPowerQueryMetadata(new ExcelPowerQueryMetadataOptions {
                    Name = "ReviewQuery",
                    WorksheetName = "Review",
                    QueryTableName = "ReviewQueryTable",
                    CommandText = "let Source = Excel.CurrentWorkbook() in Source"
                });
                ExcelPowerQueryMetadataResult second = document.AddPowerQueryMetadata(new ExcelPowerQueryMetadataOptions {
                    Name = "ReviewQueryRefresh",
                    WorksheetName = "Review",
                    CommandText = "let Source = Excel.CurrentWorkbook() in Source",
                    RefreshOnOpen = true
                });

                Assert.Equal("B1", parent.CellReference);
                Assert.False(parent.IsReply);
                Assert.True(reply.IsReply);
                Assert.True(reply.Done);
                Assert.Equal(1U, first.ConnectionId);
                Assert.Equal(2U, second.ConnectionId);
                Assert.Equal("ReviewQueryTable", first.QueryTableName);
                Assert.Equal("ReviewQueryRefreshTable", second.QueryTableName);
                Assert.Throws<ArgumentException>(() => sheet.AddThreadedComment("XFE1", "Outside sheet"));
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                ExcelWorkbookCommentReport comments = document.InspectComments();
                Assert.Equal(2, comments.ThreadedCommentCount);
                Assert.Contains(comments.ThreadedComments, comment => comment.Id == parentId && comment.Author == "Finance Reviewer");
                Assert.Contains(comments.ThreadedComments, comment => comment.ParentId == parentId && comment.Author == "Report Owner" && comment.Done);
                Assert.Empty(comments.Issues);

                ExcelDataModelReport dataModel = document.InspectDataModel();
                Assert.True(dataModel.HasDataModelOrQueries);
                Assert.Equal(1, dataModel.ConnectionPartCount);
                Assert.Equal(2, dataModel.QueryTablePartCount);
            }
        }

        private static string ReadExtendedPartText(ExtendedPart part) {
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using var reader = new StreamReader(stream);
            return reader.ReadToEnd();
        }

        private static IEnumerable<string> EnumerateWorkbookConnectionXml(OpenXmlPartContainer container) {
            foreach (IdPartPair pair in container.Parts) {
                if (pair.OpenXmlPart is ConnectionsPart connectionsPart && connectionsPart.Connections != null) {
                    yield return connectionsPart.Connections.OuterXml;
                } else if (pair.OpenXmlPart is ExtendedPart extendedPart
                    && extendedPart.ContentType.IndexOf("connections", System.StringComparison.OrdinalIgnoreCase) >= 0) {
                    yield return ReadExtendedPartText(extendedPart);
                }

                foreach (string child in EnumerateWorkbookConnectionXml(pair.OpenXmlPart)) {
                    yield return child;
                }
            }
        }

        [Fact]
        public void Test_ExcelWorkbookDiff_ReportsRightOnlyCellStyles() {
            string leftPath = Path.Combine(_directoryWithFiles, "ExcelWorkbookDiff.LeftStyle.xlsx");
            string rightPath = Path.Combine(_directoryWithFiles, "ExcelWorkbookDiff.RightStyle.xlsx");

            using (var left = ExcelDocument.Create(leftPath)) {
                left.AddWorkSheet("Data");
                left.Save();
            }

            using (var right = ExcelDocument.Create(rightPath)) {
                var sheet = right.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Styled");
                sheet.CellBold(1, 1, true);
                right.Save();
            }

            using (var left = ExcelDocument.Load(leftPath, readOnly: true))
            using (var right = ExcelDocument.Load(rightPath, readOnly: true)) {
                ExcelWorkbookDiffReport report = left.CompareWorkbook(right, new ExcelWorkbookDiffOptions {
                    CompareCells = false,
                    CompareCellStyles = true,
                    CompareWorksheetMetadata = false,
                    CompareTables = false,
                    CompareNamedRanges = false,
                    CompareComments = false
                });

                Assert.Contains(report.Differences, difference =>
                    difference.Category == "CellStyle"
                    && difference.SheetName == "Data"
                    && difference.Address == "A1"
                    && string.IsNullOrEmpty(difference.LeftValue)
                    && !string.IsNullOrEmpty(difference.RightValue));
            }
        }
    }
}
