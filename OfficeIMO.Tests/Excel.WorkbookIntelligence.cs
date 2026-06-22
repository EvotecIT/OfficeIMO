using System.Collections.Generic;
using System.IO;
using System.Linq;
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
                    new ExcelDelimitedImportOptions { Delimiter = ';', SheetName = "Import" });

                Assert.Equal(';', result.Delimiter);
                Assert.Equal("A1:B3", result.ImportResult.Range);
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
                Assert.Equal(2, dataModel.ConnectionPartCount);
                Assert.Equal(2, dataModel.QueryTablePartCount);
            }
        }
    }
}
