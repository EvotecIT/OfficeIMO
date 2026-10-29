using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Threaded = DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Test_CopyWorksheetFrom_PackageModeCarriesDrawingsCommentsAndHyperlinks(bool targetHasCommentedSheet) {
            using MemoryStream sourceStream = CreateRelatedPartCopySource();
            using var targetStream = new MemoryStream();

            using (var sourceDocument = ExcelDocument.Load(sourceStream, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Create(targetStream)) {
                if (targetHasCommentedSheet) {
                    ExcelSheet existing = targetDocument.AddWorksheet("Existing");
                    existing.CellValue(1, 1, "Existing");
                    existing.SetComment("A1", "Existing note", "Target Author");
                    existing.AddThreadedComment("B1", "Existing thread", "Target Author");
                }

                var options = new ExcelWorksheetCopyOptions { CopyMode = ExcelWorksheetCopyMode.Package };
                targetDocument.CopyWorksheetFrom(sourceDocument, "Report", "Copy1", ExcelSheetNameValidationMode.Strict, options);
                targetDocument.CopyWorksheetFrom(sourceDocument, "Report", "Copy2", ExcelSheetNameValidationMode.Strict, options);

                AssertCopiedRelatedFeatures(targetDocument, "Copy1");
                AssertCopiedRelatedFeatures(targetDocument, "Copy2");
                Assert.Empty(targetDocument.ValidateOpenXml());
                targetDocument.Save();
            }

            targetStream.Position = 0;
            using (var reopened = ExcelDocument.Load(targetStream, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                AssertCopiedRelatedFeatures(reopened, "Copy1");
                AssertCopiedRelatedFeatures(reopened, "Copy2");
                Assert.Empty(reopened.ValidateOpenXml());
                if (targetHasCommentedSheet) {
                    Assert.Equal("Existing note", Assert.Single(reopened["Existing"].GetComments()).Text);
                    Assert.Equal("Target Author", Assert.Single(reopened["Existing"].GetThreadedComments()).Author);
                }
            }

            targetStream.Position = 0;
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetStream, false)) {
                WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
                List<string> threadedIds = workbookPart.WorksheetParts
                    .SelectMany(_ => _.WorksheetThreadedCommentsParts)
                    .SelectMany(_ => _.ThreadedComments!.Elements<Threaded.ThreadedComment>())
                    .Select(_ => _.Id!.Value!)
                    .ToList();
                Assert.Equal(threadedIds.Count, threadedIds.Distinct().Count());

                List<int> shapeIds = workbookPart.WorksheetParts
                    .SelectMany(_ => _.VmlDrawingParts)
                    .SelectMany(_ => Regex.Matches(ReadPartText(_), @"_x0000_s(\d+)").Cast<Match>())
                    .Select(_ => int.Parse(_.Groups[1].Value))
                    .ToList();
                Assert.Equal(targetHasCommentedSheet ? 3 : 2, shapeIds.Count);
                Assert.Equal(shapeIds.Count, shapeIds.Distinct().Count());

                foreach (string sheetName in new[] { "Copy1", "Copy2" }) {
                    Sheet sheet = workbookPart.Workbook.Sheets!.Elements<Sheet>().Single(_ => _.Name == sheetName);
                    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
                    ChartPart chartPart = Assert.Single(worksheetPart.DrawingsPart!.ChartParts);
                    Assert.All(chartPart.ChartSpace!.Descendants<C.Formula>(), _ => Assert.Matches("^'?" + sheetName + "'?!", _.Text));
                }
            }
        }

        [Fact]
        public void Test_CopyWorksheet_WithinWorkbookCarriesDrawingsCommentsAndHyperlinks() {
            using MemoryStream stream = CreateRelatedPartCopySource();
            using (var document = ExcelDocument.Load(stream)) {
                document.CopyWorksheet("Report", "Copy1");
                document.CopyWorksheet("Report", "Copy2");

                AssertCopiedRelatedFeatures(document, "Report");
                AssertCopiedRelatedFeatures(document, "Copy1");
                AssertCopiedRelatedFeatures(document, "Copy2");
                Assert.Empty(document.ValidateOpenXml());
                document.Save();
            }

            stream.Position = 0;
            using var reopened = ExcelDocument.Load(stream, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
            AssertCopiedRelatedFeatures(reopened, "Report");
            AssertCopiedRelatedFeatures(reopened, "Copy1");
            AssertCopiedRelatedFeatures(reopened, "Copy2");
            Assert.Empty(reopened.ValidateOpenXml());
            List<string> threadedIds = reopened.Sheets
                .SelectMany(_ => _.GetThreadedComments())
                .Select(_ => _.Id!)
                .ToList();
            Assert.Equal(6, threadedIds.Count);
            Assert.Equal(6, threadedIds.Distinct().Count());
        }

        private MemoryStream CreateRelatedPartCopySource() {
            var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                ExcelSheet sheet = document.AddWorksheet("Report");
                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 2, "Sales");
                sheet.CellValue(2, 1, "North");
                sheet.CellValue(2, 2, 10);
                sheet.CellValue(3, 1, "South");
                sheet.CellValue(3, 2, 20);
                sheet.AddImage(5, 1, File.ReadAllBytes(Path.Combine(_directoryWithImages, "EvotecLogo.png")), "image/png", name: "Logo");
                sheet.AddChartFromRange("A1:B3", 5, 4, title: "Sales");
                sheet.SetComment("D1", "Check the totals", "Reviewer");
                ExcelThreadedCommentResult thread = sheet.AddThreadedComment("E1", "Is this final?", "Asker");
                sheet.AddThreadedComment("E1", "Yes", "Answerer", parentId: thread.Id);
                sheet.SetHyperlink("F1", "https://example.com/report", "Report link");
                document.Save();
            }

            stream.Position = 0;
            return stream;
        }

        private static void AssertCopiedRelatedFeatures(ExcelDocument document, string sheetName) {
            ExcelSheet sheet = document[sheetName];
            Assert.Equal("Logo", Assert.Single(sheet.Images).Name);
            Assert.Single(sheet.Charts);

            ExcelCommentInfo comment = Assert.Single(sheet.GetComments());
            Assert.Equal("D1", comment.CellReference);
            Assert.Equal("Check the totals", comment.Text);
            Assert.Equal("Reviewer", comment.Author);

            IReadOnlyList<ExcelThreadedCommentSnapshot> threaded = sheet.GetThreadedComments();
            Assert.Equal(2, threaded.Count);
            ExcelThreadedCommentSnapshot root = Assert.Single(threaded, _ => _.ParentId == null);
            ExcelThreadedCommentSnapshot reply = Assert.Single(threaded, _ => _.ParentId != null);
            Assert.Equal("E1", root.CellReference);
            Assert.Equal("Asker", root.Author);
            Assert.Equal("Is this final?", root.Text);
            Assert.Equal(root.Id, reply.ParentId);
            Assert.Equal("Answerer", reply.Author);

            ExcelHyperlinkSnapshot hyperlink = sheet.GetHyperlinks()["F1"];
            Assert.True(hyperlink.IsExternal);
            Assert.Equal("https://example.com/report", hyperlink.Target);
        }
    }
}
