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

        [Fact]
        public void Test_CopyWorksheetFrom_PackageModeDropsSlicerFramesAndRemapsCompatShapeIds() {
            using MemoryStream sourceStream = CreateRelatedPartCopySource();
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(sourceStream, true)) {
                DrawingsPart drawingsPart = spreadsheet.WorkbookPart!.WorksheetParts.Single().DrawingsPart!;
                string anchors = "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">"
                    + "<mc:Choice xmlns:sle15=\"http://schemas.microsoft.com/office/drawing/2012/slicer\" Requires=\"sle15\">"
                    + RelatedPartCopyFrameAnchor("10", "Region", "http://schemas.microsoft.com/office/drawing/2010/slicer")
                    + "</mc:Choice><mc:Fallback/></mc:AlternateContent>"
                    + RelatedPartCopyFrameAnchor("11", "Timeline", "http://schemas.microsoft.com/office/drawing/2012/timeslicer")
                    + "<xdr:twoCellAnchor xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"
                    + "<xdr:from><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
                    + "<xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>"
                    + "<xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"12\" name=\"Linked shape\">"
                    + "<a:extLst><a:ext uri=\"{63B3BB69-23CF-44E3-9099-C40C66FF867C}\">"
                    + "<a14:compatExt xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" spid=\"_x0000_s1025\"/>"
                    + "</a:ext></a:extLst></xdr:cNvPr><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>";
                string drawingXml = Regex.Replace(ReadPartText(drawingsPart), @"</(\w+:)?wsDr>\s*$", _ => anchors + _.Value);
                using var replacement = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(drawingXml));
                drawingsPart.FeedData(replacement);
            }

            sourceStream.Position = 0;
            using var targetStream = new MemoryStream();
            using (var sourceDocument = ExcelDocument.Load(sourceStream, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            using (var targetDocument = ExcelDocument.Create(targetStream)) {
                targetDocument.AddWorksheet("Existing").SetComment("A1", "Existing note", "Target Author");
                targetDocument.CopyWorksheetFrom(sourceDocument, "Report", "Copy", ExcelSheetNameValidationMode.Strict, new ExcelWorksheetCopyOptions {
                    CopyMode = ExcelWorksheetCopyMode.Package
                });
                targetDocument.Save();
            }

            targetStream.Position = 0;
            using (var reopened = ExcelDocument.Load(targetStream, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                AssertCopiedRelatedFeatures(reopened, "Copy");
                Assert.Empty(reopened.ValidateOpenXml());
            }

            targetStream.Position = 0;
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(targetStream, false)) {
                WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
                Sheet sheet = workbookPart.Workbook.Sheets!.Elements<Sheet>().Single(_ => _.Name == "Copy");
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
                string drawingXml = ReadPartText(worksheetPart.DrawingsPart!);
                Assert.DoesNotContain("drawing/2010/slicer", drawingXml);
                Assert.DoesNotContain("drawing/2012/timeslicer", drawingXml);

                string vmlShapeId = Regex.Match(ReadPartText(Assert.Single(worksheetPart.VmlDrawingParts)), @"_x0000_s\d+").Value;
                Assert.NotEqual("_x0000_s1025", vmlShapeId);
                Assert.Contains("spid=\"" + vmlShapeId + "\"", drawingXml);
            }
        }

        [Fact]
        public void Test_CopyWorksheetFrom_PackageModeRemapsChartTableAndDefinedNameReferences() {
            using MemoryStream sourceStream = CreateRelatedPartCopySource();
            using (var source = ExcelDocument.Load(sourceStream)) {
                source["Report"].AddTable("A1:B3", true, "SalesTable", ExcelTableStyle.TableStyleMedium2);
                source.Save();
            }
            sourceStream.Position = 0;
            using (SpreadsheetDocument source = SpreadsheetDocument.Open(sourceStream, true)) {
                WorkbookPart workbook = source.WorkbookPart!;
                workbook.Workbook.DefinedNames = new DefinedNames(new DefinedName { Name = "ChartOnlyName", Text = "Report!$A$2:$A$3" });
                var formulas = workbook.WorksheetParts.Single().DrawingsPart!.ChartParts.Single().ChartSpace!.Descendants<C.Formula>().ToList();
                formulas[0].Text = "[0]!ChartOnlyName";
                formulas[1].Text = "SalesTable[Sales]";
                workbook.Workbook.Save();
                workbook.WorksheetParts.Single().DrawingsPart!.ChartParts.Single().ChartSpace!.Save();
            }
            sourceStream.Position = 0;
            using var targetStream = new MemoryStream();
            using (var source = ExcelDocument.Load(sourceStream, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            using (var target = ExcelDocument.Create(targetStream)) {
                var existing = target.AddWorksheet("Existing");
                existing.CellValue(1, 1, "Sales");
                existing.CellValue(2, 1, 1);
                existing.AddTable("A1:A2", true, "SalesTable", ExcelTableStyle.TableStyleMedium2);
                target.CopyWorksheetFrom(source, "Report", "Copy", ExcelSheetNameValidationMode.Strict,
                    new ExcelWorksheetCopyOptions { CopyMode = ExcelWorksheetCopyMode.Package });
                target.Save();
            }
            targetStream.Position = 0;
            using var result = SpreadsheetDocument.Open(targetStream, false);
            WorkbookPart resultWorkbook = result.WorkbookPart!;
            Sheet copy = resultWorkbook.Workbook.Sheets!.Elements<Sheet>().Single(x => x.Name == "Copy");
            var part = (WorksheetPart)resultWorkbook.GetPartById(copy.Id!.Value!);
            string copiedTable = Assert.Single(part.TableDefinitionParts).Table!.Name!.Value!;
            Assert.NotEqual("SalesTable", copiedTable);
            var copiedFormulas = part.DrawingsPart!.ChartParts.Single().ChartSpace!.Descendants<C.Formula>().Select(x => x.Text).ToList();
            Assert.Contains("'Copy'!ChartOnlyName", copiedFormulas);
            Assert.Contains(copiedTable + "[Sales]", copiedFormulas);
            Assert.Contains(resultWorkbook.Workbook.DefinedNames!.Elements<DefinedName>(), x => x.Name == "ChartOnlyName" && x.Text!.Contains("Copy"));
        }

        [Fact]
        public void Test_CopyWorksheetFrom_PackageModeRemapsChartExternalWorkbookIndexes() {
            string sourcePath = Path.Combine(_directoryWithFiles, "ChartExternalSource.xlsx");
            string targetPath = Path.Combine(_directoryWithFiles, "ChartExternalTarget.xlsx");
            using (var source = ExcelDocument.Create(sourcePath)) {
                var sheet = source.AddWorksheet("Report");
                sheet.CellValue(1, 1, "Value");
                sheet.CellValue(2, 1, "Category");
                sheet.CellValue(1, 2, "Series");
                sheet.CellValue(2, 2, 1);
                sheet.AddChartFromRange("A1:B2", 4, 1);
                source.Save();
            }
            using (var target = ExcelDocument.Create(targetPath)) {
                target.AddWorksheet("Existing");
                target.Save();
            }
            AddExternalWorkbookReference(sourcePath);
            AddExternalWorkbookReference(targetPath);
            using (SpreadsheetDocument source = SpreadsheetDocument.Open(sourcePath, true)) {
                ChartPart chart = source.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.ChartParts.Single();
                chart.ChartSpace!.Descendants<C.Formula>().First().Text = "[1]Sheet1!$A$1:$A$2";
                chart.ChartSpace.Save();
            }
            using (var source = ExcelDocument.Load(sourcePath, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            using (var target = ExcelDocument.Load(targetPath)) {
                target.CopyWorksheetFrom(source, "Report", "Copy", ExcelSheetNameValidationMode.Strict,
                    new ExcelWorksheetCopyOptions { CopyMode = ExcelWorksheetCopyMode.Package, CopyExternalWorkbookReferences = true });
                target.Save();
            }
            using (SpreadsheetDocument result = SpreadsheetDocument.Open(targetPath, false)) {
                WorkbookPart workbook = result.WorkbookPart!;
                Sheet copy = workbook.Workbook.Sheets!.Elements<Sheet>().Single(x => x.Name == "Copy");
                var part = (WorksheetPart)workbook.GetPartById(copy.Id!.Value!);
                Assert.Contains(part.DrawingsPart!.ChartParts.Single().ChartSpace!.Descendants<C.Formula>(), x => x.Text == "[2]Sheet1!$A$1:$A$2");
            }
            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Test_CopyWorksheetPreservesChartSheetScopedName(bool useCopyFrom) {
            using MemoryStream stream = CreateRelatedPartCopySource();
            using (SpreadsheetDocument source = SpreadsheetDocument.Open(stream, true)) {
                WorkbookPart workbook = source.WorkbookPart!;
                workbook.Workbook.DefinedNames = new DefinedNames(new DefinedName { Name = "ChartLocal", LocalSheetId = 0U, Text = "Report!$A$2:$A$3" });
                var chart = workbook.WorksheetParts.Single().DrawingsPart!.ChartParts.Single();
                chart.ChartSpace!.Descendants<C.Formula>().First().Text = "Report!ChartLocal";
                chart.ChartSpace.Save();
                workbook.Workbook.Save();
            }
            stream.Position = 0;
            using (var document = ExcelDocument.Load(stream)) {
                if (useCopyFrom) {
                    document.CopyWorksheetFrom(document, "Report", "Copy", ExcelSheetNameValidationMode.Strict,
                        new ExcelWorksheetCopyOptions { CopyMode = ExcelWorksheetCopyMode.Package });
                } else {
                    document.CopyWorksheet("Report", "Copy");
                }
                document.Save();
            }
            stream.Position = 0;
            using SpreadsheetDocument result = SpreadsheetDocument.Open(stream, false);
            WorkbookPart resultWorkbook = result.WorkbookPart!;
            Assert.Contains(resultWorkbook.Workbook.DefinedNames!.Elements<DefinedName>(),
                x => x.Name == "ChartLocal" && x.LocalSheetId?.Value == 1U && x.Text!.Contains("Copy"));
            Sheet copy = resultWorkbook.Workbook.Sheets!.Elements<Sheet>().Single(x => x.Name == "Copy");
            var part = (WorksheetPart)resultWorkbook.GetPartById(copy.Id!.Value!);
            Assert.Contains(part.DrawingsPart!.ChartParts.Single().ChartSpace!.Descendants<C.Formula>(), x => x.Text!.Contains("Copy") && x.Text.Contains("ChartLocal"));
        }

        private static string RelatedPartCopyFrameAnchor(string id, string name, string graphicDataUri) {
            return "<xdr:twoCellAnchor xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"
                + "<xdr:from><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
                + "<xdr:to><xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>"
                + "<xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"" + id + "\" name=\"" + name + "\"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>"
                + "<xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></xdr:xfrm>"
                + "<a:graphic><a:graphicData uri=\"" + graphicDataUri + "\"/></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>";
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
