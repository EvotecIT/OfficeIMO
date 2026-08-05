using System;
using System.IO;
using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_MutationSnapshot_RestoresDeletedCommentPartRelationships() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.SetComment("A1", "Legacy", "Tester");
            ExcelThreadedCommentResult threaded = sheet.AddThreadedComment("B1", "Threaded", "Tester");
            WorksheetCommentsPart legacyPart = sheet.WorksheetPart.WorksheetCommentsPart!;
            WorksheetThreadedCommentsPart threadedPart = Assert.Single(sheet.WorksheetPart.WorksheetThreadedCommentsParts);
            VmlDrawingPart vmlPart = Assert.Single(sheet.WorksheetPart.VmlDrawingParts);
            string legacyRelationshipId = sheet.WorksheetPart.GetIdOfPart(legacyPart);
            string threadedRelationshipId = sheet.WorksheetPart.GetIdOfPart(threadedPart);
            string vmlRelationshipId = sheet.WorksheetPart.GetIdOfPart(vmlPart);
            using var cancellation = new CancellationTokenSource();

            Assert.Throws<OperationCanceledException>(() => sheet.ApplyTransactionalMutation(_ => {
                sheet.ClearComment("A1");
                Assert.True(sheet.RemoveThreadedComment(threaded.Id));
                Assert.Null(sheet.WorksheetPart.WorksheetCommentsPart);
                Assert.Empty(sheet.WorksheetPart.WorksheetThreadedCommentsParts);
                Assert.Empty(sheet.WorksheetPart.VmlDrawingParts);
                cancellation.Cancel();
            }, 0, new ExcelMutationPlanOptions(), cancellation.Token));

            Assert.Equal(legacyRelationshipId, sheet.WorksheetPart.GetIdOfPart(sheet.WorksheetPart.WorksheetCommentsPart!));
            Assert.Equal(threadedRelationshipId, sheet.WorksheetPart.GetIdOfPart(Assert.Single(sheet.WorksheetPart.WorksheetThreadedCommentsParts)));
            Assert.Equal(vmlRelationshipId, sheet.WorksheetPart.GetIdOfPart(Assert.Single(sheet.WorksheetPart.VmlDrawingParts)));
            Assert.Equal("Legacy", Assert.Single(sheet.GetComments()).Text);
            Assert.Equal("Threaded", Assert.Single(sheet.GetThreadedComments()).Text);
            Assert.Empty(document.ValidateDocument());
        }

        [Fact]
        public void Test_AutoFilterCriteria_RejectQualifiedRangesWithoutReplacingLocalState() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            document.AddWorksheet("Other");
            sheet.AutoFilterBlanks("A1:C3", 1);

            Assert.Throws<ArgumentException>(() => sheet.AutoFilterBlanks("Other!A1:C3", 1));
            Assert.Throws<ArgumentException>(() => sheet.AutoFilterTopBottom("Data!A1:C3", 2, 5));

            ExcelAutoFilterInfo filter = Assert.Single(sheet.GetAutoFilters());
            Assert.Equal("A1:C3", filter.Range);
            Assert.Equal(1U, Assert.Single(filter.Columns).ColumnOffset);
        }

        [Fact]
        public void Test_TableResize_RejectsMergedCellIntersectionBeforeMutation() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "A");
            sheet.CellValue(2, 1, 1);
            sheet.CellValue(3, 1, 2);
            sheet.AddTable("A1:A3", true, "Sales", OfficeIMO.Excel.TableStyle.TableStyleMedium2);
            sheet.Range("B2:C2").Merge();

            Assert.Throws<InvalidOperationException>(() =>
                sheet.ResizeTable("Sales", "A1:C3"));

            Assert.Equal("A1:A3", Assert.Single(document.GetTables()).Range);
        }

        [Fact]
        public void Test_StructuralColumns_RejectsEmptyBoundaryColumnDefinition() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            var columns = new Columns(new Column {
                Min = (uint)A1.MaxColumns,
                Max = (uint)A1.MaxColumns,
                Width = 12D,
                CustomWidth = true,
            });
            sheet.WorksheetPart.Worksheet.InsertBefore(columns, sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>());

            Assert.Throws<InvalidOperationException>(() => sheet.PlanInsertColumns(1));

            Column column = Assert.Single(sheet.WorksheetPart.Worksheet.Descendants<Column>());
            Assert.Equal((uint)A1.MaxColumns, column.Max!.Value);
        }

        [Fact]
        public void Test_SparklineType_LineDefaultDoesNotSplitImportedGroup() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.AddSparklines("A1:C2", "D1:D2");
            X14.SparklineGroup group = Assert.Single(sheet.WorksheetPart.Worksheet.Descendants<X14.SparklineGroup>());
            group.Type = null;

            Assert.Equal(0, sheet.SetSparklineType("D1", ExcelSparklineType.Line));
            Assert.Single(sheet.WorksheetPart.Worksheet.Descendants<X14.SparklineGroup>());
            Assert.Equal(2, sheet.GetSparklines().Count);
        }

        [Fact]
        public void Test_StructuralColumns_RemapsWorksheetFilterCriteria() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.AutoFilterBlanks("A1:C3", 1);
            sheet.AutoFilterTopBottom("A1:C3", 2, 5);

            sheet.InsertColumns(2);

            ExcelAutoFilterInfo inserted = Assert.Single(sheet.GetAutoFilters());
            Assert.Equal("A1:D3", inserted.Range);
            Assert.Equal(new uint[] { 2U, 3U }, inserted.Columns.Select(column => column.ColumnOffset).OrderBy(value => value));

            sheet.DeleteColumns(3);

            ExcelAutoFilterInfo deleted = Assert.Single(sheet.GetAutoFilters());
            Assert.Equal("A1:C3", deleted.Range);
            Assert.Equal(new uint[] { 2U }, deleted.Columns.Select(column => column.ColumnOffset));
        }

        [Fact]
        public void Test_StructuralColumns_RemapsManualColumnPageBreaks() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.AddManualColumnPageBreak(2);
            sheet.AddManualColumnPageBreak(4);

            sheet.InsertColumns(2);
            Assert.Equal(new[] { 3, 5 }, sheet.GetManualColumnPageBreaks());

            sheet.DeleteColumns(3);
            Assert.Equal(new[] { 4 }, sheet.GetManualColumnPageBreaks());
            ColumnBreaks columnBreaks = sheet.WorksheetPart.Worksheet.GetFirstChild<ColumnBreaks>()!;
            Assert.Equal(1U, columnBreaks.Count!.Value);
            Assert.Equal(1U, columnBreaks.ManualBreakCount!.Value);
        }
    }
}
