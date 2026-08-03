using System;
using System.IO;
using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("validation")]
        [InlineData("conditional-formatting")]
        [InlineData("allowed-edit")]
        [InlineData("ignored-error")]
        public void Test_RangeMove_RejectsPartiallyIntersectingRangeMetadata(string kind) {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, 1);
            sheet.CellValue(1, 2, 2);
            switch (kind) {
                case "validation":
                    sheet.WorksheetPart.Worksheet.Append(new DataValidations(
                        new DataValidation {
                            Type = DataValidationValues.Whole,
                            SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:B1" }
                        }) { Count = 1U });
                    break;
                case "conditional-formatting":
                    sheet.AddConditionalFormulaRule("A1:B1", "A1>0");
                    break;
                case "allowed-edit":
                    sheet.Protect();
                    sheet.SetAllowedEditRange("Inputs", new[] { "A1:B1" });
                    break;
                case "ignored-error":
                    sheet.AddIgnoredErrorRegion(new[] { "A1:B1" }, ExcelIgnoredErrorKind.NumberStoredAsText);
                    break;
            }

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => sheet.PlanMoveRange("A1", "D1"));

            Assert.Contains("range metadata", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(1, sheet.CellAt(1, 1).GetValue<int>());
            Assert.Equal(2, sheet.CellAt(1, 2).GetValue<int>());
        }

        [Fact]
        public void Test_RangeMove_ReplacesDestinationLegacyAndThreadedComments() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Legacy source");
            sheet.CellValue(1, 2, "Threaded source");
            sheet.SetComment("A1", "Keep legacy", "Tester");
            sheet.SetComment("E1", "Replace legacy", "Tester");
            sheet.AddThreadedComment("B1", "Keep threaded", "Tester");
            sheet.AddThreadedComment("F1", "Replace threaded", "Tester");

            sheet.MoveRange("A1:B1", "E1");

            ExcelCommentInfo legacy = Assert.Single(sheet.GetComments());
            Assert.Equal("E1", legacy.CellReference);
            Assert.Equal("Keep legacy", legacy.Text);
            ExcelThreadedCommentSnapshot threaded = Assert.Single(sheet.GetThreadedComments());
            Assert.Equal("F1", threaded.CellReference);
            Assert.Equal("Keep threaded", threaded.Text);
            Assert.Single(sheet.WorksheetPart.VmlDrawingParts);
            Assert.Empty(document.ValidateDocument());
        }

        [Fact]
        public void Test_MutationSnapshot_RestoresNamedSheetViews() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            AddNamedSheetViewFilter(sheet, "A1:B2");
            NamedSheetViewsPart part = Assert.Single(sheet.WorksheetPart.NamedSheetViewsParts);
            using var cancellation = new CancellationTokenSource();

            Assert.Throws<OperationCanceledException>(() => sheet.ApplyTransactionalMutation(_ => {
                part.NamedSheetViews!.Descendants()
                    .Single(element => element.LocalName == "nsvFilter")
                    .SetAttribute(new OpenXmlAttribute("ref", string.Empty, "C1:D2"));
                part.NamedSheetViews.Save();
                cancellation.Cancel();
            }, 0, new ExcelMutationPlanOptions(), cancellation.Token));

            OpenXmlElement restored = Assert.Single(sheet.WorksheetPart.NamedSheetViewsParts)
                .NamedSheetViews!.Descendants()
                .Single(element => element.LocalName == "nsvFilter");
            Assert.Equal("A1:B2", restored.GetAttribute("ref", string.Empty).Value);
        }

        [Fact]
        public void Test_ColumnAndCellEdits_HonorDrawingPlacementModes() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet absoluteSheet = CreateDrawingSheet(document);
            Xdr.TwoCellAnchor absolute = ReplaceWithTwoCellAnchor(
                absoluteSheet,
                fromRow: 4,
                toRow: 7,
                toRowOffset: "0",
                Xdr.EditAsValues.Absolute);
            absolute.FromMarker!.ColumnId!.Text = (A1.MaxColumns - 1).ToString();
            absolute.ToMarker!.ColumnId!.Text = (A1.MaxColumns - 1).ToString();

            absoluteSheet.InsertColumns(1);

            Assert.Equal((A1.MaxColumns - 1).ToString(), absolute.FromMarker.ColumnId.Text);
            Assert.Equal((A1.MaxColumns - 1).ToString(), absolute.ToMarker.ColumnId!.Text);

            ExcelSheet columnMoveOnlySheet = CreateDrawingSheet(document);
            Xdr.TwoCellAnchor columnMoveOnly = ReplaceWithTwoCellAnchor(
                columnMoveOnlySheet,
                fromRow: 4,
                toRow: 7,
                toRowOffset: "0",
                Xdr.EditAsValues.OneCell);
            columnMoveOnly.FromMarker!.ColumnId!.Text = "3";
            columnMoveOnly.ToMarker!.ColumnId!.Text = "8";

            columnMoveOnlySheet.InsertColumns(4);

            Assert.Equal("4", columnMoveOnly.FromMarker.ColumnId.Text);
            Assert.Equal("9", columnMoveOnly.ToMarker.ColumnId!.Text);

            ExcelSheet cellMoveOnlySheet = CreateDrawingSheet(document);
            Xdr.TwoCellAnchor cellMoveOnly = ReplaceWithTwoCellAnchor(
                cellMoveOnlySheet,
                fromRow: 4,
                toRow: 7,
                toRowOffset: "0",
                Xdr.EditAsValues.OneCell);
            cellMoveOnly.FromMarker!.ColumnId!.Text = "3";
            cellMoveOnly.ToMarker!.ColumnId!.Text = "8";

            cellMoveOnlySheet.InsertCells("D5", ExcelCellShiftDirection.Right);

            Assert.Equal("4", cellMoveOnly.FromMarker.ColumnId.Text);
            Assert.Equal("9", cellMoveOnly.ToMarker.ColumnId!.Text);
        }

        [Fact]
        public void Test_FileBackedCloneDestinationObservesCancellationDuringWrites() {
            using var destination = new MemoryStream();
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            using var bounded = new ExcelBoundedSeekableStream(
                destination,
                maximumBytes: 1024,
                leaveOpen: true,
                cancellation.Token);

            Assert.Throws<OperationCanceledException>(() => bounded.WriteByte(1));
            Assert.Equal(0, destination.Length);
        }

        [Theory]
        [InlineData((ExcelIgnoredErrorKind)512)]
        [InlineData(ExcelIgnoredErrorKind.NumberStoredAsText | (ExcelIgnoredErrorKind)512)]
        public void Test_IgnoredErrors_RejectUndefinedFlags(ExcelIgnoredErrorKind errors) {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");

            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.AddIgnoredErrorRegion(new[] { "A1" }, errors));
            Assert.Empty(sheet.GetIgnoredErrorRegions());
        }

        [Fact]
        public void Test_FormulaDependencies_FilterWhitespaceSeparatedCellLikeFunctions() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, 10);
            sheet.CellFormula(1, 2, "LOG10 (A1)");

            ExcelFormulaDependencyNode node = Assert.IsType<ExcelFormulaDependencyNode>(
                document.InspectFormulas().DependencyGraph.FindNode("Data", "B1"));

            Assert.Equal(new[] { "Data!A1" }, node.Dependencies);
            Assert.Empty(node.FormulaDependencies);
        }

        [Fact]
        public void Test_AutoFilterState_UsesTop10SchemaDefaults() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.AutoFilterAdd("A1:B2");
            AutoFilter filter = sheet.WorksheetPart.Worksheet.GetFirstChild<AutoFilter>()!;
            filter.Append(new FilterColumn(new Top10 { Val = 5D }) { ColumnId = 0U });

            ExcelAutoFilterColumnInfo state = Assert.Single(Assert.Single(sheet.GetAutoFilters()).Columns);

            Assert.True(state.Top);
            Assert.False(state.Percent);
            Assert.Equal(5D, state.TopValue);
        }
    }
}
