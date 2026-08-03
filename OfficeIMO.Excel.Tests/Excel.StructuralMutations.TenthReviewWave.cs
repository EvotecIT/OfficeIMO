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
        public void Test_FormulaSyntaxTree_PreservesSpacedFunctionCallsAndLexicalBindings() {
            ExcelFormulaSyntaxTree references = ExcelFormulaSyntaxTree.Parse("=LOG10 (A1)+A1");

            Assert.Equal(2, references.Nodes.OfType<ExcelFormulaReferenceSyntax>().Count());
            Assert.Equal("=LOG10 (B1)+B1", references.Rewrite(reference => reference.Offset(0, 1)));

            ExcelFormulaSyntaxTree names = ExcelFormulaSyntaxTree.Parse(
                "=LET(Input,1,Input+TaxRate)+LAMBDA(Input,Input+TaxRate)(2)");
            string rewritten = names.RewriteNames(name =>
                string.Equals(name, "Input", StringComparison.OrdinalIgnoreCase) ? "RenamedInput" :
                string.Equals(name, "TaxRate", StringComparison.OrdinalIgnoreCase) ? "Rate2026" : name);

            Assert.Equal(
                "=LET(Input,1,Input+Rate2026)+LAMBDA(Input,Input+Rate2026)(2)",
                rewritten);
        }

        [Fact]
        public void Test_StructuralRows_EnforcesAffectedCellBudgetAndReportsAppliedCount() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Header");
            sheet.CellValue(2, 1, 10);
            sheet.CellValue(3, 1, 20);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                sheet.PlanInsertRows(
                    2,
                    options: new ExcelMutationPlanOptions { MaximumAffectedCells = 1 }));

            Assert.Contains("MaximumAffectedCells", exception.Message, StringComparison.Ordinal);
            ExcelRowMutationPlan plan = sheet.PlanInsertRows(
                2,
                options: new ExcelMutationPlanOptions { MaximumAffectedCells = 2 });
            ExcelMutationResult result = plan.ApplyWithDiagnostics();
            Assert.Equal(2, result.AffectedCells);
        }

        [Fact]
        public void Test_ReferenceAlgebra_PreservesWholeRowAndWholeColumnKinds() {
            ExcelReference columns = ExcelReference.Parse("A:C");
            ExcelReference columnOverlap = columns.Intersect(ExcelReference.Parse("B:D"))!;
            ExcelReference columnUnion = columns.BoundingUnion(ExcelReference.Parse("B:D"));
            ExcelReference[] columnRemainder = columns.Except(ExcelReference.Parse("B:B")).ToArray();

            Assert.Equal(ExcelReferenceKind.WholeColumn, columnOverlap.Kind);
            Assert.Equal("B:C", columnOverlap.ToString());
            Assert.Equal(ExcelReferenceKind.WholeColumn, columnUnion.Kind);
            Assert.Equal("A:D", columnUnion.ToString());
            Assert.Equal(new[] { "A:A", "C:C" }, columnRemainder.Select(item => item.ToString()));
            Assert.All(columnRemainder, item => Assert.Equal(ExcelReferenceKind.WholeColumn, item.Kind));

            ExcelReference rows = ExcelReference.Parse("1:3");
            ExcelReference rowOverlap = rows.Intersect(ExcelReference.Parse("2:4"))!;
            Assert.Equal(ExcelReferenceKind.WholeRow, rowOverlap.Kind);
            Assert.Equal("2:3", rowOverlap.ToString());
        }

        [Fact]
        public void Test_ColumnCellAndMoveMutations_RemapSparklineDestinations() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet columns = document.AddWorksheet("Columns");
            columns.AddSparklines("A1:B1", "D1");
            columns.InsertColumns(1);
            Assert.Equal("E1", Assert.Single(columns.WorksheetPart.Worksheet
                .Descendants<X14.Sparkline>()).ReferenceSequence!.Text);

            ExcelSheet cells = document.AddWorksheet("Cells");
            cells.AddSparklines("A1:B1", "D1");
            cells.InsertCells("A1", ExcelCellShiftDirection.Right);
            Assert.Equal("E1", Assert.Single(cells.WorksheetPart.Worksheet
                .Descendants<X14.Sparkline>()).ReferenceSequence!.Text);

            ExcelSheet moved = document.AddWorksheet("Moved");
            moved.AddSparklines("A1:B1", "D1");
            moved.MoveRange("D1", "E1");
            Assert.Equal("E1", Assert.Single(moved.WorksheetPart.Worksheet
                .Descendants<X14.Sparkline>()).ReferenceSequence!.Text);
        }

        [Fact]
        public void Test_ColumnCellAndMoveMutations_RemapQueryTableSortRanges() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet columns = document.AddWorksheet("Columns");
            SortState columnSort = AddQueryTableRefreshSortState(document, columns, "A1:C10");
            columns.InsertColumns(1);
            Assert.Equal("B1:D10", columnSort.Reference!.Value);
            Assert.Equal("B1:B10", Assert.Single(columnSort.Elements<SortCondition>()).Reference!.Value);

            ExcelSheet cells = document.AddWorksheet("Cells");
            SortState cellSort = AddQueryTableRefreshSortState(document, cells, "A1:C1");
            cells.InsertCells("A1", ExcelCellShiftDirection.Right);
            Assert.Equal("B1:D1", cellSort.Reference!.Value);

            ExcelSheet moved = document.AddWorksheet("Moved");
            SortState movedSort = AddQueryTableRefreshSortState(document, moved, "A1:C1");
            moved.MoveRange("A1:C1", "B2");
            Assert.Equal("B2:D2", movedSort.Reference!.Value);
        }

        [Fact]
        public void Test_MutationSnapshot_RestoresQueryTableRoots() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            AddQueryTableRefreshSortState(document, sheet, "A1:C10");
            QueryTablePart queryPart = Assert.Single(sheet.WorksheetPart.QueryTableParts);
            using var cancellation = new CancellationTokenSource();

            Assert.Throws<OperationCanceledException>(() => sheet.ApplyTransactionalMutation(_ => {
                Assert.Single(queryPart.QueryTable!.Descendants<SortState>()).Reference = "B2:D11";
                cancellation.Cancel();
            }, 0, new ExcelMutationPlanOptions(), cancellation.Token));

            Assert.Equal(
                "A1:C10",
                Assert.Single(queryPart.QueryTable!.Descendants<SortState>()).Reference!.Value);
        }
    }
}
