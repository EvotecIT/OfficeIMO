using System;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentFormulaTests {
    [Fact]
    public void FormulaMutationInvalidatesCachedValueAndAllowsExplicitReplacement() {
        OdsDocument document = OdsDocument.Create();
        OdsCell cell = document.AddSheet("Data").Cell(0, 0);
        cell.SetNumber(42D);

        cell.Formula = "of:=1+1";

        Assert.Equal(OdsCellValueKind.Empty, cell.Value.Kind);
        cell.SetNumber(2D);
        Assert.Equal(2D, cell.Value.AsDouble());
        Assert.Equal("of:=1+1", cell.Formula);
    }

    [Fact]
    public void EvaluatesArithmeticRangesFunctionsAndCrossSheetReferences() {
        OdsDocument document = OdsDocument.Create();
        OdsSheet data = document.AddSheet("Data");
        OdsSheet other = document.AddSheet("Other");
        data.Cell(0, 0).SetNumber(2D);
        data.Cell(1, 0).SetNumber(3D);
        other.Cell(0, 0).SetNumber(4D);
        data.Cell(0, 1).Formula = "of:=SUM([.A1:.A2])*2+POWER(2;3)+[$'Other'.$A$1]";

        OdsFormulaEvaluationResult result = OdsFormulaEvaluator.EvaluateCell(document, "Data", 0, 1);

        Assert.True(result.Success, result.Error);
        Assert.Equal(22D, result.Value.AsNumber());
        Assert.InRange(result.Operations, 1, 100);
    }

    [Fact]
    public void ReportsCyclesUnsupportedFunctionsAndRangeBoundsWithoutExecutingAnything() {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        sheet.Cell(0, 0).Formula = "of:=[.B1]";
        sheet.Cell(0, 1).Formula = "of:=[.A1]";

        OdsFormulaEvaluationResult cycle = OdsFormulaEvaluator.EvaluateCell(document, "Data", 0, 0);
        OdsFormulaEvaluationResult unsupported = OdsFormulaEvaluator.EvaluateExpression(document, "Data", "of:=WEBSERVICE(\"https://example.com\")");
        OdsFormulaEvaluationResult bounded = OdsFormulaEvaluator.EvaluateExpression(document, "Data", "of:=SUM([.A1:.A10])",
            new OdsFormulaEvaluationOptions { MaximumRangeCells = 2 });

        Assert.False(cycle.Success);
        Assert.Contains("Cyclic", cycle.Error, StringComparison.OrdinalIgnoreCase);
        Assert.False(unsupported.Success);
        Assert.Contains("Unsupported", unsupported.Error, StringComparison.OrdinalIgnoreCase);
        Assert.False(bounded.Success);
        Assert.Contains("range-cell limit", bounded.Error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void RecalculationUpdatesSuccessfulCachedValuesAndLeavesFailuresReported() {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        sheet.Cell(0, 0).SetNumber(4D);
        sheet.Cell(0, 1).Formula = "of:=[.A1]*3";
        sheet.Cell(0, 2).Formula = "of:=1/0";

        OdsRecalculationReport report = document.Recalculate();

        Assert.Equal(2, report.FormulaCells);
        Assert.Equal(1, report.UpdatedCells);
        Assert.Equal(1, report.FailedCells);
        Assert.Equal(12D, sheet.GetValue(0, 1).AsDouble());
        Assert.Equal(OdsCellValueKind.Empty, sheet.GetValue(0, 2).Kind);
        Assert.True(document.Validate().IsValid);
    }

    [Theory]
    [InlineData("of:=ROUND(2.5;0)", 3D)]
    [InlineData("of:=ROUND(-2.5;0)", -3D)]
    [InlineData("of:=ROUND(1234;-2)", 1200D)]
    [InlineData("of:=ROUND(1250;-2)", 1300D)]
    [InlineData("of:=ROUND(-1250;-2)", -1300D)]
    [InlineData("of:=-2^2", -4D)]
    [InlineData("of:=2^-2", 0.25D)]
    public void FormulaEvaluationUsesSpreadsheetNumericSemantics(string formula, double expected) {
        OdsDocument document = OdsDocument.Create();
        document.AddSheet("Data");

        OdsFormulaEvaluationResult result = OdsFormulaEvaluator.EvaluateExpression(document, "Data", formula);

        Assert.True(result.Success, result.Error);
        Assert.Equal(expected, result.Value.AsNumber());
    }
}
