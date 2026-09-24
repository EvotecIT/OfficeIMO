using OfficeIMO.Excel;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.OpenDocument;
using System;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class SpreadsheetCustomValidationConversionTests {
    [Fact]
    public void RelativeCellComparisonRoundTripsWithItsBaseCell() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Data Set");
        sheet.ValidationCustomFormula("C5:C7", "C5>0", allowBlank: false,
            errorTitle: "Positive", errorMessage: "Enter a positive number.");

        OdfConversionResult<OdsDocument> toOds = source.ToOpenDocumentResult();
        OdsDocument ods = OdsDocument.Load(new MemoryStream(toOds.Value.ToBytes()));
        Assert.True(ods.Validate().IsValid);
        OdsValidation validation = Assert.Single(ods.Validations);
        Assert.Equal("of:is-true-formula([.C5]>0)", validation.Condition);
        Assert.Equal("$'Data Set'.$C$5", validation.BaseCellAddress);
        Assert.False(validation.AllowEmptyCell);
        Assert.DoesNotContain(toOds.Report.Mappings, mapping => mapping.Feature == "validations"
            && mapping.Status == OdfConversionMappingStatus.Unsupported);

        OdfConversionResult<ExcelDocument> toExcel = ods.ToExcelDocumentResult();
        using ExcelDocument reopened = ExcelDocument.Load(new MemoryStream(toExcel.Value.ToBytes()));
        ExcelDataValidationInfo result = Assert.Single(reopened.Sheets.Single().GetDataValidations());
        Assert.Equal("custom", result.Type);
        Assert.Equal("C5>0", result.Formula1);
        Assert.Equal("C5:C7", result.Range);
        Assert.False(result.AllowBlank);
        Assert.Equal("Positive", result.ErrorTitle);
        Assert.Equal("Enter a positive number.", result.Error);
        Assert.DoesNotContain(toExcel.Report.Mappings, mapping => mapping.Feature == "validations"
            && mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void FunctionsAndDisjointCustomRangesRemainExplicitLoss() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Data");
        source.AddWorksheet("Other");
        sheet.ValidationCustomFormula("C1:C2", "SUM(A1:B1)>10");
        sheet.ValidationCustomFormula("E1 E3", "E1>0");
        sheet.ValidationCustomFormula("G1", "Other!A1>0");

        OdfConversionResult<OdsDocument> result = source.ToOpenDocumentResult();
        Assert.Empty(result.Value.Validations);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "validations"
            && mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 3);
    }

    [Theory]
    [InlineData("D4<>\"\"")]
    [InlineData("D4<E4")]
    [InlineData("D4>=-1")]
    [InlineData("D4<+2.5")]
    [InlineData("(D4)>0")]
    [InlineData("D4>(0)")]
    [InlineData("(D4>0)")]
    [InlineData("AND(D4>0,D4<10)")]
    [InlineData("OR(D4=\"\",AND(D4>=-1,D4<=1))")]
    [InlineData("NOT(D4=\"\")")]
    [InlineData("AND(D4>0,NOT(OR(D4=5,D4=10)))")]
    public void TextAndCellComparisonOperandsRoundTrip(string formula) {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Data");
        sheet.ValidationCustomFormula("D4", formula);

        OdfConversionResult<OdsDocument> toOds = source.ToOpenDocumentResult();
        OdsDocument ods = OdsDocument.Load(new MemoryStream(toOds.Value.ToBytes()));
        Assert.Single(ods.Validations);
        Assert.True(ods.Validate().IsValid);
        OdfConversionResult<ExcelDocument> toExcel = ods.ToExcelDocumentResult();
        using ExcelDocument reopened = ExcelDocument.Load(new MemoryStream(toExcel.Value.ToBytes()));
        Assert.Equal(formula, Assert.Single(reopened.Sheets.Single().GetDataValidations()).Formula1);
    }

    [Fact]
    public void OdsCustomRuleWithoutMatchingBaseCellRemainsExplicitLoss() {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        OdsValidation validation = source.AddValidation("Custom",
            OdsValidationConditionSyntax.CreateFormula("[.A1]>0"));
        validation.BaseCellAddress = "$'Other'.$A$1";
        sheet.Cell(0, 0).ValidationName = validation.Name;

        OdfConversionResult<ExcelDocument> result = source.ToExcelDocumentResult();
        Assert.Empty(result.Value.Sheets.Single().GetDataValidations());
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "validations"
            && mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void MalformedCustomComparisonRemainsExplicitLoss() {
        using ExcelDocument source = ExcelDocument.Create();
        source.AddWorksheet("Data").ValidationCustomFormula("C5", "C5>0()");

        OdfConversionResult<OdsDocument> result = source.ToOpenDocumentResult();
        Assert.Empty(result.Value.Validations);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "validations"
            && mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Theory]
    [InlineData("AND(C5>0)")]
    [InlineData("OR(C5>0,)")]
    [InlineData("AND(C5>0,SUM(C5)>1)")]
    [InlineData("NOT()")]
    [InlineData("NOT(C5>0,C5<10)")]
    [InlineData("NOT(SUM(C5)>1)")]
    public void UnsupportedBooleanCustomFormulaRemainsExplicitLoss(string formula) {
        using ExcelDocument source = ExcelDocument.Create();
        source.AddWorksheet("Data").ValidationCustomFormula("C5", formula);

        OdfConversionResult<OdsDocument> result = source.ToOpenDocumentResult();
        Assert.Empty(result.Value.Validations);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "validations"
            && mapping.Status == OdfConversionMappingStatus.Unsupported);
    }
}
