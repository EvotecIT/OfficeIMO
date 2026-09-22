using OfficeIMO.Excel;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.OpenDocument;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class SpreadsheetTemporalValidationConversionTests {
    [Theory]
    [InlineData(ExcelDateSystem.NineteenHundred)]
    [InlineData(ExcelDateSystem.NineteenFour)]
    public void ExcelDateAndWholeSecondTimeValidationsSurviveOdsAndXlsxReopen(
        ExcelDateSystem dateSystem) {
        var firstDate = new DateTime(2024, 2, 29);
        var secondDate = new DateTime(2024, 12, 31);
        var time = new TimeSpan(9, 30, 15);
        using ExcelDocument source = ExcelDocument.Create();
        source.DateSystem = dateSystem;
        ExcelSheet sheet = source.AddWorksheet("Data");
        sheet.ValidationDate("A1:A3", ExcelDataValidationOperator.Between,
            firstDate, secondDate, allowBlank: false);
        sheet.ValidationTime("B1:B3", ExcelDataValidationOperator.Equal, time);

        OdfConversionResult<OdsDocument> toOds = source.ToOpenDocumentResult();
        Assert.DoesNotContain(toOds.Report.Mappings, mapping => mapping.Feature == "validations"
            && mapping.Status == OdfConversionMappingStatus.Unsupported);
        OdsDocument ods = OdsDocument.Load(new MemoryStream(toOds.Value.ToBytes()));
        Assert.True(ods.Validate().IsValid);
        Assert.Contains(ods.Validations, validation => validation.ParsedCondition?.ValueKind == OdsValidationValueKind.Date
            && validation.ParsedCondition.FirstOperand == "DATE(2024;2;29)"
            && validation.ParsedCondition.SecondOperand == "DATE(2024;12;31)"
            && !validation.AllowEmptyCell);
        Assert.Contains(ods.Validations, validation => validation.ParsedCondition?.ValueKind == OdsValidationValueKind.Time
            && validation.ParsedCondition.FirstOperand == "TIME(9;30;15)");

        OdfConversionResult<ExcelDocument> toExcel = ods.ToExcelDocumentResult();
        using ExcelDocument converted = toExcel.Value;
        using ExcelDocument reopened = ExcelDocument.Load(new MemoryStream(converted.ToBytes()));
        var validations = reopened.Sheets.Single().GetDataValidations();
        Assert.Equal(2, validations.Count);
        var date = Assert.Single(validations, validation => validation.Type == "date");
        Assert.Equal("between", date.Operator);
        Assert.False(date.AllowBlank);
        Assert.Equal(ExcelDateSystemConverter.ToSerial(firstDate, ExcelDateSystem.NineteenHundred),
            double.Parse(date.Formula1!, CultureInfo.InvariantCulture), 8);
        Assert.Equal(ExcelDateSystemConverter.ToSerial(secondDate, ExcelDateSystem.NineteenHundred),
            double.Parse(date.Formula2!, CultureInfo.InvariantCulture), 8);
        var clock = Assert.Single(validations, validation => validation.Type == "time");
        Assert.Equal("equal", clock.Operator);
        Assert.Equal(time.TotalDays, double.Parse(clock.Formula1!, CultureInfo.InvariantCulture), 12);
        Assert.DoesNotContain(toExcel.Report.Mappings, mapping => mapping.Feature == "validations"
            && mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void ProducerStyleDateAndTimeConstantsProjectToExcel() {
        OdsDocument source = OdsDocument.Create();
        OdsValidation date = source.AddValidation("DateWindow", OdsValidationConditionSyntax.Create(
            OdsValidationValueKind.Date, OdsValidationComparison.Between,
            "DATE(2026;1;1)", "DATE(2026;1;31)"));
        OdsValidation time = source.AddValidation("StartTime", OdsValidationConditionSyntax.Create(
            OdsValidationValueKind.Time, OdsValidationComparison.GreaterThanOrEqual,
            "TIME(8;30;0)"));
        OdsSheet sheet = source.AddSheet("Data");
        sheet.Cell(0, 0).ValidationName = date.Name;
        sheet.Cell(0, 1).ValidationName = time.Name;

        OdsDocument reopened = OdsDocument.Load(new MemoryStream(source.ToBytes()));
        OdfConversionResult<ExcelDocument> conversion = reopened.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;
        var validations = target.Sheets.Single().GetDataValidations();
        Assert.Equal(2, validations.Count);
        Assert.Contains(validations, validation => validation.Type == "date"
            && validation.Operator == "between");
        Assert.Contains(validations, validation => validation.Type == "time"
            && validation.Operator == "greaterThanOrEqual");
        Assert.DoesNotContain(conversion.Report.Mappings, mapping => mapping.Feature == "validations"
            && mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void DynamicDateRuleAndFractionalTimeRemainExplicitLosses() {
        OdsDocument ods = OdsDocument.Create();
        OdsValidation date = ods.AddValidation("DynamicDate", OdsValidationConditionSyntax.Create(
            OdsValidationValueKind.Date, OdsValidationComparison.GreaterThan, "TODAY()"));
        ods.AddSheet("Data").Cell(0, 0).ValidationName = date.Name;
        OdfConversionResult<ExcelDocument> toExcel = ods.ToExcelDocumentResult();
        using ExcelDocument target = toExcel.Value;
        Assert.Empty(target.Sheets.Single().GetDataValidations());
        Assert.Contains(toExcel.Report.Mappings, mapping => mapping.Feature == "validations"
            && mapping.Status == OdfConversionMappingStatus.Unsupported);

        using ExcelDocument excel = ExcelDocument.Create();
        excel.AddWorksheet("Data").ValidationTime("A1", ExcelDataValidationOperator.Equal,
            TimeSpan.FromMilliseconds(500));
        OdfConversionResult<OdsDocument> toOds = excel.ToOpenDocumentResult();
        Assert.Empty(toOds.Value.Validations);
        Assert.Contains(toOds.Report.Mappings, mapping => mapping.Feature == "validations"
            && mapping.Status == OdfConversionMappingStatus.Unsupported);
    }
}
