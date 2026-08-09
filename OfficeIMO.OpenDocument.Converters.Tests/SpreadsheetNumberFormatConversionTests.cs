using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class SpreadsheetNumberFormatConversionTests {
    [Fact]
    public void WholeNumberValidationWithNonIntegerLexicalOperandIsReportedUnsupported() {
        byte[] package;
        using (ExcelDocument authored = ExcelDocument.Create()) {
            authored.AddWorksheet("Data").ValidationWholeNumber(
                "A1",
                ExcelDataValidationOperator.GreaterThan,
                1,
                allowBlank: true);
            package = authored.ToBytes();
        }
        using (var stream = new MemoryStream(package)) {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(stream, true)) {
                WorkbookPart workbookPart = spreadsheet.WorkbookPart
                    ?? throw new InvalidDataException("The regression workbook has no workbook part.");
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.Single();
                Worksheet worksheet = worksheetPart.Worksheet
                    ?? throw new InvalidDataException("The regression workbook has no worksheet XML.");
                DataValidation validation = worksheet.Descendants<DataValidation>().Single();
                validation.Formula1 = new Formula1("1E3");
                worksheet.Save();
            }
            package = stream.ToArray();
        }
        using ExcelDocument source = ExcelDocument.Load(new MemoryStream(package));

        OdfConversionResult<OdsDocument> conversion = source.ToOpenDocumentResult();

        Assert.Empty(conversion.Value.Validations);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "validations"
            && mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void OdsValidationListWithCommaBearingItemIsReportedUnsupported() {
        OdsDocument source = OdsDocument.Create();
        OdsValidation validation = source.AddValidation(
            "Names",
            OdsValidationConditionSyntax.CreateList(new[] { "Last, First" }));
        OdsCell cell = source.AddSheet("Data").Cell(0, 0);
        cell.SetString("Last, First");
        cell.ValidationName = validation.Name;

        OdfConversionResult<ExcelDocument> conversion = source.ToExcelDocumentResult();
        using ExcelDocument target = conversion.Value;

        Assert.Empty(target.Sheets.Single().GetDataValidations());
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "validations"
            && mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void ScalarAndListValidationsRoundTripThroughTypedOdfConditions() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Data");
        sheet.ValidationWholeNumber("B2:B3", ExcelDataValidationOperator.Between, 1, 10, allowBlank: false);
        sheet.SetDataValidationMessages("B2:B3", new ExcelDataValidationMessageOptions {
            PromptTitle = "Allowed values",
            Prompt = "Enter 1 through 10.",
            ShowInputMessage = true,
            ErrorTitle = "Outside range",
            Error = "The value must be between 1 and 10.",
            ShowErrorMessage = true,
            ErrorStyle = ExcelDataValidationErrorStyle.Warning,
            PreserveShowMessageFlags = true
        });
        sheet.ValidationList("C2:C3", new[] { "New", "On \"Hold\"", "Done" });

        OdfConversionResult<OdsDocument> toOds = source.ToOpenDocumentResult();
        Assert.Equal(2, toOds.Value.Validations.Count);
        Assert.Contains(toOds.Value.Validations, validation =>
            validation.ParsedCondition?.ValueKind == OdsValidationValueKind.WholeNumber
            && validation.ParsedCondition.Comparison == OdsValidationComparison.Between
            && !validation.AllowEmptyCell
            && validation.HelpTitle == "Allowed values"
            && validation.HelpText == "Enter 1 through 10."
            && validation.ErrorTitle == "Outside range"
            && validation.ErrorMessageType == OdsValidationMessageType.Warning);
        Assert.Contains(toOds.Value.Validations, validation =>
            validation.ParsedCondition?.ValueKind == OdsValidationValueKind.List
            && validation.ParsedCondition.ListValues.SequenceEqual(new[] { "New", "On \"Hold\"", "Done" }));
        Assert.NotNull(toOds.Value.GetSheet("Data")!.Cell(1, 1).ValidationName);
        Assert.NotNull(toOds.Value.GetSheet("Data")!.Cell(2, 2).ValidationName);
        Assert.Contains(toOds.Report.Mappings, mapping =>
            mapping.Feature == "validations" && mapping.Status == OdfConversionMappingStatus.Converted && mapping.Count == 2);

        using ExcelDocument roundTrip = toOds.Value.ToExcelDocumentResult().Value;
        IReadOnlyList<ExcelDataValidationInfo> validations = roundTrip.Sheets.Single().GetDataValidations();
        Assert.Equal(2, validations.Count);
        Assert.Contains(validations, validation => validation.Type == "whole"
            && validation.Formula1 == "1" && validation.Formula2 == "10"
            && validation.PromptTitle == "Allowed values" && validation.ShowInputMessage
            && validation.ErrorTitle == "Outside range" && validation.ShowErrorMessage
            && validation.ErrorStyle == "warning");
        Assert.Contains(validations, validation => validation.Type == "list" && validation.Formula1 == "\"New,On \"\"Hold\"\",Done\"");
    }

    [Fact]
    public void LegacyExcelCommentsRoundTripThroughOdsAnnotations() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Data");
        sheet.SetComment("C4", "Review  this\tvalue\nbefore release", "Alice");

        OdfConversionResult<OdsDocument> toOds = source.ToOpenDocumentResult();
        OdsAnnotation annotation = Assert.Single(toOds.Value.GetSheet("Data")!.Cell(3, 2).Annotations);
        Assert.Equal("Alice", annotation.Creator);
        Assert.Equal("Review  this\tvalue\nbefore release", annotation.Text);
        Assert.Contains(toOds.Report.Mappings, mapping =>
            mapping.Feature == "comments" && mapping.Status == OdfConversionMappingStatus.Converted);

        using ExcelDocument roundTrip = toOds.Value.ToExcelDocumentResult().Value;
        ExcelCommentInfo comment = Assert.Single(roundTrip.Sheets.Single().GetComments());
        Assert.Equal("C4", comment.CellReference);
        Assert.Equal("Alice", comment.Author);
        Assert.Equal("Review  this\tvalue\nbefore release", comment.Text);
    }

    [Fact]
    public void ThreadedExcelCommentsBecomeExplicitlyApproximatedOdsAnnotations() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Data");
        DateTime timestamp = new DateTime(2026, 8, 9, 8, 30, 0, DateTimeKind.Utc);
        ExcelThreadedCommentResult root = sheet.AddThreadedComment(new ExcelThreadedCommentOptions {
            Address = "A1",
            Text = "Root",
            Author = "Alice",
            Date = timestamp,
            Done = true
        });
        sheet.AddThreadedComment(new ExcelThreadedCommentOptions {
            Address = "A1",
            Text = "Reply",
            Author = "Bob",
            ParentId = root.Id,
            Date = timestamp.AddMinutes(5)
        });

        OdfConversionResult<OdsDocument> conversion = source.ToOpenDocumentResult();
        OdsAnnotation annotation = Assert.Single(conversion.Value.GetSheet("Data")!.Cell(0, 0).Annotations);

        Assert.Equal("Alice", annotation.Creator);
        Assert.Contains("Root", annotation.Text, StringComparison.Ordinal);
        Assert.Contains("Reply", annotation.Text, StringComparison.Ordinal);
        Assert.Contains("Bob", annotation.Text, StringComparison.Ordinal);
        Assert.Contains("resolved", annotation.Text, StringComparison.Ordinal);
        Assert.Contains(conversion.Report.Mappings, mapping =>
            mapping.Feature == "threaded-comments"
            && mapping.Status == OdfConversionMappingStatus.Approximated
            && mapping.Count == 2);
    }

    [Fact]
    public void Conversion_Loss_Policy_Distinguishes_Approximation_From_Skipped_Or_Unsupported_Content() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Data");
        sheet.CellAt(1, 1).SetValue(1D);
        sheet.CellAt(1, 2).SetFormula("A1+1");

        OdfConversionResult<OdsDocument> accepted = source.ToOpenDocumentResult(
            new ExcelOpenDocumentConversionOptions {
                LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported
            });

        Assert.True(accepted.HasLoss);
        Assert.False(accepted.Report.HasSkippedOrUnsupported);

        OdfConversionLossException exception = Assert.Throws<OdfConversionLossException>(() =>
            source.ToOpenDocumentResult(new ExcelOpenDocumentConversionOptions {
                LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss
            }));
        Assert.Contains(exception.Report.Mappings, mapping =>
            mapping.Feature == "formulas" && mapping.Status == OdfConversionMappingStatus.Approximated);
    }

    [Fact]
    public void Quoted_Percent_Is_Not_Converted_To_Percentage_Scaling() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelCell cell = source.AddWorksheet("Data").CellAt(1, 1);
        cell.SetValue(0.5D);
        cell.SetNumberFormat("0.00\"%\"");

        OdfConversionResult<OdsDocument> conversion = source.ToOpenDocumentResult();
        OdsDataStyle style = Assert.Single(conversion.Value.DataStyles);

        Assert.Equal(OdsDataStyleKind.Number, style.Kind);
        Assert.Equal(2, style.DecimalPlaces);
        Assert.Contains(conversion.Report.Mappings, mapping =>
            mapping.Feature == "cell-format-details" && mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void Excel_Thousands_Scaling_Comma_Is_Not_Misreported_As_Grouping() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelCell cell = source.AddWorksheet("Data").CellAt(1, 1);
        cell.SetValue(12000D);
        cell.SetNumberFormat("#,##0,");

        OdfConversionResult<OdsDocument> conversion = source.ToOpenDocumentResult();

        Assert.Contains(conversion.Report.Mappings, mapping =>
            mapping.Feature == "cell-format-details"
            && mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() =>
            source.ToOpenDocumentResult(new ExcelOpenDocumentConversionOptions {
                LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported
            }));
    }

    [Fact]
    public void Percentage_Decimals_And_Grouping_RoundTrip_Through_Typed_Style() {
        OdsDocument source = OdsDocument.Create();
        OdsDataStyle style = source.AddPercentageStyle("Rate", 3, useGrouping: true);
        OdsCell cell = source.AddSheet("Data").Cell(0, 0);
        cell.SetPercentage(0.125M);
        cell.NumberFormatName = style.Name;

        using ExcelDocument target = source.ToExcelDocumentResult().Value;
        ExcelCellSnapshot converted = Assert.Single(target.CreateInspectionSnapshot().Worksheets.Single().Cells);

        Assert.Equal("#,##0.000%", converted.Style!.NumberFormatCode);
    }

    [Fact]
    public void Currency_Symbol_And_Decimals_RoundTrip_From_Ods() {
        OdsDocument source = OdsDocument.Create();
        OdsDataStyle style = source.AddCurrencyStyle("Amount", "EUR", 1, useGrouping: true);
        OdsCell cell = source.AddSheet("Data").Cell(0, 0);
        cell.SetCurrency(12.5M, "EUR");
        cell.NumberFormatName = style.Name;

        using ExcelDocument target = source.ToExcelDocumentResult().Value;
        ExcelCellSnapshot converted = Assert.Single(target.CreateInspectionSnapshot().Worksheets.Single().Cells);

        Assert.Equal("\"EUR\" #,##0.0", converted.Style!.NumberFormatCode);
    }
}
