using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Xlsb.Biff12;
using System.IO.Compression;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Data validations written by the native XLSB writer, read back and rewritten in loaded XLSB files.</summary>
public sealed class ExcelXlsbDataValidationTests {
    [Theory]
    [InlineData("explicit-list", "list", null, "\"Red,Green,Blue\"", null)]
    [InlineData("same-sheet-list", "list", null, "$H$1:$H$3", null)]
    [InlineData("other-sheet-list", "list", null, "'Lists'!$A$1:$A$3", null)]
    [InlineData("named-list", "list", null, "Colors", null)]
    [InlineData("whole-between", "whole", "between", "1", "10")]
    [InlineData("decimal-greater", "decimal", "greaterThan", "1.5", null)]
    [InlineData("date-before", "date", "lessThanOrEqual", "46053", null)]
    [InlineData("time-between", "time", "between", "0.5", "0.75")]
    [InlineData("text-length", "textLength", "notEqual", "5", null)]
    [InlineData("custom-relative", "custom", null, "ISNUMBER(B2)", null)]
    public void Data_validation_survives_native_xlsb_save(string shape, string type, string? @operator, string formula1, string? formula2) {
        byte[] xlsx = CreateXlsx(document => {
            ExcelSheet lists = document.AddSheet("Lists");
            lists.Cell(1, 1).SetValue("Red");
            ExcelSheet sheet = document.AddSheet("Data");
            document.SetNamedRange("Colors", "'Lists'!$A$1:$A$3", save: false);
            ExcelRange range = sheet.Range("B2:B5");
            switch (shape) {
                case "explicit-list": range.Validation().List("Red", "Green", "Blue"); break;
                case "same-sheet-list": range.Validation().ListRange("$H$1:$H$3"); break;
                case "other-sheet-list": range.Validation().ListRange("$A$1:$A$3", "Lists"); break;
                case "named-list": range.Validation().ListNamedRange("Colors"); break;
                case "whole-between": range.Validation().WholeNumber(ExcelDataValidationOperator.Between, 1, 10, errorTitle: "Range", errorMessage: "1 to 10"); break;
                case "decimal-greater": range.Validation().Decimal(ExcelDataValidationOperator.GreaterThan, 1.5); break;
                case "date-before": range.Validation().Date(ExcelDataValidationOperator.LessThanOrEqual, new DateTime(2026, 1, 31)); break;
                case "time-between": range.Validation().Time(ExcelDataValidationOperator.Between, TimeSpan.FromHours(12), TimeSpan.FromHours(18)); break;
                case "text-length": range.Validation().TextLength(ExcelDataValidationOperator.NotEqual, 5); break;
                case "custom-relative": range.Validation().CustomFormula("=ISNUMBER(B2)"); break;
            }
        });

        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(SaveAsXlsb(xlsx)));
        ExcelDataValidationInfo validation = Assert.Single(reloaded["Data"].GetDataValidations());
        Assert.Equal("B2:B5", validation.Range);
        Assert.Equal(type, validation.Type);
        Assert.Equal(@operator, validation.Operator);
        Assert.Equal(formula1, validation.Formula1);
        Assert.Equal(formula2, validation.Formula2);
        Assert.True(validation.AllowBlank);
        Assert.False(validation.SuppressDropDown);
        if (shape == "whole-between") {
            Assert.True(validation.ShowErrorMessage);
            Assert.Equal("Range", validation.ErrorTitle);
            Assert.Equal("1 to 10", validation.Error);
        }
    }

    [Fact]
    public void Data_validation_messages_flags_and_ranges_survive_native_xlsb_save() {
        byte[] xlsx = CreateXlsx(document => {
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Range("A1:A3").Validation().List(new[] { "Yes", "No" }, allowBlank: false);
            sheet.SetDataValidationMessages("A1", new ExcelDataValidationMessageOptions {
                PromptTitle = "Answer", Prompt = "Pick yes or no", ErrorTitle = "Invalid", Error = "Only yes or no",
                ErrorStyle = ExcelDataValidationErrorStyle.Warning, SuppressDropDown = true
            });
            sheet.Range("D4").Validation().CustomFormula("=LEN($D$4)<10");
            DataValidation custom = sheet.WorksheetPart.Worksheet!.GetFirstChild<DataValidations>()!.Elements<DataValidation>().Last();
            custom.SequenceOfReferences!.InnerText = "D4 F6:G7";
            custom.ImeMode = DataValidationImeModeValues.Hiragana;
            custom.ErrorStyle = DataValidationErrorStyleValues.Information;
        });

        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(SaveAsXlsb(xlsx)));
        ExcelDataValidationInfo[] validations = reloaded["Data"].GetDataValidations().ToArray();
        Assert.Equal(2, validations.Length);
        ExcelDataValidationInfo list = validations[0];
        Assert.Equal("A1:A3", list.Range);
        Assert.False(list.AllowBlank);
        Assert.True(list.SuppressDropDown);
        Assert.Equal("warning", list.ErrorStyle);
        Assert.True(list.ShowInputMessage);
        Assert.True(list.ShowErrorMessage);
        Assert.Equal("Answer", list.PromptTitle);
        Assert.Equal("Pick yes or no", list.Prompt);
        Assert.Equal("Invalid", list.ErrorTitle);
        Assert.Equal("Only yes or no", list.Error);
        ExcelDataValidationInfo custom = validations[1];
        Assert.Equal("D4 F6:G7", custom.Range);
        Assert.Equal("LEN($D$4)<10", custom.Formula1);
        Assert.Equal("information", custom.ErrorStyle);
        Assert.False(custom.ShowInputMessage);
        DataValidation customElement = reloaded["Data"].WorksheetPart.Worksheet!.GetFirstChild<DataValidations>()!.Elements<DataValidation>().Last();
        Assert.Equal(DataValidationImeModeValues.Hiragana, customElement.ImeMode!.Value);
    }

    [Fact]
    public void Editing_data_validations_of_a_loaded_xlsb_keeps_the_other_native_records() {
        byte[] source = SaveAsXlsb(CreateXlsx(document => {
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("x");
            sheet.Range("A2:A4").Validation().List("a", "b");
            sheet.AddManualRowPageBreak(10, save: false);
        }));

        byte[] pageBreakEdit;
        using (ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(source))) {
            loaded["Data"].AddManualRowPageBreak(20, save: false);
            pageBreakEdit = loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
        }
        // Unchanged validations are copied record for record.
        Assert.Equal(ReadRecords(source, 573, 574), ReadRecords(pageBreakEdit, 573, 574));

        byte[] validationEdit;
        using (ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(source))) {
            loaded["Data"].Range("C1").Validation().WholeNumber(ExcelDataValidationOperator.GreaterThan, 0);
            loaded["Data"].UpdateDataValidationMessages("A2", options => options.Prompt = "Choose");
            validationEdit = loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
        }

        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(validationEdit));
        ExcelDataValidationInfo[] validations = reloaded["Data"].GetDataValidations().ToArray();
        Assert.Equal(2, validations.Length);
        Assert.Equal("\"a,b\"", validations[0].Formula1);
        Assert.Equal("Choose", validations[0].Prompt);
        Assert.Equal("C1", validations[1].Range);
        Assert.Equal("greaterThan", validations[1].Operator);
        Assert.Equal("0", validations[1].Formula1);
        Assert.Equal(new[] { 10 }, reloaded["Data"].GetManualRowPageBreaks());
        Assert.Equal(ReadRecords(source, 392, 393), ReadRecords(validationEdit, 392, 393));
    }

    [Theory]
    [InlineData("single-cell-other-sheet")]
    [InlineData("long-error-title")]
    public void Data_validations_xlsb_cannot_hold_are_refused(string shape) {
        byte[] xlsx = CreateXlsx(document => {
            document.AddSheet("Lists");
            ExcelSheet sheet = document.AddSheet("Data");
            if (shape == "single-cell-other-sheet") {
                // DVParsedFormula forbids PtgRef3d (MS-XLSB 2.5.98.8).
                sheet.Range("A1").Validation().ListRange("$A$1", "Lists");
            } else {
                // DValStrings limits the error title to 32 characters (MS-XLSB 2.5.37).
                sheet.Range("A1").Validation().WholeNumber(ExcelDataValidationOperator.Equal, 1, errorTitle: new string('t', 33));
            }
        });

        using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(xlsx));
        Assert.Throws<NotSupportedException>(() => loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb }));
    }

    private static byte[] CreateXlsx(Action<ExcelDocument> build) {
        using ExcelDocument document = ExcelDocument.Create();
        build(document);
        return document.ToBytes();
    }

    private static byte[] SaveAsXlsb(byte[] xlsx) {
        using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(xlsx));
        return loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
    }

    private static string[] ReadRecords(byte[] xlsb, int beginType, int endType) {
        using var archive = new ZipArchive(new MemoryStream(xlsb), ZipArchiveMode.Read);
        using var part = new MemoryStream();
        using (Stream entry = archive.GetEntry("xl/worksheets/sheet1.bin")!.Open()) entry.CopyTo(part);
        part.Position = 0;
        var records = XlsbRecordReader.ReadAll(part).ToList();
        int begin = records.FindIndex(record => record.Type == beginType);
        int end = records.FindIndex(record => record.Type == endType);
        Assert.True(begin >= 0 && end > begin);
        return records.Skip(begin).Take(end - begin + 1).Select(record => record.Type + ":" + Convert.ToBase64String(record.Data)).ToArray();
    }
}
