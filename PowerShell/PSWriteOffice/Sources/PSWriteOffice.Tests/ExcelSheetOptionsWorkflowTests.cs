using System.Management.Automation;
using System.Management.Automation.Runspaces;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using PSWriteOffice.Cmdlets.Excel;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class ExcelSheetOptionsWorkflowTests {
    [Theory]
    [InlineData("2", "Lookups", true)]
    [InlineData("'2'", "2", true)]
    [InlineData("$selected", "Lookups", true)]
    [InlineData("$selected", "Lookups", false)]
    public void SelectorsEditOnlyTheirNativeSheetAndPersistVeryHiddenThemeColor(string selector, string expected, bool explicitDocument) {
        using var book = ExcelDocument.Create();
        var numericName = book.AddSheet("2");
        var lookup = book.AddSheet("Lookups");
        numericName.CellValue(1, 1, "Keep first data");
        lookup.CellValue(1, 1, "Keep lookup data");
        var other = expected == "2" ? lookup : numericName;
        book.SetActiveWorksheet(other);
        using var runspace = Open(book);
        runspace.SessionStateProxy.SetVariable("selected", lookup);
        var output = Invoke(runspace, "Set-OfficeExcelSheet " + (explicitDocument ? "-Document $book " : string.Empty)
            + "-Sheet " + selector + " -Visibility VeryHidden -TabColor Accent1 -PassThru");
        Assert.Same(book[expected], Assert.Single(output).BaseObject);
        Assert.False(other.Hidden);
        using var bytes = Save(book);
        using var saved = ExcelDocument.Load(bytes);
        Assert.True(saved[expected].VeryHidden);
        Assert.Equal("Keep first data", saved["2"].CellAt(1, 1).GetValue<string>());
        Assert.Equal("Keep lookup data", saved["Lookups"].CellAt(1, 1).GetValue<string>());
        bytes.Position = 0;
        using var package = SpreadsheetDocument.Open(bytes, false);
        var nativeSheet = package.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().Single(sheet => sheet.Name!.Value == expected);
        var part = (WorksheetPart)package.WorkbookPart!.GetPartById(nativeSheet.Id!.Value!);
        Assert.Equal(4U, part.Worksheet!.GetFirstChild<SheetProperties>()!.TabColor!.Theme!.Value);
    }

    [Fact]
    public void ShowingActivatingAndClearingKeepUnboundPropertiesAndAreQuietByDefault() {
        using var book = ExcelDocument.Create();
        book.AddSheet("Summary");
        var lookup = book.AddSheet("Lookups");
        lookup.SetVeryHidden(true);
        lookup.SetTabColor("#336699");
        using var runspace = Open(book);
        Assert.Empty(Invoke(runspace, "Set-OfficeExcelSheet -Document $book -Sheet Lookups -Active:$false -NoTabColor:$false"));
        Assert.True(book["Lookups"].VeryHidden);
        Assert.Equal(0, book.ActiveWorksheetIndex);
        using (var unchanged = Save(book))
        using (var previousPackage = SpreadsheetDocument.Open(unchanged, false)) {
            var previousSheet = previousPackage.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().Single(sheet => sheet.Name!.Value == "Lookups");
            Assert.Equal("FF336699", ((WorksheetPart)previousPackage.WorkbookPart!.GetPartById(previousSheet.Id!.Value!)).Worksheet!
                .GetFirstChild<SheetProperties>()?.TabColor?.Rgb?.Value);
        }
        Assert.Empty(Invoke(runspace, "$book | Set-OfficeExcelSheet -Sheet 2 -NoTabColor"));
        Assert.True(book["Lookups"].VeryHidden);
        Assert.Equal(0, book.ActiveWorksheetIndex);
        Assert.Empty(Invoke(runspace, "$book | Set-OfficeExcelSheet -Sheet 2 -Visibility Visible -Active"));
        using var bytes = Save(book);
        using var saved = ExcelDocument.Load(bytes);
        Assert.False(saved["Lookups"].Hidden);
        Assert.Equal(1, saved.ActiveWorksheetIndex);
        bytes.Position = 0;
        using var package = SpreadsheetDocument.Open(bytes, false);
        var native = package.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().Single(sheet => sheet.Name!.Value == "Lookups");
        Assert.Null(((WorksheetPart)package.WorkbookPart!.GetPartById(native.Id!.Value!)).Worksheet!.GetFirstChild<SheetProperties>()?.TabColor);
    }

    [Theory]
    [InlineData("-Sheet 3 -TabColor Red")]
    [InlineData("-Sheet $foreign -TabColor Red")]
    [InlineData("-Sheet Lookups -Visibility Hidden -Active -TabColor Red")]
    [InlineData("-Sheet Lookups -Visibility 99 -TabColor Red")]
    [InlineData("-Sheet Lookups -TabColor Red -NoTabColor")]
    public void InvalidSelectionsOrPropertyCombinationsLeaveTheWorkbookUnchanged(string arguments) {
        using var book = ExcelDocument.Create();
        book.AddSheet("Summary");
        var lookup = book.AddSheet("Lookups");
        lookup.SetTabColor("#336699");
        using var foreignBook = ExcelDocument.Create();
        using var runspace = Open(book);
        runspace.SessionStateProxy.SetVariable("foreign", foreignBook.AddSheet("Lookups"));
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("Set-OfficeExcelSheet -Document $book " + arguments);
        var error = Record.Exception(() => shell.Invoke());
        Assert.True(error != null || shell.HadErrors);
        Assert.False(lookup.Hidden);
        Assert.Equal(0, book.ActiveWorksheetIndex);
        using var bytes = Save(book);
        using var package = SpreadsheetDocument.Open(bytes, false);
        var native = package.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().Single(sheet => sheet.Name!.Value == "Lookups");
        Assert.Equal("FF336699", ((WorksheetPart)package.WorkbookPart!.GetPartById(native.Id!.Value!)).Worksheet!.GetFirstChild<SheetProperties>()!.TabColor!.Rgb!.Value);
    }

    private static Runspace Open(ExcelDocument book) {
        var state = InitialSessionState.CreateDefault();
        state.Commands.Add(new SessionStateCmdletEntry("Set-OfficeExcelSheet", typeof(SetOfficeExcelSheetCommand), null));
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        runspace.SessionStateProxy.SetVariable("book", book);
        return runspace;
    }

    private static System.Collections.ObjectModel.Collection<PSObject> Invoke(Runspace runspace, string script) {
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript(script);
        var output = shell.Invoke();
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
        return output;
    }

    private static MemoryStream Save(ExcelDocument book) {
        var bytes = new MemoryStream();
        book.Save(bytes);
        bytes.Position = 0;
        return bytes;
    }
}
