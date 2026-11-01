using System.Management.Automation;
using OfficeIMO.Excel;
using OfficeIMO.Excel.GoogleSheets;
using OfficeIMO.GoogleWorkspace;

namespace PSWriteOffice.Tests;

public sealed partial class ConversionCmdletWorkflowTests {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void GoogleSheetsFactoryFeedsNativeFormulaPropertyAndProtectionRequestsWithoutASession(bool domainUsersCanEdit) {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Forecast");
        sheet.CellValue(1, 1, 3);
        sheet.CellValue(2, 1, 7);
        sheet.CellFormula(3, 1, "MYTOTAL(A1:A2)");
        sheet.Protect();
        using var runspace = OpenRunspace();
        runspace.SessionStateProxy.SetVariable("document", document);
        runspace.SessionStateProxy.SetVariable("domainUsersCanEdit", domainUsersCanEdit);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("""
            New-OfficeExcelGoogleSheetsOptions
            $options = New-OfficeExcelGoogleSheetsOptions -Title 'Forecast review' -FolderId folder -DriveId drive `
                -SpreadsheetId existing -ExpectedDriveVersion 42 -UnsupportedFormula Error -FunctionMapping @{ MYTOTAL = 'SUM' } `
                -Locale pl_PL -TimeZone Europe/Warsaw -RecalculationInterval Hour `
                -Editor editor@example.invalid -DomainUsersCanEdit:$domainUsersCanEdit
            Export-OfficeExcelGoogleSpreadsheet -Document $document -Options $options -AsBatch -ErrorAction Stop
            $options
            """);
        var output = shell.Invoke();
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
        Assert.Equal(3, output.Count);
        var defaults = Assert.IsType<GoogleSheetsSaveOptions>(output[0].BaseObject);
        Assert.Equal(GoogleSheetsReplaceConflictMode.RequireMatchingDriveVersion, defaults.Replace.ConflictMode);
        Assert.Equal(GoogleSheetsUnsupportedFormulaMode.PreserveWithWarning, defaults.Formulas.UnsupportedFormulaMode);
        Assert.Equal(UnsupportedFeatureMode.Error, defaults.UnsupportedFeatures.Charts);
        Assert.Equal(UnsupportedFeatureMode.WarnAndSkip, defaults.UnsupportedFeatures.PrintLayout);
        var batch = Assert.IsType<GoogleSheetsBatch>(output[1].BaseObject);
        Assert.Equal("Forecast review", batch.Title);
        var properties = Assert.Single(batch.Requests.OfType<GoogleSheetsUpdateSpreadsheetPropertiesRequest>());
        Assert.Equal("pl_PL", properties.Locale);
        Assert.Equal("Europe/Warsaw", properties.TimeZone);
        Assert.Equal(GoogleSheetsRecalculationInterval.Hour, properties.RecalculationInterval);
        var formula = Assert.Single(batch.Requests.OfType<GoogleSheetsUpdateCellsRequest>().SelectMany(request => request.Cells),
            cell => cell.Value.Kind == GoogleSheetsCellValueKind.Formula);
        Assert.Equal("SUM(A1:A2)", Assert.IsType<string>(formula.Value.Value).TrimStart('='));
        var protection = Assert.Single(batch.Requests.OfType<GoogleSheetsAddProtectedRangeRequest>());
        Assert.Equal(domainUsersCanEdit, protection.DomainUsersCanEdit);
        Assert.Equal("editor@example.invalid", Assert.Single(protection.EditorEmailAddresses));
        var options = Assert.IsType<GoogleSheetsSaveOptions>(output[2].BaseObject);
        Assert.Equal("existing", options.Location.ExistingFileId);
        Assert.Equal("folder", options.Location.FolderId);
        Assert.Equal("drive", options.Location.DriveId);
        Assert.Equal(42, options.Replace.ExpectedDriveVersion);
        Assert.True(sheet.IsProtected);
    }
}
