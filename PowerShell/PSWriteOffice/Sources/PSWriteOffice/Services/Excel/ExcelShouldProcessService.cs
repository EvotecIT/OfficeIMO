using System;
using System.Management.Automation;
using OfficeIMO.Excel;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelShouldProcessService
{
    public static bool ShouldProcessWorkbook(PSCmdlet cmdlet, ExcelDocument document, string? inputPath, string action)
    {
        ArgumentNullException.ThrowIfNull(cmdlet);

        var target = !string.IsNullOrWhiteSpace(document?.FilePath)
            ? document!.FilePath!
            : !string.IsNullOrWhiteSpace(inputPath)
                ? cmdlet.SessionState.Path.GetUnresolvedProviderPathFromPSPath(inputPath!)
                : "Excel workbook";

        return cmdlet.ShouldProcess(target, action);
    }

    public static bool ShouldProcessTarget(PSCmdlet cmdlet, string target, string action)
    {
        ArgumentNullException.ThrowIfNull(cmdlet);

        return cmdlet.ShouldProcess(string.IsNullOrWhiteSpace(target) ? "Excel workbook" : target, action);
    }
}
