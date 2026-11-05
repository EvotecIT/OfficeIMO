using OfficeIMO.Excel;
using System.Management.Automation;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelTableStyleOptionService
{
    public static bool? IsSwitchPresent(PSCmdlet cmdlet, string parameterName, SwitchParameter value)
    {
        return cmdlet.MyInvocation.BoundParameters.ContainsKey(parameterName) ? value.IsPresent : (bool?)null;
    }

    public static void Apply(
        ExcelSheet sheet,
        string? tableOrRange,
        ExcelTableStyle style,
        bool? showFirstColumn,
        bool? showLastColumn,
        bool? noRowStripes,
        bool? showColumnStripes)
    {
        if (string.IsNullOrWhiteSpace(tableOrRange))
        {
            return;
        }

        if (!showFirstColumn.HasValue && !showLastColumn.HasValue && !noRowStripes.HasValue && !showColumnStripes.HasValue)
        {
            return;
        }

        sheet.SetTableStyle(
            tableOrRange!,
            style,
            showFirstColumn: showFirstColumn,
            showLastColumn: showLastColumn,
            showRowStripes: noRowStripes.HasValue ? !noRowStripes.Value : (bool?)null,
            showColumnStripes: showColumnStripes);
    }
}
