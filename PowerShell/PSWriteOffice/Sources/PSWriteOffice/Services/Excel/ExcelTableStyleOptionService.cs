using OfficeIMO.Excel;
using System.Management.Automation;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelTableStyleOptionService
{
    public static bool IsSwitchPresent(PSCmdlet cmdlet, string parameterName, SwitchParameter value)
    {
        return value.IsPresent;
    }

    public static void Apply(
        ExcelSheet sheet,
        string? tableOrRange,
        ExcelTableStyle style,
        bool showFirstColumn,
        bool showLastColumn,
        bool noRowStripes,
        bool showColumnStripes)
    {
        if (string.IsNullOrWhiteSpace(tableOrRange))
        {
            return;
        }

        if (!showFirstColumn && !showLastColumn && !noRowStripes && !showColumnStripes)
        {
            return;
        }

        sheet.SetTableStyle(
            tableOrRange!,
            style,
            showFirstColumn: showFirstColumn ? true : null,
            showLastColumn: showLastColumn ? true : null,
            showRowStripes: noRowStripes ? false : null,
            showColumnStripes: showColumnStripes ? true : null);
    }
}
