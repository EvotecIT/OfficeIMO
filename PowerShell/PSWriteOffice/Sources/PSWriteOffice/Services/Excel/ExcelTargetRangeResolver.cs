using System.Management.Automation;
using OfficeIMO.Excel;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelTargetRangeResolver
{
    public static string Resolve(
        ExcelSheet sheet,
        string? range,
        string? headerName,
        string? tableName,
        int headerRow,
        bool includeHeader,
        string? pivotTableName = null,
        bool pivotDataBody = true)
    {
        if (!string.IsNullOrWhiteSpace(pivotTableName))
        {
            return sheet.GetPivotTableRange(
                pivotTableName!.Trim(),
                pivotDataBody ? ExcelPivotRangeTarget.DataBody : ExcelPivotRangeTarget.WholeTable);
        }

        if (!string.IsNullOrWhiteSpace(range))
        {
            return range!.Trim();
        }

        if (string.IsNullOrWhiteSpace(headerName))
        {
            throw new PSArgumentException("Provide -Range or -HeaderName.");
        }

        if (sheet.TryGetColumnRangeByHeader(headerName!, tableName, headerRow, includeHeader, true, out string resolved))
            return resolved;
        throw new PSArgumentException($"Column header '{headerName}' was not found on worksheet '{sheet.Name}'.", nameof(headerName));
    }

    public static string? ResolveOptional(
        ExcelSheet sheet,
        string? range,
        string? headerName,
        string? tableName,
        int headerRow,
        bool includeHeader,
        string? pivotTableName = null,
        bool pivotDataBody = true)
    {
        if (!string.IsNullOrWhiteSpace(range) || !string.IsNullOrWhiteSpace(headerName) || !string.IsNullOrWhiteSpace(pivotTableName))
        {
            return Resolve(sheet, range, headerName, tableName, headerRow, includeHeader, pivotTableName, pivotDataBody);
        }

        return null;
    }
}
