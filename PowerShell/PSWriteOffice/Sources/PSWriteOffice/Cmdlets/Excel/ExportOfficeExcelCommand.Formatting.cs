using System.Collections;
using System.Globalization;
using System.Management.Automation;
using System.Security;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;
using System.Data;

namespace PSWriteOffice.Cmdlets.Excel;

public sealed partial class ExportOfficeExcelCommand {
    private ExcelColumnFormatPlan? BuildColumnFormatPlan()
        => ExcelColumnFormatPlanService.Build(ColumnFormat, null, null, null, null, null, null, null, Decimals, Culture?.Name);

    private IReadOnlyList<ExcelColumnFormatResult> ApplyColumnFormatPlan(
        ExcelSheet sheet,
        ExcelColumnFormatPlan? plan,
        int? headerRow = null,
        bool throwOnMissing = true,
        bool requireExplicitHeaderRow = false)
    {
        if (plan == null)
        {
            return Array.Empty<ExcelColumnFormatResult>();
        }

        if (requireExplicitHeaderRow && !headerRow.HasValue)
        {
            var headers = string.Join(", ", GetColumnFormatHeaders(plan));
            WriteVerbose($"Column format headers were not applied on worksheet '{sheet.Name}' because no header row was emitted.");
            if (throwOnMissing && !IgnoreMissingColumnFormat.IsPresent)
            {
                throw new PSArgumentException($"Column format headers were not found on worksheet '{sheet.Name}': {headers}. Export-time column formats require a header row; remove -NoHeader or use -IgnoreMissingColumnFormat for optional columns.");
            }

            return Array.Empty<ExcelColumnFormatResult>();
        }

        var results = sheet.ApplyColumnFormatPlan(
            plan,
            includeHeader: false,
            autoFit: AutoFitFormattedColumn.IsPresent,
            headerRow: headerRow);
        var missing = results.Where(static result => !result.Applied).ToArray();
        foreach (var result in missing)
        {
            WriteVerbose(result.Warning);
        }

        if (missing.Length > 0 && throwOnMissing && !IgnoreMissingColumnFormat.IsPresent)
        {
            var headers = string.Join(", ", missing.Select(static result => result.Header));
            throw new PSArgumentException($"Column format headers were not found on worksheet '{sheet.Name}': {headers}. Use -IgnoreMissingColumnFormat for optional columns.");
        }

        return results;
    }

    private static string[] GetColumnFormatHeaders(ExcelColumnFormatPlan plan)
        => plan.Rules
            .Select(static rule => rule.Header.Trim())
            .Where(static header => header.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

    private bool ColumnFormatRequiresMaterializedCells(ExcelColumnFormatPlan? plan)
    {
        return plan?.Rules.Any(static rule => rule.IncludeHeader) ?? false;
    }

    private void ThrowForDataSetColumnFormatsMissingEverywhere(ExcelColumnFormatPlan? plan, HashSet<string> formattedHeaders)
    {
        if (plan == null || IgnoreMissingColumnFormat.IsPresent)
        {
            return;
        }

        var missingEverywhere = plan.Rules
            .Select(static rule => rule.Header)
            .Where(header => !formattedHeaders.Contains(header))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        if (missingEverywhere.Length == 0)
        {
            return;
        }

        throw new PSArgumentException($"Column format headers were not found in any DataSet worksheet: {string.Join(", ", missingEverywhere)}. Use -IgnoreMissingColumnFormat for optional columns.");
    }

    private void ApplyTableStyleOptions(ExcelSheet sheet, string? tableOrRange, ExcelTableStyle style)
    {
        ExcelTableStyleOptionService.Apply(
            sheet,
            tableOrRange,
            style,
            ExcelTableStyleOptionService.IsSwitchPresent(this, nameof(ShowFirstColumn), ShowFirstColumn),
            ExcelTableStyleOptionService.IsSwitchPresent(this, nameof(ShowLastColumn), ShowLastColumn),
            ExcelTableStyleOptionService.IsSwitchPresent(this, nameof(NoRowStripes), NoRowStripes),
            ExcelTableStyleOptionService.IsSwitchPresent(this, nameof(ShowColumnStripes), ShowColumnStripes));
    }

    private int? ResolveColumnFormatHeaderRow(
        ExcelDocument document,
        ExcelSheet sheet,
        ExcelColumnFormatPlan? plan,
        bool appendToExistingSheet,
        string? appendTableName,
        bool includeHeaders,
        int dataStartRow)
    {
        if (includeHeaders)
        {
            return dataStartRow;
        }

        if (!appendToExistingSheet || plan == null)
        {
            return null;
        }

        var tableHeaderRow = ResolveExistingTableHeaderRow(document, sheet, appendTableName);
        if (NoHeader.IsPresent)
        {
            return tableHeaderRow;
        }

        return tableHeaderRow ?? FindExistingHeaderRow(sheet, plan, allowPartialMatch: IgnoreMissingColumnFormat.IsPresent);
    }

    private int? ResolveExistingTableHeaderRow(ExcelDocument document, ExcelSheet sheet, string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName))
        {
            return null;
        }

        var table = document.GetTables().FirstOrDefault(candidate =>
            string.Equals(candidate.SheetName, sheet.Name, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(candidate.Name, tableName, StringComparison.OrdinalIgnoreCase));
        if (table != null && table.HasHeaderRow && !string.IsNullOrWhiteSpace(table.Range) &&
            ExcelA1.TryParseRange(table.Range, out var tableStartRow, out _, out _, out _))
        {
            return Math.Max(1, tableStartRow);
        }

        return null;
    }

    private static int? FindExistingHeaderRow(ExcelSheet sheet, ExcelColumnFormatPlan plan, bool allowPartialMatch)
    {
        var requestedHeaders = GetColumnFormatHeaders(plan);
        if (requestedHeaders.Length == 0)
        {
            return null;
        }

        var usedRange = sheet.UsedRange.Address;
        if (!ExcelA1.TryParseRange(usedRange, out var firstRow, out var firstColumn, out var lastRow, out var lastColumn))
        {
            var (row, column) = ExcelA1.ParseCellRef(usedRange);
            firstRow = lastRow = row;
            firstColumn = lastColumn = column;
        }

        var headerRow = Math.Max(1, firstRow);
        var rowHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var column = Math.Max(1, firstColumn); column <= lastColumn; column++)
        {
            if (sheet.TryGetCellText(headerRow, column, out var text) && !string.IsNullOrWhiteSpace(text))
            {
                rowHeaders.Add(text.Trim());
            }
        }

        var hasRequestedHeader = allowPartialMatch
            ? requestedHeaders.Any(rowHeaders.Contains)
            : requestedHeaders.All(rowHeaders.Contains);

        return hasRequestedHeader ? headerRow : null;
    }

    private int ResolveFrozenTopRows(ExcelDocument document, ExcelSheet sheet, bool appendToExistingSheet, string? appendTableName, int dataStartRow, bool includeHeaders)
    {
        if (!FreezeTopRow.IsPresent)
        {
            return 0;
        }

        if (!appendToExistingSheet)
        {
            return Math.Max(1, includeHeaders ? dataStartRow : StartRow);
        }

        if (!string.IsNullOrWhiteSpace(appendTableName))
        {
            var table = document.GetTables().FirstOrDefault(candidate =>
                string.Equals(candidate.SheetName, sheet.Name, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(candidate.Name, appendTableName, StringComparison.OrdinalIgnoreCase));
            if (table != null && !string.IsNullOrWhiteSpace(table.Range) &&
                ExcelA1.TryParseRange(table.Range, out var tableStartRow, out _, out _, out _))
            {
                return Math.Max(1, tableStartRow);
            }
        }

        return Math.Max(1, StartRow);
    }


}
