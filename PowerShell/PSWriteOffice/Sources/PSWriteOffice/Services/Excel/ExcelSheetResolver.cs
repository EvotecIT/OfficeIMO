using System;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelSheetResolver
{
    /// <summary>Resolves a typed selection against its explicit, enclosing or native worksheet's workbook.</summary>
    public static ExcelSheet ResolveTarget(PSCmdlet cmdlet, ExcelDocument? document, OfficeSheetSelector? selection) {
        ExcelDocument workbook = document ?? OfficeDocumentContext.Find<ExcelDocument>() ?? ExcelDslContext.Current?.Document
            ?? (selection?.Sheet as ExcelSheet)?.Document ?? ExcelDslContext.Require(cmdlet).Document;
        if (selection != null) return selection.Resolve(workbook.Sheets, sheet => sheet.Name);
        return document != null ? Resolve(workbook, null, null) : ExcelDslContext.Require(cmdlet).RequireSheet();
    }

    /// <summary>Resolves an explicit worksheet, workbook selection, or the active PowerShell composition scope.</summary>
    public static ExcelSheet ResolveTarget(PSCmdlet cmdlet, ExcelSheet? worksheet, ExcelDocument? document, string? sheetName, int? sheetIndex)
    {
        if (worksheet != null && document != null)
        {
            throw new PSArgumentException("Use either -Worksheet or -Document, not both.");
        }

        if (sheetName != null && sheetIndex.HasValue)
        {
            throw new PSArgumentException("Use either -Sheet or -SheetIndex, not both.");
        }

        if (document == null && (sheetName != null || sheetIndex.HasValue))
        {
            throw new PSArgumentException("Use -Sheet or -SheetIndex together with -Document, or pipe a workbook to the command.");
        }

        return worksheet ?? (document != null
            ? Resolve(document, sheetName, sheetIndex)
            : ExcelDslContext.Require(cmdlet).RequireSheet());
    }

    public static ExcelSheet Resolve(ExcelDocument document, string? sheetName, int? sheetIndex)
    {
        if (document == null)
        {
            throw new ArgumentNullException(nameof(document));
        }

        if (sheetName != null)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new PSArgumentException("Sheet cannot be empty or whitespace. Omit -Sheet to use the default worksheet.", "Sheet");
            }

            return document[sheetName];
        }

        if (sheetIndex.HasValue)
        {
            if (sheetIndex.Value < 0 || sheetIndex.Value >= document.Sheets.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(sheetIndex), "SheetIndex is out of range.");
            }

            return document.Sheets[sheetIndex.Value];
        }

        if (document.Sheets.Count == 0)
        {
            throw new InvalidOperationException("Workbook contains no worksheets.");
        }

        return document.Sheets[0];
    }

    public static ExcelSheet? ResolveOptional(ExcelDocument document, string? sheetName, int? sheetIndex)
    {
        if (sheetName == null && !sheetIndex.HasValue)
        {
            return null;
        }

        return Resolve(document, sheetName, sheetIndex);
    }
}
