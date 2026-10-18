using System;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Excel;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelWorkbookCommandService
{
    public static ExcelWorkbookCommandScope ResolveWorkbook(
        PSCmdlet cmdlet,
        string parameterSetName,
        string inputPath,
        ExcelDocument? document,
        bool readOnly)
    {
        if (string.Equals(parameterSetName, "Path", StringComparison.OrdinalIgnoreCase))
        {
            return OpenWorkbook(cmdlet, inputPath, readOnly);
        }

        if (string.Equals(parameterSetName, "Document", StringComparison.OrdinalIgnoreCase))
        {
            return new ExcelWorkbookCommandScope(document ?? throw new PSArgumentException("Provide an Excel document."), ownsDocument: false);
        }

        return new ExcelWorkbookCommandScope(ExcelDslContext.Require(cmdlet).Document, ownsDocument: false);
    }

    public static ExcelWorkbookCommandScope OpenWorkbook(PSCmdlet cmdlet, string inputPath, bool readOnly)
    {
        var resolvedPath = cmdlet.SessionState.Path.GetUnresolvedProviderPathFromPSPath(inputPath);
        return new ExcelWorkbookCommandScope(ExcelDocumentService.LoadDocument(resolvedPath, readOnly, autoSave: false), ownsDocument: true);
    }

    public static ExcelSheet ResolveSheet(
        PSCmdlet cmdlet,
        ExcelDocument document,
        string parameterSetName,
        string? sheetName,
        int? sheetIndex)
    {
        if (string.Equals(parameterSetName, "Context", StringComparison.OrdinalIgnoreCase)
            && string.IsNullOrWhiteSpace(sheetName)
            && !sheetIndex.HasValue)
        {
            return ExcelDslContext.Require(cmdlet).RequireSheet();
        }

        return ExcelSheetResolver.Resolve(document, sheetName, sheetIndex);
    }

    public static IEnumerable<ExcelSheet> ResolveSheets(
        PSCmdlet cmdlet,
        ExcelDocument document,
        string parameterSetName,
        string? sheetName,
        int? sheetIndex)
    {
        if (string.Equals(parameterSetName, "Context", StringComparison.OrdinalIgnoreCase)
            && string.IsNullOrWhiteSpace(sheetName)
            && !sheetIndex.HasValue)
        {
            yield return ExcelDslContext.Require(cmdlet).RequireSheet();
            yield break;
        }

        if (!string.IsNullOrWhiteSpace(sheetName) || sheetIndex.HasValue)
        {
            yield return ExcelSheetResolver.Resolve(document, sheetName, sheetIndex);
            yield break;
        }

        foreach (var sheet in document.Sheets)
        {
            yield return sheet;
        }
    }

    public static ExcelWorkbookCommandScope ResolveSourceWorkbook(
        PSCmdlet cmdlet,
        ExcelDocument fallbackDocument,
        ExcelDocument? sourceDocument,
        string? sourcePath,
        bool readOnly)
    {
        ValidateSingleSource(sourceDocument, sourcePath);
        if (!string.IsNullOrWhiteSpace(sourcePath))
        {
            return OpenWorkbook(cmdlet, sourcePath!, readOnly);
        }

        return new ExcelWorkbookCommandScope(sourceDocument ?? fallbackDocument, ownsDocument: false);
    }

    public static string ResolveSheetNameOrCurrent(
        PSCmdlet cmdlet,
        ExcelDocument document,
        string parameterSetName,
        string? sheetName)
    {
        if (!string.IsNullOrWhiteSpace(sheetName))
        {
            return sheetName!;
        }

        if (string.Equals(parameterSetName, "Context", StringComparison.OrdinalIgnoreCase))
        {
            return ExcelDslContext.Require(cmdlet).RequireSheet().Name;
        }

        if (document.Sheets.Count == 0)
        {
            throw new InvalidOperationException("Workbook contains no worksheets.");
        }

        return document.Sheets[0].Name;
    }

    public static void ValidateSingleSource(ExcelDocument? sourceDocument, string? sourcePath)
    {
        if (sourceDocument != null && !string.IsNullOrWhiteSpace(sourcePath))
        {
            throw new PSArgumentException("Specify either -SourceDocument or -SourcePath, not both.");
        }
    }
}
