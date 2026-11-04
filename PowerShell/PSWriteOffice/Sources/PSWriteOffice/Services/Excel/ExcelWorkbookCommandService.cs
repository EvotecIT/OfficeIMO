using System;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;

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

    /// <summary>Opens one native workbook source, preserving ownership of live documents and literal FileInfo paths.</summary>
    public static ExcelWorkbookCommandScope ResolveSourceWorkbook(PSCmdlet cmdlet, ExcelDocument fallbackDocument, object? source)
    {
        source = OfficeArgumentValue.Unwrap(source);
        if (source == null) return new ExcelWorkbookCommandScope(fallbackDocument, ownsDocument: false);
        if (source is ExcelDocument document) return new ExcelWorkbookCommandScope(document, ownsDocument: false);
        string path = source is FileInfo file ? file.FullName : source as string
            ?? throw new PSArgumentException("Source accepts a workbook path, FileInfo or native ExcelDocument.");
        IReadOnlyList<string> paths = OfficePath.ResolveInput(cmdlet.SessionState, path, literal: source is FileInfo);
        if (paths.Count != 1) throw new PSArgumentException("Source must resolve to exactly one workbook.");
        return new ExcelWorkbookCommandScope(ExcelDocumentService.LoadDocument(paths[0], readOnly: true, autoSave: false), ownsDocument: true);
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


}
