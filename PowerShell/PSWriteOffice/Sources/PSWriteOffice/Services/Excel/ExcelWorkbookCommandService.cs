using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelWorkbookCommandService
{
    public static ExcelWorkbookCommandScope OpenWorkbook(PSCmdlet cmdlet, string inputPath, bool readOnly)
    {
        var paths = OfficePath.ResolveInput(cmdlet.SessionState, inputPath, literal: false);
        if (paths.Count != 1) throw new PSArgumentException("The workbook path must resolve to exactly one file.");
        string resolvedPath = paths[0];
        return new ExcelWorkbookCommandScope(ExcelDocumentService.LoadDocument(resolvedPath, readOnly, autoSave: false), ownsDocument: true);
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

}
