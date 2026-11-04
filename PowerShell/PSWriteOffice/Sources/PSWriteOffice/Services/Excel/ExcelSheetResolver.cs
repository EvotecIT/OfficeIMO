using System;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelSheetResolver
{
    /// <summary>Resolves mixed selectors and wildcard names once, retaining workbook order and native membership.</summary>
    public static IReadOnlyList<ExcelSheet> ResolveMany(ExcelDocument document, OfficeSheetSelector[]? selectors, ExcelSheet? defaultSheet = null) =>
        OfficeSheetSelector.ResolveMany(document, selectors, defaultSheet);

    /// <summary>Resolves a typed selection against its explicit, enclosing or native worksheet's workbook.</summary>
    public static ExcelSheet ResolveTarget(PSCmdlet cmdlet, ExcelDocument? document, OfficeSheetSelector? selection) {
        ExcelDocument workbook = document ?? OfficeDocumentContext.Find<ExcelDocument>() ?? ExcelDslContext.Current?.Document
            ?? (selection?.Sheet as ExcelSheet)?.Document ?? ExcelDslContext.Require(cmdlet).Document;
        if (selection != null) return selection.Resolve(workbook.Sheets, sheet => sheet.Name);
        return document != null ? Resolve(workbook) : ExcelDslContext.Require(cmdlet).RequireSheet();
    }

    /// <summary>Resolves a selector, or the first worksheet when omitted.</summary>
    public static ExcelSheet Resolve(ExcelDocument document, OfficeSheetSelector? selector = null) {
        ArgumentNullException.ThrowIfNull(document);
        if (selector != null) return selector.Resolve(document.Sheets, sheet => sheet.Name);
        if (document.Sheets.Count == 0) throw new InvalidOperationException("Workbook contains no worksheets.");
        return document.Sheets[0];
    }

}
