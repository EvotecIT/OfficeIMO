using System;
using System.Collections.Generic;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Models.Excel;

/// <summary>Combines native workbook import policies with PowerShell worksheet selectors.</summary>
public sealed class ExcelMergeOptions {
    internal ExcelMergeOptions(ExcelWorkbookMergeOptions nativeOptions, OfficeSheetSelector[]? sheets) {
        NativeOptions = nativeOptions;
        Sheet = sheets == null ? null : Array.AsReadOnly((OfficeSheetSelector[])sheets.Clone());
    }
    /// <summary>Native policies captured independently when Merge-OfficeDocument starts.</summary>
    public ExcelWorkbookMergeOptions NativeOptions { get; }
    /// <summary>Worksheets selected in appended source files. Null imports the native selection; an empty selection imports none.</summary>
    public IReadOnlyList<OfficeSheetSelector>? Sheet { get; }
}
