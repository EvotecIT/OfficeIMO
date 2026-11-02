using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Models.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Services;

internal static partial class OfficeMergeDispatch {
    private static void MergeExcel(ExcelDocument target, IReadOnlyList<IDisposable> owned, int firstAppend, ExcelMergeOptions settings) {
        var sources = owned.Skip(firstAppend).Cast<ExcelDocument>().ToArray();
        OfficeSheetSelector[]? selectors = settings.Sheet?.ToArray();
        if (selectors != null) {
            foreach (OfficeSheetSelector selector in selectors) {
                if (selector.Sheet is not ExcelSheet sheet) continue;
                if (string.IsNullOrWhiteSpace(sheet.Document.FilePath) ||
                    !sources.Any(source => OfficePath.AreEquivalent(source.FilePath!, sheet.Document.FilePath!)))
                    throw new PSArgumentException("A sheet object must identify a worksheet in an appended saved source file.");
            }
        }
        var plans = new List<(ExcelDocument Source, ExcelWorkbookMergeOptions Options)>();
        foreach (ExcelDocument source in sources) {
            var native = settings.NativeOptions.Clone();
            if (selectors != null) {
                // Sheet objects address their own saved source. Names and positions apply to every appended source.
                var local = new List<OfficeSheetSelector>();
                foreach (OfficeSheetSelector selector in selectors) {
                    if (selector.Sheet is ExcelSheet sheet) {
                        if (OfficePath.AreEquivalent(source.FilePath!, sheet.Document.FilePath!))
                            local.Add(OfficeSheetSelector.From(source.GetSheet(sheet.Name)));
                    } else local.Add(selector);
                }
                if (local.Count == 0) continue;
                IReadOnlyList<ExcelSheet> selected = ExcelSheetResolver.ResolveMany(source, local.ToArray());
                // Native empty SheetNames means all sheets; a selector matching none must not import anything.
                if (selected.Count == 0) continue;
                native.SheetNames = selected.Select(sheet => sheet.Name).ToArray();
            } else if (native.SheetNames != null) {
                foreach (string name in native.SheetNames) source.GetSheet(name);
            }
            plans.Add((source, native));
        }
        // Complete all source selection before the first native append to the borrowed/base workbook.
        foreach (var plan in plans) target.MergeWorkbookFrom(plan.Source, plan.Options);
    }
}
