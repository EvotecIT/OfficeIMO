namespace OfficeIMO.Excel {
    /// <summary>
    /// The models formulas calculate over. A workbook that has not opened its SDK package calculates over its own
    /// worksheet XML, with the recorded changes applied; its results are recorded for the owned save. A workbook with
    /// cell metadata (dynamic arrays and rich values), or one whose markup the owned reader leaves to the model, opens
    /// the SDK package and calculates over it.
    /// </summary>
    public partial class ExcelDocument {
        private ExcelXmlFormulaWorkbookModel? _xmlFormulaWorkbook;

        /// <summary>The workbook as formulas see it.</summary>
        internal ExcelFormulaWorkbookModel FormulaWorkbookModel =>
            (ExcelFormulaWorkbookModel?)TryGetXmlFormulaWorkbook() ?? SdkFormulaWorkbookModel;

        /// <summary>Whether the workbook has a cell metadata part, which dynamic arrays and rich values need.</summary>
        internal bool HasCellMetadataPart => FormulaWorkbookModel.HasCellMetadata;

        /// <summary>The workbook read from its own XML, or null when it calculates over the SDK model.</summary>
        private ExcelXmlFormulaWorkbookModel? TryGetXmlFormulaWorkbook() {
            ExcelDeferredWorkbook? workbook = _deferredWorkbook;
            byte[]? bytes = _deferredPackageBytes;
            if (workbook == null || bytes == null) return null;
            ExcelXmlFormulaWorkbookModel? model = _xmlFormulaWorkbook;
            if (model != null && ReferenceEquals(model.PackageBytes, bytes) && ReferenceEquals(model.Workbook, workbook)) return model;
            if (WorkbookHasCellMetadata(workbook, bytes)) return null;
            model = new ExcelXmlFormulaWorkbookModel(this, workbook, bytes);
            if (!model.TryReadParts()) return null;
            return _xmlFormulaWorkbook = model;
        }

        /// <summary>The sheet as formulas see it.</summary>
        internal ExcelFormulaSheetModel GetFormulaSheetModel(ExcelSheet sheet) {
            if (sheet.DeferredEntry is { } entry && TryGetXmlFormulaWorkbook() is { } workbook) {
                lock (entry) {
                    if (entry.FormulaModel is ExcelXmlFormulaSheetModel model && ReferenceEquals(model.PackageBytes, workbook.PackageBytes)
                        && model.TrySync()) {
                        return model;
                    }
                    model = ExcelXmlFormulaSheetModel.TryRead(this, sheet, entry, workbook)!;
                    if (model != null && model.TrySync()) {
                        entry.FormulaModel = model;
                        return model;
                    }
                    entry.FormulaModel = null;
                }
                OpenDeferredSdk("Formulas on worksheet '" + entry.Name + "' need the model");
            }
            return sheet.SdkFormulaModel;
        }

        /// <summary>A handle to the worksheet with the name, from the model formulas read the workbook through.</summary>
        internal ExcelSheet? FindSheetForFormulas(string name) {
            ExcelFormulaWorkbookModel workbook = FormulaWorkbookModel;
            foreach (ExcelFormulaSheetEntry entry in workbook.SheetEntries()) {
                if (entry.IsWorksheet && string.Equals(entry.Name, name, StringComparison.OrdinalIgnoreCase)) return workbook.CreateSheet(entry);
            }
            return null;
        }

        /// <summary>
        /// Records a calculated result for the owned save; false when the workbook has opened its SDK package and the
        /// result goes to the model instead. Consecutive results for one cell keep the latest.
        /// </summary>
        internal bool RecordDeferredFormulaResult(ExcelSheet sheet, ExcelDeferredSheet entry, ExcelXmlFormulaCellRecord cell) {
            ExcelDeferredCellEdit edit = ExcelDeferredCellEdit.FormulaResult(cell.Row, cell.Column, cell.Value, cell.Type,
                cell.HasFormula ? cell.Calculate : null);
            lock (_deferredEditsGate) {
                if (_deferredWorkbook == null || !ReferenceEquals(sheet.DeferredEntry, entry)) return false;
                ExcelDeferredCellStore? store = entry.Cells;
                if (store == null) {
                    entry.Cells = store = new ExcelDeferredCellStore();
                    _deferredEditedSheets.Add(sheet);
                    store.Add(edit);
                } else if (store.Count > 0 && store[store.Count - 1] is { Kind: ExcelDeferredCellKind.FormulaResult } last
                    && last.Row == edit.Row && last.Column == edit.Column) {
                    store[store.Count - 1] = edit;
                } else {
                    store.Add(edit);
                }
            }
            if (!_packageDirty) MarkPackageDirty();
            return true;
        }
    }
}
