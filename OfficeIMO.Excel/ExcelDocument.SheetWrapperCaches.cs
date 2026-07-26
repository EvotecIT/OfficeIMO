using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private readonly object _sheetWrapperCacheLock = new object();
        private readonly List<WeakReference<ExcelSheet>> _sheetWrappers = new();

        internal void RegisterSheetWrapper(ExcelSheet sheet) {
            lock (_sheetWrapperCacheLock) {
                _sheetWrappers.Add(new WeakReference<ExcelSheet>(sheet));
            }
        }

        internal void ResetStructuralMutationCaches(WorksheetPart worksheetPart) {
            var wrappers = new List<ExcelSheet>();
            lock (_sheetWrapperCacheLock) {
                for (int index = _sheetWrappers.Count - 1; index >= 0; index--) {
                    if (!_sheetWrappers[index].TryGetTarget(out ExcelSheet? sheet)) {
                        _sheetWrappers.RemoveAt(index);
                        continue;
                    }

                    if (ReferenceEquals(sheet.DeferredMetadataWorksheetPart, worksheetPart)) {
                        wrappers.Add(sheet);
                    }
                }
            }

            foreach (ExcelSheet sheet in wrappers) {
                sheet.ResetStructuralMutationCachesLocal();
            }
        }
    }
}
