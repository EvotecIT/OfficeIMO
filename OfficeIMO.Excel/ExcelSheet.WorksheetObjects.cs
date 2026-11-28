using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const int WorksheetMarkup = OfficeXmlNamespaces.SpreadsheetML;

        internal T ReadWorksheetObject<T>(Func<OfficeXmlNode, T> read) =>
            _excelDocument.ExecuteReadWithoutMaterializing(() => {
                if (DeferredEntry is { } entry && _excelDocument.TryReadWorksheetObject(entry, read, out T result)) return result;
                return read(ViewWorksheetRoot);
            });

        private void EditWorksheetObject(Func<OfficeXmlNode, bool> edit, bool save = false,
            Action<OfficeOpcEditablePackage, ExcelDeferredSheet, OfficeXmlNode>? afterOwnedEdit = null) {
            Locking.ExecuteWrite(_excelDocument.EnsureLock(), () => {
                if (_excelDocument.HasDeferredDirectDataSetImport) _excelDocument.MaterializeDeferredDataSetImport();
                if (DeferredEntry is { } entry && _excelDocument.TryEditWorksheetObject(entry, edit, afterOwnedEdit)) return;
                if (!edit(ViewWorksheetRoot)) return;
                if (save) ExcelPartRef.WorksheetForMetadata(this).SaveRoot();
                MarkRequiresSavePreparation();
            });
        }

        private static OfficeXmlNode AddWorksheetObject(
            OfficeXmlNode worksheet, string name, params string[] subsequentElements) {
            OfficeXmlNode existing = worksheet.Child(WorksheetMarkup, name);
            if (!existing.IsNone) return existing;

            OfficeXmlNode before = default;
            foreach (OfficeXmlNode child in worksheet.Children()) {
                if (child.Is(WorksheetMarkup, "extLst") ||
                    child.NamespaceId == WorksheetMarkup && Array.IndexOf(subsequentElements, child.LocalName) >= 0) {
                    before = child;
                    break;
                }
            }
            return worksheet.InsertBefore(worksheet.CreateElement(WorksheetMarkup, name), before);
        }
    }

    public partial class ExcelDocument {
        internal bool TryReadWorksheetObject<T>(ExcelDeferredSheet sheet, Func<OfficeXmlNode, T> read, out T result) {
            ThrowIfDisposed();
            if (_deferredPackageBytes is not { } bytes || _deferredWorkbook == null ||
                _directDataSetSaveCandidate != null || HasPendingDirectCellValues) {
                result = default!;
                return false;
            }

            using var source = new MemoryStream(bytes, writable: false);
            using OfficeOpcEditablePackage package = OfficeOpcEditablePackage.Open(source, leaveOpen: true);
            OfficeXmlTree worksheet = package.GetXml(sheet.PartName);
            result = read(new OfficeXmlNode(worksheet, worksheet.Root));
            return true;
        }

        internal bool TryReadWorksheetPackage<T>(ExcelDeferredSheet sheet, Func<OfficeOpcPackage, ExcelDeferredSheet, T> read, out T result) =>
            TryReadWorkbookObjectPackage(sheet, read, out result);

        private bool TryReadWorkbookObjectPackage<TState, T>(TState state, Func<OfficeOpcPackage, TState, T> read, out T result) {
            ThrowIfDisposed();
            if (_deferredPackageBytes is not { } bytes || _deferredWorkbook == null) {
                result = default!;
                return false;
            }
            using var source = new MemoryStream(bytes, writable: false);
            using OfficeOpcPackage package = OfficeOpcPackage.Open(source, leaveOpen: true);
            result = read(package, state);
            return true;
        }

        internal bool TryEditWorksheetObject(ExcelDeferredSheet sheet, Func<OfficeXmlNode, bool> edit,
            Action<OfficeOpcEditablePackage, ExcelDeferredSheet, OfficeXmlNode>? afterOwnedEdit = null) =>
            TryEditWorksheetPackage(sheet, (package, entry) => {
                OfficeXmlTree worksheet = package.GetXml(entry.PartName);
                var root = new OfficeXmlNode(worksheet, worksheet.Root);
                if (!edit(root)) return false;
                afterOwnedEdit?.Invoke(package, entry, root);
                return true;
            });

        internal bool TryEditWorksheetPackage(ExcelDeferredSheet sheet, Func<OfficeOpcEditablePackage, ExcelDeferredSheet, bool> edit) =>
            WithDeferredOpenGate(() => {
                ThrowIfDisposed();
                if (_deferredPackageBytes is not { } bytes || _deferredWorkbook == null ||
                    _directDataSetSaveCandidate != null || HasPendingDirectCellValues) return false;

                using var source = new MemoryStream(bytes, writable: false);
                using OfficeOpcEditablePackage package = OfficeOpcEditablePackage.Open(source, leaveOpen: true);
                if (!edit(package, sheet)) return true;
                using var output = new MemoryStream(bytes.Length + 512);
                package.Save(output);
                _deferredPackageBytes = output.ToArray();
                _deferredPackageWrittenHere = true;
                sheet.IsNewSheet = false;
                ForgetDeferredTables();
                MarkPackageDirty();
                return true;
            });
    }
}
