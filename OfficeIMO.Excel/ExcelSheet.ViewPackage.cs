using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private sealed class RecentWorksheetView {
            internal RecentWorksheetView(byte[] bytes, string partName, ExcelWorksheetViewInfo info) {
                Bytes = new WeakReference<byte[]>(bytes);
                PartName = partName;
                Info = info;
            }

            internal WeakReference<byte[]> Bytes { get; }
            internal string PartName { get; }
            internal ExcelWorksheetViewInfo Info { get; }
        }

        private RecentWorksheetView? _recentWorksheetView;

        internal ExcelWorksheetViewInfo? ReadDirectWorksheetView(ExcelSheet sheet) =>
            ExecuteReadWithoutMaterializing(() => {
                var candidate = _directDataSetSaveCandidate;
                if (candidate == null || !candidate.IsValid || !candidate.IsDeferred || candidate.Model.Sheets.Count != 1) return null;
                var model = candidate.Model.Sheets[0];
                string? markup = model.Metadata?.SheetViewsXml;
                return !string.IsNullOrEmpty(markup) && string.Equals(model.SheetName, sheet.Name, StringComparison.Ordinal)
                    ? ExcelSheet.ReadWorksheetViewMarkup(markup!)
                    : null;
            });

        internal ExcelWorksheetViewInfo? ReadDeferredWorksheetView(ExcelDeferredSheet sheet) =>
            ExecuteReadWithoutMaterializing(() => {
                ThrowIfDisposed();
                if (_deferredPackageBytes is not { } bytes || _deferredWorkbook == null
                    || _directDataSetSaveCandidate != null || HasPendingDirectCellValues) return null;

                RecentWorksheetView? recent = Volatile.Read(ref _recentWorksheetView);
                if (recent != null && recent.Bytes.TryGetTarget(out byte[]? previous)
                    && ReferenceEquals(bytes, previous) && sheet.PartName == recent.PartName) return recent.Info;

                using var source = new MemoryStream(bytes, writable: false);
                using OfficeOpcPackage package = OfficeOpcPackage.Open(source, leaveOpen: true);
                OfficeXmlPullReader reader = package.OpenXml(sheet.PartName, out OfficePooledBytes markup, processMarkupCompatibility: false);
                ExcelWorksheetViewInfo info;
                using (markup) info = ExcelSheet.ReadWorksheetViewStart(reader);
                Volatile.Write(ref _recentWorksheetView, new RecentWorksheetView(bytes, sheet.PartName, info));
                return info;
            });

        internal bool TryEditDeferredWorksheetView(ExcelDeferredSheet sheet, Func<OfficeXmlNode, bool> edit) {
            return WithDeferredOpenGate(() => {
                ThrowIfDisposed();
                if (_deferredPackageBytes is not { } bytes || _deferredWorkbook == null
                    || _directDataSetSaveCandidate != null || HasPendingDirectCellValues) return false;

                using var source = new MemoryStream(bytes, writable: false);
                using OfficeOpcEditablePackage package = OfficeOpcEditablePackage.Open(source, leaveOpen: true);
                bool editedMarkup;
                byte[]? replacement;
                using (OfficePooledBytes part = package.ReadPart(sheet.PartName)) {
                    editedMarkup = ExcelSheet.TryEditWorksheetViewMarkup(part, edit, out replacement);
                }
                if (editedMarkup) {
                    if (replacement == null) return true;
                    package.ReplacePart(sheet.PartName, replacement, replacement.Length);
                } else {
                    OfficeXmlTree tree = package.GetXml(sheet.PartName);
                    var worksheet = new OfficeXmlNode(tree, tree.Root);
                    if (!edit(worksheet)) return true;
                    ExcelSheet.CleanupSheetViewArtifacts(worksheet);
                }

                using var output = new MemoryStream(bytes.Length + 512);
                package.Save(output);
                _deferredPackageBytes = output.ToArray();
                _deferredPackageWrittenHere = true;
                // A new sheet with view metadata needs the worksheet writer that keeps its existing XML.
                sheet.IsNewSheet = false;
                ForgetDeferredTables();
                MarkPackageDirty();
                return true;
            });
        }
    }
}
