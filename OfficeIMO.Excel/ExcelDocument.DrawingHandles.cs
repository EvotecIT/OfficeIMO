using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DrawingIdentity = System.ValueTuple<string, uint, string, string>;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private readonly List<WeakReference<ExcelDrawingHandle>> _drawingHandles = new();
        private IReadOnlyList<ExcelDrawingHandle.ReloadSnapshot>? _drawingHandlesBeforePackageClose;
        private int _drawingHandlePruneThreshold = 256;

        internal void RegisterDrawingHandle(ExcelDrawingHandle handle) {
            lock (_sheetWrapperCacheLock) {
                if (_drawingHandles.Count >= _drawingHandlePruneThreshold) PruneCollectedDrawingHandles();
                _drawingHandles.Add(new WeakReference<ExcelDrawingHandle>(handle));
            }
        }

        /// <summary>Compacts collected references in one pass while the sheet-wrapper lock is held.</summary>
        private void PruneCollectedDrawingHandles() {
            _drawingHandles.RemoveAll(reference => !reference.TryGetTarget(out _));
            // Grow with live aliases so registering them does not repeatedly scan the whole list.
            _drawingHandlePruneThreshold = (int)Math.Min(int.MaxValue, Math.Max(256L, (long)_drawingHandles.Count * 2));
        }

        /// <summary>Snapshots live drawing identity before package replacement or a transactional mutation.</summary>
        internal IReadOnlyList<ExcelDrawingHandle.ReloadSnapshot> CaptureDrawingHandlesForReload(
            IReadOnlyDictionary<uint, WorksheetPart>? activeParts = null) {
            activeParts ??= CaptureActiveSheetPartsForReload();
            var snapshots = new List<ExcelDrawingHandle.ReloadSnapshot>();
            lock (_sheetWrapperCacheLock) {
                PruneCollectedDrawingHandles();
                foreach (WeakReference<ExcelDrawingHandle> reference in _drawingHandles) {
                    if (!reference.TryGetTarget(out ExcelDrawingHandle? handle)) continue;
                    ExcelDrawingHandle.ReloadSnapshot? snapshot = handle.Capture(activeParts);
                    if (snapshot != null) snapshots.Add(snapshot);
                }
            }
            return snapshots;
        }

        /// <summary>Reconnects captured originals, indexing each restored drawing part once for all aliases.</summary>
        internal static void RebindDrawingHandlesAfterPackageReload(IReadOnlyList<ExcelDrawingHandle.ReloadSnapshot> snapshots) {
            var indexes = new Dictionary<DrawingsPart, Dictionary<DrawingIdentity, OpenXmlElement?>>();
            foreach (ExcelDrawingHandle.ReloadSnapshot snapshot in snapshots) {
                DrawingsPart? part = snapshot.Sheet.DeferredMetadataWorksheetPart.DrawingsPart;
                if (part?.WorksheetDrawing == null) continue;
                if (!indexes.TryGetValue(part, out Dictionary<DrawingIdentity, OpenXmlElement?>? elements)) {
                    elements = new Dictionary<DrawingIdentity, OpenXmlElement?>();
                    foreach (OpenXmlElement element in part.WorksheetDrawing.Descendants()
                        .Where(item => item is Xdr.GraphicFrame || item is Xdr.Picture)) {
                        if (!ExcelDrawingHandle.TryGetIdentity(element, part, out DrawingIdentity identity)) continue;
                        // Ambiguous malformed drawing identities must not attach a retained handle to an arbitrary node.
                        if (elements.ContainsKey(identity)) elements[identity] = null;
                        else elements.Add(identity, element);
                    }
                    indexes.Add(part, elements);
                }
                snapshot.Restore(part, elements);
            }
        }
    }
}
