using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DrawingIdentity = System.ValueTuple<string, uint, string, string>;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private readonly List<WeakReference<ExcelDrawingHandle>> _drawingHandles = new();
        private IReadOnlyList<ExcelDrawingHandle.ReloadSnapshot>? _drawingHandlesBeforePackageClose;

        internal void RegisterDrawingHandle(ExcelDrawingHandle handle) {
            lock (_sheetWrapperCacheLock) { _drawingHandles.Add(new WeakReference<ExcelDrawingHandle>(handle)); }
        }

        /// <summary>Snapshots live drawing identity before package replacement or a transactional mutation.</summary>
        internal IReadOnlyList<ExcelDrawingHandle.ReloadSnapshot> CaptureDrawingHandlesForReload(
            IReadOnlyDictionary<uint, WorksheetPart>? activeParts = null) {
            activeParts ??= CaptureActiveSheetPartsForReload();
            var snapshots = new List<ExcelDrawingHandle.ReloadSnapshot>();
            lock (_sheetWrapperCacheLock) {
                for (int index = _drawingHandles.Count - 1; index >= 0; index--) {
                    if (!_drawingHandles[index].TryGetTarget(out ExcelDrawingHandle? handle)) {
                        _drawingHandles.RemoveAt(index);
                        continue;
                    }
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
