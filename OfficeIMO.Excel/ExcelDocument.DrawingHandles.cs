using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;
using DrawingIdentity = System.ValueTuple<string, uint, string, string>;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        // Created on first use and without field initializers, which would run in the document's constructors.
        private List<WeakReference<ExcelDrawingHandle>>? _drawingHandlesValue;
        private List<WeakReference<ExcelDrawingHandle>> _drawingHandles => _drawingHandlesValue ??= new List<WeakReference<ExcelDrawingHandle>>();
        private IReadOnlyList<ExcelDrawingHandle.ReloadSnapshot>? _drawingHandlesBeforePackageClose;
        // Zero until the first prune grows it; the list is compacted from 256 handles on.
        private int _drawingHandlePruneThresholdValue;
        private int _drawingHandlePruneThreshold {
            get => _drawingHandlePruneThresholdValue == 0 ? 256 : _drawingHandlePruneThresholdValue;
            set => _drawingHandlePruneThresholdValue = value;
        }

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
        private IReadOnlyList<ExcelDrawingHandle.ReloadSnapshot> CaptureDrawingHandles(IReadOnlyDictionary<uint, ExcelPartRef> activeParts) {
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
            var indexes = new Dictionary<ExcelPartRef, Dictionary<DrawingIdentity, OfficeXmlNode>>();
            foreach (ExcelDrawingHandle.ReloadSnapshot snapshot in snapshots) {
                ExcelPartRef part = ExcelPartRef.WorksheetForMetadata(snapshot.Sheet).Drawings;
                OfficeXmlNode root = part.Root;
                if (root.IsNone) continue;
                if (!indexes.TryGetValue(part, out Dictionary<DrawingIdentity, OfficeXmlNode>? elements)) {
                    elements = new Dictionary<DrawingIdentity, OfficeXmlNode>();
                    IndexDrawingElements(root, part, elements);
                    indexes.Add(part, elements);
                }
                snapshot.Restore(part, elements);
            }
        }

        /// <summary>Indexes graphic frames and pictures in document order, as the SDK's descendant walk visits them.</summary>
        private static void IndexDrawingElements(OfficeXmlNode parent, ExcelPartRef part, Dictionary<DrawingIdentity, OfficeXmlNode> elements) {
            foreach (OfficeXmlNode element in parent.Children()) {
                if (element.Is(Xdr, "graphicFrame") || element.Is(Xdr, "pic")) {
                    if (ExcelDrawingHandle.TryGetIdentity(element, part, out DrawingIdentity identity)) {
                        // Ambiguous malformed drawing identities must not attach a retained handle to an arbitrary node.
                        if (elements.ContainsKey(identity)) elements[identity] = OfficeXmlNode.None;
                        else elements.Add(identity, element);
                    }
                }
                IndexDrawingElements(element, part, elements);
            }
        }
    }
}
