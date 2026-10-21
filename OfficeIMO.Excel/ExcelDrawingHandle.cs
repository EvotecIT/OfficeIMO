using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DrawingIdentity = System.ValueTuple<string, uint, string, string>;

namespace OfficeIMO.Excel {
    /// <summary>Tracks a live drawing wrapper through replacement of its worksheet package.</summary>
    internal sealed class ExcelDrawingHandle {
        private OpenXmlElement _element;
        private DrawingsPart _part;
        private readonly Action? _afterReload;

        internal ExcelDrawingHandle(OpenXmlElement element, DrawingsPart part, ExcelSheet sheet, Action? afterReload = null) {
            _element = element ?? throw new ArgumentNullException(nameof(element));
            _part = part ?? throw new ArgumentNullException(nameof(part));
            Sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));
            _afterReload = afterReload;
            sheet.Document.RegisterDrawingHandle(this);
        }

        // Retaining a drawing also retains its parent, even when worksheet caching is disabled.
        internal ExcelSheet Sheet { get; }
        internal OpenXmlElement Element { get { EnsureActive(); return _element; } }
        internal DrawingsPart Part { get { EnsureActive(); return _part; } }

        private bool IsActive => Sheet.SheetElement.Parent != null
            && ReferenceEquals(_part, Sheet.DeferredMetadataWorksheetPart.DrawingsPart)
            && ReferenceEquals(_element.Ancestors<Xdr.WorksheetDrawing>().FirstOrDefault(), _part.WorksheetDrawing);

        private void EnsureActive() {
            if (!IsActive) throw new InvalidOperationException("This drawing is no longer attached to the active worksheet.");
        }

        /// <summary>Captures only attached originals, so removed or newly replaced drawings cannot acquire an old identity.</summary>
        internal ReloadSnapshot? Capture(IReadOnlyDictionary<uint, WorksheetPart> activeParts) {
            uint? sheetId = Sheet.SheetElement.SheetId?.Value;
            if (sheetId == null || !activeParts.TryGetValue(sheetId.Value, out WorksheetPart? part)
                || !ReferenceEquals(part, Sheet.DeferredMetadataWorksheetPart) || !IsActive
                || !TryGetIdentity(_element, _part, out DrawingIdentity identity)) return null;
            return new ReloadSnapshot(this, identity);
        }

        internal static bool TryGetIdentity(OpenXmlElement element, DrawingsPart part, out DrawingIdentity identity) {
            identity = default;
            uint? id = element.Descendants<Xdr.NonVisualDrawingProperties>().FirstOrDefault()?.Id?.Value;
            string? relationshipId = element is Xdr.GraphicFrame frame
                ? frame.Graphic?.GraphicData?.GetFirstChild<C.ChartReference>()?.Id?.Value
                : (element as Xdr.Picture)?.BlipFill?.Blip?.Embed?.Value
                    ?? (element as Xdr.Picture)?.BlipFill?.Blip?.Link?.Value;
            if (id == null || string.IsNullOrEmpty(relationshipId)) return false;
            string? target = part.TryGetPartById(relationshipId!, out OpenXmlPart? related)
                ? related!.Uri.ToString()
                : part.ExternalRelationships.FirstOrDefault(item => item.Id == relationshipId)?.Uri.ToString();
            if (target == null) return false;
            identity = (element.LocalName, id.Value, relationshipId!, target);
            return true;
        }

        internal sealed class ReloadSnapshot {
            private readonly ExcelDrawingHandle _handle;
            private readonly DrawingIdentity _identity;
            internal ReloadSnapshot(ExcelDrawingHandle handle, DrawingIdentity identity) { _handle = handle; _identity = identity; }
            internal ExcelSheet Sheet => _handle.Sheet;

            internal void Restore(DrawingsPart part, IReadOnlyDictionary<DrawingIdentity, OpenXmlElement?> elements) {
                if (!elements.TryGetValue(_identity, out OpenXmlElement? element) || element == null) return;
                _handle._element = element;
                _handle._part = part;
                _handle._afterReload?.Invoke();
            }
        }
    }
}
