using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;
using DrawingIdentity = System.ValueTuple<string, uint, string, string>;

namespace OfficeIMO.Excel {
    /// <summary>Tracks a live drawing wrapper through replacement of its worksheet package.</summary>
    internal sealed class ExcelDrawingHandle {
        private OfficeXmlNode _element;
        private ExcelPartRef _part;
        private readonly Action? _afterReload;

        internal ExcelDrawingHandle(OfficeXmlNode element, ExcelPartRef part, ExcelSheet sheet, Action? afterReload = null) {
            if (element.IsNone) throw new ArgumentNullException(nameof(element));
            if (part.IsNone) throw new ArgumentNullException(nameof(part));
            _element = element;
            _part = part;
            Sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));
            _afterReload = afterReload;
            sheet.Document.RegisterDrawingHandle(this);
        }

        // Retaining a drawing also retains its parent, even when worksheet caching is disabled.
        internal ExcelSheet Sheet { get; }
        internal OfficeXmlNode Element { get { EnsureActive(); return _element; } }
        internal ExcelPartRef Part { get { EnsureActive(); return _part; } }

        private bool IsActive => ExcelPartRef.IsSheetAttached(Sheet)
            && _part == ExcelPartRef.WorksheetForMetadata(Sheet).Drawings
            && _element.AncestorOrSelf(Xdr, "wsDr") == _part.Root;

        private void EnsureActive() {
            if (!IsActive) throw new InvalidOperationException("This drawing is no longer attached to the active worksheet.");
        }

        /// <summary>Captures only attached originals, so removed or newly replaced drawings cannot acquire an old identity.</summary>
        internal ReloadSnapshot? Capture(IReadOnlyDictionary<uint, ExcelPartRef> activeParts) {
            uint? sheetId = ExcelPartRef.SheetId(Sheet);
            if (sheetId == null || !activeParts.TryGetValue(sheetId.Value, out ExcelPartRef part)
                || part != ExcelPartRef.WorksheetForMetadata(Sheet) || !IsActive
                || !TryGetIdentity(_element, _part, out DrawingIdentity identity)) return null;
            return new ReloadSnapshot(this, identity);
        }

        internal static bool TryGetIdentity(OfficeXmlNode element, ExcelPartRef part, out DrawingIdentity identity) {
            identity = default;
            uint? id = element.FirstDescendant(Xdr, "cNvPr").UInt32Attribute("id");
            string? relationshipId;
            if (element.Is(Xdr, "graphicFrame")) {
                relationshipId = element.Child(A, "graphic").Child(A, "graphicData").Child(C, "chart").RelationshipAttribute("id");
            } else if (element.Is(Xdr, "pic")) {
                OfficeXmlNode blip = element.Child(Xdr, "blipFill").Child(A, "blip");
                relationshipId = blip.RelationshipAttribute("embed") ?? blip.RelationshipAttribute("link");
            } else {
                relationshipId = null;
            }
            if (id == null || string.IsNullOrEmpty(relationshipId)) return false;
            ExcelPartRef related = part.GetPartById(relationshipId);
            string? target = related.Exists
                ? related.Uri!.ToString()
                : part.GetExternalRelationshipUri(relationshipId!)?.ToString();
            if (target == null) return false;
            identity = (element.LocalName, id.Value, relationshipId!, target);
            return true;
        }

        internal sealed class ReloadSnapshot {
            private readonly ExcelDrawingHandle _handle;
            private readonly DrawingIdentity _identity;
            internal ReloadSnapshot(ExcelDrawingHandle handle, DrawingIdentity identity) { _handle = handle; _identity = identity; }
            internal ExcelSheet Sheet => _handle.Sheet;

            internal void Restore(ExcelPartRef part, IReadOnlyDictionary<DrawingIdentity, OfficeXmlNode> elements) {
                if (!elements.TryGetValue(_identity, out OfficeXmlNode element) || element.IsNone) return;
                _handle._element = element;
                _handle._part = part;
                _handle._afterReload?.Invoke();
            }
        }
    }
}
