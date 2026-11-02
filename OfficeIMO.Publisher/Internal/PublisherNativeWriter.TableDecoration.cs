using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherNativeWriter {
    private readonly Dictionary<uint, IReadOnlyList<PublisherTableBorderSegment>> _tableBorderTargets = new();
    private PublisherEscherShapeEdits? _tableBorderRecords;
    /// <summary>Edits utility fills and shared edges within the qualified existing drawing and allocation cluster.</summary>
    private void CompareTableDecoration(PublisherTableState before, PublisherTableState after) {
        if (before.CellFills.SequenceEqual(after.CellFills) && before.BorderSegments.SequenceEqual(after.BorderSegments)) return;
        string location = "Contents/object/" + before.Id;
        PublisherTableArtwork[] artwork = _imported.Artwork.TableArtwork.Where(item => item.TableId == before.Id).ToArray();
        for (int i = 0; i < before.CellFills.Count; i++) {
            _token.ThrowIfCancellationRequested();
            if (before.CellFills[i] == after.CellFills[i]) continue;
            PublisherTableArtwork[] owners = artwork.Where(item => item.Value(0x2001) == 0 && item.Value(0x2003) == i).ToArray();
            if (owners.Length != 1) {
                Reject(location, "Cell fill saving requires one unambiguous existing native utility record."); continue;
            }
            TableFill(owners[0], before.CellFills[i], after.CellFills[i], location);
        }
        if (before.BorderSegments.SequenceEqual(after.BorderSegments)) return;
        var oldBorders = new Dictionary<(int, int, int, int), PublisherTableBorderSegment>();
        var ownersByEdge = new Dictionary<(int, int, int, int), PublisherTableArtwork>();
        foreach (PublisherTableBorderSegment border in before.BorderSegments) {
            _token.ThrowIfCancellationRequested();
            var key = PublisherEscherTableBorderPlan.Key(border);
            if (oldBorders.ContainsKey(key)) { Reject(location, "Native shared edge ownership is ambiguous."); return; }
            oldBorders.Add(key, border);
        }
        foreach (PublisherTableArtwork owner in artwork.Where(item => item.Value(0x2001) is 1 or 2)) {
            var key = ((int)owner.Value(0x2004), (int)owner.Value(0x2005), (int)owner.Value(0x2006), (int)owner.Value(0x2007));
            if (!oldBorders.ContainsKey(key) || ownersByEdge.ContainsKey(key) || !QualifiedBorder(owner)) {
                Reject(location, "Native utility edges include unsupported or ambiguous paint/ownership."); return;
            }
            ownersByEdge.Add(key, owner);
        }
        if (ownersByEdge.Count != oldBorders.Count || !DisjointBorders(before.BorderSegments) || !DisjointBorders(after.BorderSegments)) {
            Reject(location, "Native border editing requires complete non-overlapping shared grid ownership."); return;
        }
        // Publisher paints newly authored ranges as cell edges. Retain source
        // ranges verbatim; split only new ranges to reproduce native joins.
        var nativeBorders = new List<PublisherTableBorderSegment>();
        foreach (PublisherTableBorderSegment edge in after.BorderSegments) {
            bool horizontal = edge.FirstRow == edge.LastRow;
            int start = horizontal ? edge.FirstColumn : edge.FirstRow, end = horizontal ? edge.LastColumn : edge.LastRow;
            int count = oldBorders.ContainsKey(PublisherEscherTableBorderPlan.Key(edge)) ? 1 : end - start;
            if (nativeBorders.Count + (long)count > _sourceLimits.MaxItems) { Reject(location, "Native cell-edge allocation exceeds the source item limit."); return; }
            if (count == 1) nativeBorders.Add(edge);
            else for (int index = start; index < end; index++) {
                _token.ThrowIfCancellationRequested();
                nativeBorders.Add(new(horizontal ? edge.FirstRow : index, horizontal ? index : edge.FirstColumn,
                    horizontal ? edge.LastRow : index + 1, horizontal ? index + 1 : edge.LastColumn, edge.Side));
            }
        }
        var newKeys = new HashSet<(int, int, int, int)>();
        foreach (PublisherTableBorderSegment current in nativeBorders) {
            _token.ThrowIfCancellationRequested();
            var key = PublisherEscherTableBorderPlan.Key(current);
            if (!newKeys.Add(key) || current.Side.Color.A != 255 || current.Side.Width <= 0
                || current.Side.Width > (current.FirstRow == current.LastRow ? after.Height : after.Width)
                || current.Side.DashStyle != OfficeStrokeDashStyle.Solid || current.Side.LineKind != OfficeBorderLineKind.Single
                || current.Side.DoubleLineSeparation != 0 || !Emu(current.Side.Width, 0, out int width) || width <= 0) {
                Reject(location, "Shared border saving requires unique grid ranges and opaque solid single lines with exact positive signed-EMU widths."); return;
            }
            if (!oldBorders.TryGetValue(key, out PublisherTableBorderSegment prior)) continue;
            PublisherTableArtwork owner = ownersByEdge[key];
            if (prior.Side.Color != current.Side.Color) Color(owner, current.Side.Color);
            if (prior.Side.Width != current.Side.Width) Scalar(owner, 0x01CB, (uint)width);
        }
        if (!newKeys.SetEquals(oldBorders.Keys)) _tableBorderTargets.Add(before.Id, nativeBorders);

        void Color(PublisherTableArtwork owner, OfficeColor value) => Scalar(owner, 0x0181, (uint)(value.R | value.G << 8 | value.B << 16));
        void Scalar(PublisherTableArtwork owner, ushort property, uint value) {
            PublisherNativeShapeBinding? binding = owner.Shape.NativeBinding;
            PublisherNativePropertySlot[] slots = binding?.Properties.Where(item => item.Id == property).ToArray() ?? Array.Empty<PublisherNativePropertySlot>();
            if (binding == null || !binding.Matches(new PublisherBinaryData(_imported.NativeSource!.EscherBytes, "Escher/EscherStm"))
                || slots.Length != 1 || (slots[0].RawOperationId & 0xC000) != 0) {
                Reject(location, $"Table utility property 0x{property:X4} has no matching unambiguous local scalar slot."); return;
            }
            Patch(slots[0].Value, value);
        }
    }

    private bool QualifiedBorder(PublisherTableArtwork owner) {
        var shape = owner.Shape; var style = shape.Style; var binding = shape.NativeBinding;
        return binding != null && binding.Matches(new PublisherBinaryData(_imported.NativeSource!.EscherBytes, "Escher/EscherStm"))
            && shape.Type == 1 && !shape.Hidden && shape.Transform.RotationDegrees.GetValueOrDefault() == 0
            && style.ShadowEnabled != true && style.FillEnabled != false && style.FillType.GetValueOrDefault() == 0 && style.FillOpacity.GetValueOrDefault(1) == 1
            && !owner.Anchor.Keys.Any(key => key is not (0x6802 or 0x2001 or 0x2004 or 0x2005 or 0x2006 or 0x2007))
            && !binding.Properties.Any(item => item.Id == 0x01BF && (item.RawOperationId & 0xC000) != 0)
            && new ushort[] { 0x0181, 0x01CB }.All(id => binding.Properties.Count(item => item.Id == id) == 1
                && binding.Properties.Single(item => item.Id == id).RawOperationId == id);
    }

    private bool DisjointBorders(IReadOnlyList<PublisherTableBorderSegment> borders) {
        var last = new Dictionary<(bool, int), int>();
        foreach (var edge in borders.OrderBy(item => item.FirstRow == item.LastRow).ThenBy(item => item.FirstRow == item.LastRow ? item.FirstRow : item.FirstColumn)
            .ThenBy(item => item.FirstRow == item.LastRow ? item.FirstColumn : item.FirstRow)) {
            _token.ThrowIfCancellationRequested(); bool horizontal = edge.FirstRow == edge.LastRow;
            var line = (horizontal, horizontal ? edge.FirstRow : edge.FirstColumn);
            int start = horizontal ? edge.FirstColumn : edge.FirstRow, end = horizontal ? edge.LastColumn : edge.LastRow;
            if (last.TryGetValue(line, out int prior) && start < prior) return false;
            last[line] = end;
        }
        return true;
    }

    private void PlanTableBorderRecords() {
        if (_tableBorderTargets.Count == 0) return;
        var context = new PublisherParseContext(_sourceOptions, _token);
        PublisherEscherTableBorderPlan? plan = PublisherEscherTableBorderPlan.Create(_imported.NativeSource!.EscherBytes, _imported.Artwork, _tableBorderTargets, context);
        if (plan == null) { Reject("Escher/tables", "Utility edits require one existing direct utility drawing, consistent shape/allocation counts and room in its current SPID cluster."); return; }
        foreach (var edit in plan.Scalars) Patch(edit.Slot, edit.Value);
        _tableBorderRecords = plan.Shapes;
    }
}
