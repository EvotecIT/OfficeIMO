using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherNativeWriter {
    /// <summary>Patches only qualified existing utility scalars; adding records, visibility changes and grid edits remain refused.</summary>
    private void CompareTableDecoration(PublisherTableState before, PublisherTableState after) {
        if (before.CellFills.SequenceEqual(after.CellFills) && before.BorderSegments.SequenceEqual(after.BorderSegments)) return;
        string location = "Contents/object/" + before.Id;
        PublisherTableArtwork[] artwork = _imported.Artwork.TableArtwork.Where(item => item.TableId == before.Id).ToArray();
        for (int i = 0; i < before.CellFills.Count; i++) {
            _token.ThrowIfCancellationRequested();
            if (before.CellFills[i] == after.CellFills[i]) continue;
            PublisherTableArtwork[] owners = artwork.Where(item => item.Value(0x2001) == 0 && item.Value(0x2003) == i).ToArray();
            if (!before.CellFills[i].HasValue || !after.CellFills[i].HasValue || after.CellFills[i]!.Value.A != 255 || owners.Length != 1) {
                Reject(location, "Cell fill saving requires one existing visible solid utility fill; visibility changes and record insertion are not qualified."); continue;
            }
            Color(owners[0], after.CellFills[i]!.Value);
        }
        if (before.BorderSegments.Count != after.BorderSegments.Count) {
            Reject(location, "Adding or removing native shared border records is not qualified."); return;
        }
        for (int i = 0; i < before.BorderSegments.Count; i++) {
            _token.ThrowIfCancellationRequested();
            PublisherTableBorderSegment prior = before.BorderSegments[i], current = after.BorderSegments[i];
            if (prior.Equals(current)) continue;
            if (prior.FirstRow != current.FirstRow || prior.FirstColumn != current.FirstColumn
                || prior.LastRow != current.LastRow || prior.LastColumn != current.LastColumn
                || current.Side.Color.A != 255 || current.Side.Width <= 0
                || current.Side.DashStyle != OfficeStrokeDashStyle.Solid || current.Side.LineKind != OfficeBorderLineKind.Single
                || current.Side.DoubleLineSeparation != 0 || !Emu(current.Side.Width, 0, out int width) || width <= 0) {
                Reject(location, "Shared border saving requires unchanged endpoints and an opaque single solid line with exact positive signed-EMU width."); continue;
            }
            uint kind = prior.FirstRow == prior.LastRow ? 1U : 2U;
            PublisherTableArtwork[] owners = artwork.Where(item => item.Value(0x2001) == kind
                && item.Value(0x2004) == prior.FirstRow && item.Value(0x2005) == prior.FirstColumn
                && item.Value(0x2006) == prior.LastRow && item.Value(0x2007) == prior.LastColumn).ToArray();
            if (owners.Length != 1) { Reject(location, "The shared border has no unambiguous native utility record."); continue; }
            if (prior.Side.Color != current.Side.Color) Color(owners[0], current.Side.Color);
            if (prior.Side.Width != current.Side.Width) Scalar(owners[0], 0x01CB, (uint)width);
        }

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
}
