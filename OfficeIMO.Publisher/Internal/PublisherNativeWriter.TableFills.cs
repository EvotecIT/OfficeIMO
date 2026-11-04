using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherNativeWriter {
    private bool _hasTableFillProperties;

    /// <summary>Retains unrelated boolean uses and latent color while changing solid-fill visibility or inserting a local color scalar.</summary>
    private void TableFill(PublisherTableArtwork owner, OfficeColor? before, OfficeColor? after, string location) {
        PublisherNativeShapeBinding? binding = owner.Shape.NativeBinding;
        var style = owner.Shape.Style;
        if (binding == null || !binding.Matches(new PublisherBinaryData(_imported.NativeSource!.EscherBytes, "Escher/EscherStm"))
            || owner.Shape.Type != 1 || owner.Shape.Hidden || owner.Shape.Transform.RotationDegrees.GetValueOrDefault() != 0
            || owner.Anchor.Keys.Any(key => key is not (0x6802 or 0x2001 or 0x2003))
            || binding.Properties.Any(item => item.Id == 0x01BF && (item.RawOperationId & 0xC000) != 0)
            || style.ShadowEnabled == true || style.FillType.GetValueOrDefault() != 0 || style.FillOpacity.GetValueOrDefault(1) != 1
            || after is { A: not 255 }) {
            Reject(location, "The native cell utility has no qualified solid-fill binding, paint or ownership profile."); return;
        }
        if (before.HasValue != after.HasValue) {
            const uint useFilled = 0x00100000, filled = 0x00000010;
            PublisherNativePropertySlot[] flags = binding.Properties.Where(item => item.Id == 0x01BF && (item.Value.Original & useFilled) != 0).ToArray();
            if (flags.Length != 1) {
                Reject(location, "Cell visibility requires exactly one existing local scalar owning the native filled use-bit."); return;
            }
            uint value = flags[0].Value.Original;
            Patch(flags[0].Value, after.HasValue ? value | filled : value & ~filled);
            _hasTableFillProperties = true;
        }
        if (!after.HasValue) return; // Keep the latent source color when hiding a cell fill.
        PublisherNativePropertySlot[] colors = binding.Properties.Where(item => item.Id == 0x0181).ToArray();
        OfficeColor color = after.Value;
        uint encoded = (uint)(color.R | color.G << 8 | color.B << 16);
        if (colors.Length == 1 && (colors[0].RawOperationId & 0xC000) == 0) {
            Patch(colors[0].Value, encoded); return;
        }
        if (colors.Length != 0) { Reject(location, "Cell color has ambiguous, complex or BLIP local property ownership."); return; }
        var changes = _shapeProperties.TryGetValue(binding.BodyOffset, out var prior)
            ? prior.ToDictionary(item => item.Key, item => item.Value) : new Dictionary<ushort, OfficeArtProperty>();
        if (changes.ContainsKey(0x0181)) { Reject(location, "Semantic owners disagree about the native cell color insertion."); return; }
        changes.Add(0x0181, new OfficeArtProperty(0, 0x0181, encoded));
        _shapeProperties[binding.BodyOffset] = changes;
        _hasTableFillProperties = true;
    }
}
