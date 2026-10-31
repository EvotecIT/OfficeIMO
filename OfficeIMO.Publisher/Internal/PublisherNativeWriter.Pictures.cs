using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherNativeWriter {
    private bool _hasPictureCrop;
    private readonly List<uint> _quantizedPictureCrops = new();

    private void PictureCrop(PublisherPicture before, PublisherPicture after) {
        if (before.Crop.Equals(after.Crop)) return;
        PublisherNativeShapeBinding? binding = Binding(before.Id);
        if (binding == null) return;
        IReadOnlyList<OfficeArtProperty> crop;
        OfficeImageSourceCrop encoded;
        try { crop = OfficeArtPictureProperties.EncodeCrop(after.Crop, out encoded); }
        catch (ArgumentException) {
            Reject("Escher/object/" + before.Id, "The crop has no positive visible source area after native 16.16 rounding."); return;
        }
        var changes = _shapeProperties.TryGetValue(binding.BodyOffset, out var prior)
            ? prior.ToDictionary(pair => pair.Key, pair => pair.Value) : new Dictionary<ushort, OfficeArtProperty>();
        foreach (OfficeArtProperty property in crop) changes.Add(property.PropertyId, property);
        _shapeProperties[binding.BodyOffset] = changes;
        _hasPictureCrop = true;
        if (!encoded.Equals(after.Crop)) _quantizedPictureCrops.Add(before.Id);
    }
}
