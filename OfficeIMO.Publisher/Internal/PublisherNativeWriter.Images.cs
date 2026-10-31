using System.Text;
using OfficeIMO.Drawing.Binary;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherNativeWriter {
    private PublisherNativeStreamEdit? _imageStore;
    private byte[]? _imageDelay;
    private readonly Dictionary<uint, string> _pictureNames = new();
    private readonly List<string> _imageRenumbering = new();

    private void PrepareImages() {
        foreach (var entry in _baseline.Images) {
            _token.ThrowIfCancellationRequested();
            if (_current.Images.TryGetValue(entry.Key, out PublisherImage? same) && ReferenceEquals(same, entry.Value)) continue;
            if (!_current.Images.TryGetValue(entry.Key, out PublisherImage? current)
                || current.ContentType != entry.Value.ContentType || !current.Bytes.SequenceEqual(entry.Value.Bytes)) {
                Reject("Escher/images", "Original source assets must remain present and byte-identical."); return;
            }
        }
        var additions = _current.Images.Values.Where(image => !_baseline.Images.ContainsKey(image.Id)).OrderBy(image => image.Id).ToArray();
        var changed = _baseline.Pictures.Values.Where(picture => picture.ImageId != _current.Pictures[picture.Id].ImageId).ToArray();
        if (additions.Length == 0 && changed.Length == 0) return;
        var context = new PublisherParseContext(new PublisherReadOptions { Limits = _sourceLimits }, _token);
        if (!TryPrepareImages(additions, changed, context))
            Reject("Escher/images", "Picture replacement has an unsupported image-store, source binding or asset profile.");
    }

    private bool TryPrepareImages(PublisherImage[] additions, PublisherPicture[] changed, PublisherParseContext context) {
        if (additions.Length == 0 || changed.Length == 0 || _imported.Artwork.ImageStores.Count != 1) return false;
        PublisherEscherRecord store = _imported.Artwork.ImageStores[0];
        byte[] source = _imported.NativeSource!.EscherBytes;
        var data = new PublisherBinaryData(source, "Escher/EscherStm");
        if ((store.Initial & 15) != 15 || data.U16(store.Offset - 6) != 0xF001
            || data.U32(store.Offset - 4) != store.Length || data.U16(store.Offset - 8) != store.Initial) return false;
        var entries = new PublisherEscherReader(data, null, context).Records(store.Offset, store.End);
        if (entries.Count != store.Initial >> 4 || entries.Count + additions.Length > 4095
            || entries.Any(entry => entry.Kind != 0xF007 || entry.Length < 36)) return false;
        // Appending must not resolve a previously missing reference in any
        // native artwork role, including fill/print bindings outside the model.
        if (_imported.Artwork.ImageReferences.Any(reference => reference > entries.Count && reference <= entries.Count + additions.Length)) return false;
        var decrements = new Dictionary<int, uint>();
        var counts = additions.ToDictionary(image => image.Id, _ => 0U);
        var indices = additions.Select((image, index) => (image.Id, Index: entries.Count + index + 1)).ToDictionary(pair => pair.Id, pair => pair.Index);
        foreach (PublisherPicture before in changed) {
            _token.ThrowIfCancellationRequested();
            PublisherPicture after = _current.Pictures[before.Id];
            if (!before.CanEdit || !before.ImageId.HasValue || before.ImageId.Value < 1 || before.ImageId.Value > entries.Count
                || !after.ImageId.HasValue || !counts.ContainsKey(after.ImageId.Value)) return false;
            counts[after.ImageId.Value]++;
            decrements[before.ImageId.Value] = decrements.TryGetValue(before.ImageId.Value, out uint count) ? count + 1 : 1;
            PublisherNativeShapeBinding? binding = Binding(before.Id); if (binding == null) return false;
            PublisherEscherShape shape = _artwork[before.Id];
            var references = shape.Properties.Where(property => property.PropertyId == 0x104).ToArray();
            var names = shape.Properties.Where(property => property.PropertyId == 0x105).ToArray();
            if (references.Length != 1 || references[0].RawOperationId != 0x4104 || references[0].Value != before.ImageId.Value
                || names.Length != 1 || names[0].RawOperationId != 0xC105 || !names[0].HasCompleteComplexData) return false;
            PublisherImage image = _current.Images[after.ImageId.Value];
            if (image.ContentType is not ("image/png" or "image/jpeg")) return false;
            string name = "Image" + indices[image.Id].ToString(System.Globalization.CultureInfo.InvariantCulture)
                + (image.ContentType == "image/png" ? ".png" : ".jpg");
            byte[] filename = Encoding.Unicode.GetBytes(Path.GetFileNameWithoutExtension(name) + "\0");
            _shapeProperties.Add(binding.BodyOffset, new Dictionary<ushort, OfficeArtProperty> {
                [0x104] = new(0, 0x4104, (uint)indices[image.Id]),
                [0x105] = new(0, 0xC105, (uint)filename.Length, filename.Length, complexData: filename)
            });
            _pictureNames.Add(before.Id, name);
        }
        if (counts.Values.Any(count => count == 0)) return false;
        long delayLength = _imported.NativeSource.ImageDelayBytes?.Length ?? 0;
        bool delayed = _imported.NativeSource.ImageDelayBytes != null;
        long addedBytes = additions.Sum(image => (long)image.ByteCount + 25);
        if (addedBytes > _maximumBytes - delayLength) throw new InvalidDataException("Publisher native output byte limit exceeded.");
        using var payload = new MemoryStream();
        using var delay = new MemoryStream(delayed ? checked((int)(delayLength + addedBytes)) : 0);
        if (delayed) delay.Write(_imported.NativeSource.ImageDelayBytes!, 0, (int)delayLength);
        for (int index = 0; index < entries.Count; index++) {
            _token.ThrowIfCancellationRequested();
            PublisherEscherRecord entry = entries[index];
            byte[] bytes = new byte[entry.Length + 8]; Buffer.BlockCopy(source, entry.Offset - 8, bytes, 0, bytes.Length);
            if (decrements.TryGetValue(index + 1, out uint decrement)) {
                uint old = data.U32(entry.Offset + 24); if (decrement > old) return false;
                BitConverter.GetBytes(old - decrement).CopyTo(bytes, 32);
            }
            payload.Write(bytes, 0, bytes.Length);
        }
        foreach (PublisherImage image in additions) {
            _token.ThrowIfCancellationRequested();
            byte[] entry;
            if (delayed) {
                entry = OfficeArtBlipStoreEntryWriter.CreateDelayed(image.Bytes, image.ContentType, checked((uint)delay.Length), counts[image.Id]);
                byte[] blip = OfficeArtBlipStoreEntryWriter.CreateBlipRecord(image.Bytes, image.ContentType); delay.Write(blip, 0, blip.Length);
            } else entry = OfficeArtBlipStoreEntryWriter.CreateEmbedded(image.Bytes, image.ContentType, counts[image.Id]);
            if (entry.Length > _maximumBytes - payload.Length) throw new InvalidDataException("Publisher native output byte limit exceeded.");
            payload.Write(entry, 0, entry.Length);
            if (image.Id != indices[image.Id]) _imageRenumbering.Add(image.Id + " -> " + indices[image.Id]);
        }
        byte[] replacement = new byte[checked((int)payload.Length + 8)];
        BitConverter.GetBytes((ushort)(((entries.Count + additions.Length) << 4) | 15)).CopyTo(replacement, 0);
        BitConverter.GetBytes((ushort)0xF001).CopyTo(replacement, 2);
        BitConverter.GetBytes((uint)payload.Length).CopyTo(replacement, 4); payload.Position = 0;
        _ = payload.Read(replacement, 8, (int)payload.Length);
        _imageStore = new(store.Offset - 8, store.Length + 8, replacement);
        if (delayed) _imageDelay = delay.ToArray();
        return true;
    }
}
