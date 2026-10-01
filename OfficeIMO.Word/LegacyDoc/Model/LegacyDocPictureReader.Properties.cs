using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;

namespace OfficeIMO.Word.LegacyDoc.Model;

internal static partial class LegacyDocPictureReader {
    private static bool TryReadInlinePictureCrop(byte[] data, int offset, int length, out OfficeImageSourceCrop crop) {
        crop = default;
        int end = checked(offset + length);
        uint flags = 0;
        var properties = new List<OfficeArtProperty>();
        while (offset < end) {
            if (!TryReadOfficeArtHeader(data, offset, end, out ushort type, out ushort instance, out int recordLength)) return false;
            int payload = offset + 8;
            if (type == 0xF00A) {
                if (recordLength != 8) return false;
                flags = unchecked((uint)LegacyDocFib.ReadInt32(data, payload + 4));
            } else if (type is 0xF00B or 0xF121 or 0xF122) {
                var table = OfficeArtPropertyTableReader.Read(data, payload, recordLength, instance);
                if (table.Count != instance || table.Any(property => property.IsComplex && property.AvailableComplexDataLength != property.Value)) return false;
                properties.AddRange(table);
            }
            offset = checked(payload + recordLength);
        }
        OfficeArtPictureProperties picture = OfficeArtPictureProperties.Decode(properties);
        OfficeArtShapeTransform transform = OfficeArtShapeTransform.Decode(flags, properties);
        if (picture.HasPictureEffect || transform.FlipHorizontal || transform.FlipVertical || transform.RotationDegrees.GetValueOrDefault() != 0) return false;
        if (properties.Any(property => property.PropertyId is >= 0x0100 and <= 0x0103 && property.IsComplex)) return false;
        try {
            crop = OfficeImageSourceCrop.FromStrictFractions(picture.CropFromLeft ?? 0, picture.CropFromTop ?? 0,
                picture.CropFromRight ?? 0, picture.CropFromBottom ?? 0);
            return true;
        } catch (ArgumentOutOfRangeException) { return false; }
    }
}
