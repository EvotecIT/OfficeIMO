using System;
using System.Linq;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeHeifImageComposer {
    private void ApplyColor(OfficeRasterImage image, OfficeHeifItemInfo item) {
        var properties = item.PropertyAssociations.Where(p => p.PropertyType == "colr").ToArray();
        if (properties.Length == 0 && item.ItemType is "grid" or "iden") {
            uint[] children = Targets(item.ItemId, "dimg");
            if (children.Length > 0) { ApplyColor(image, FindItem(children[0])); return; }
        }
        OfficeHeifItemPropertyAssociationInfo? icc = null, nclx = null;
        foreach (var property in properties) {
            Require(property.PayloadLength >= 4, "Truncated HEIF color property.");
            int p = property.PayloadOffset;
            string kind = System.Text.Encoding.ASCII.GetString(_bytes, p, 4);
            if (kind is "prof" or "rICC") {
                Require(icc == null, "Ambiguous HEIF ICC profiles.");
                icc = property;
            } else if (kind == "nclx") {
                Require(nclx == null && property.PayloadLength == 11, "Invalid HEIF nclx property.");
                nclx = property;
            } else if (property.IsEssential) throw new NotSupportedException("Unsupported HEIF color property: " + kind + ".");
        }
        if (icc != null) {
            int length = icc.PayloadLength - 4;
            Require(length > 0 && length <= 4 * 1024 * 1024, "HEIF ICC profile exceeds the profile limit.");
            long rgbBytes = (long)image.Width * image.Height * 3;
            Budget(image.PixelBuffer.LongLength * 2 + rgbBytes + length * 33L + 4096);
            var profile = new byte[length];
            Buffer.BlockCopy(_bytes, icc.PayloadOffset + 4, profile, 0, length);
            var rgb = new byte[checked((int)rgbBytes)];
            for (int i = 0, j = 0; i < image.PixelBuffer.Length; i += 4, j += 3) {
                if ((i & 16383) == 0) _options.CancellationToken.ThrowIfCancellationRequested();
                Buffer.BlockCopy(image.PixelBuffer, i, rgb, j, 3);
            }
            OfficeIccRasterConversionStatus status = OfficeIccRasterConverter.TryConvertToSrgb(rgb, image.Width, image.Height, profile,
                new OfficeIccRasterConversionOptions {
                    MaximumPixels = _options.MaximumDecodedPixels,
                    MaximumManagedBytes = OfficeRasterGuards.MaximumDecodedBytes - _retained - image.PixelBuffer.LongLength,
                    CancellationToken = _options.CancellationToken
                }, out var converted);
            Require(status == OfficeIccRasterConversionStatus.Converted && converted != null, "HEIF ICC conversion failed: " + status + ".");
            for (int i = 0; i < image.PixelBuffer.Length; i += 4) {
                if ((i & 16383) == 0) _options.CancellationToken.ThrowIfCancellationRequested();
                Buffer.BlockCopy(converted!.PixelBuffer, i, image.PixelBuffer, i, 3);
            }
        } else if (nclx != null) {
            int p = nclx.PayloadOffset + 4;
            OfficeAvifColorConverter.ConvertRgbToSrgb(image, (ushort)U16(_bytes, p), (ushort)U16(_bytes, p + 2), _options.CancellationToken);
        }
    }
}
