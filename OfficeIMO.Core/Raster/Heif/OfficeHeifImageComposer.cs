using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeHeifImageComposer {
    internal static bool TryDecode(byte[] bytes, OfficeRasterDecodeOptions options, out OfficeRasterImage? image,
        out OfficeRasterCodecAvailability availability, out string? diagnostic,
        out OfficeImageOrientation sourceOrientation, out bool orientationNormalized) {
        image = null;
        diagnostic = null;
        sourceOrientation = OfficeImageOrientation.Normal;
        orientationNormalized = false;
        availability = new OfficeRasterCodecAvailability(OfficeRasterCodecUnavailableReason.None);
        try {
            if (!TryCreate(bytes, options, out var composer)) return false;
            Layout layout = composer!.ReadLayout(composer._info.PrimaryItem!, new HashSet<uint>(), 0);
            sourceOrientation = (OfficeImageOrientation)layout.SourceExif;
            IOfficeRasterImageCodec? codec = OfficeRasterImageCodecs.Resolve(options.ImageCodec);
            if (codec == null) {
                availability = new OfficeRasterCodecAvailability(OfficeRasterCodecUnavailableReason.NoPlatformCodec);
                diagnostic = availability.Message;
                return false;
            }
            if (codec is IOfficeRasterImageCodecAvailability reporting) {
                availability = reporting.GetAvailability();
                options.CancellationToken.ThrowIfCancellationRequested();
                if (!availability.IsAvailable) { diagnostic = availability.Message; return false; }
            }
            image = composer.Decode(layout, codec, 0);
            // auxl points from the auxiliary item to its master; HDR/depth auxiliaries are not alpha.
            uint masterId = layout.Item.ItemType == "tmap" ? layout.Tiles[0].ItemId : layout.Item.ItemId;
            var alphaItems = composer._info.References.Where(r => r.ReferenceType == "auxl" && r.ToItemIds.Contains(masterId))
                .Select(r => composer.FindItem(r.FromItemId))
                .Where(i => i.AuxiliaryType is "urn:mpeg:hevc:2015:auxid:1" or "urn:mpeg:mpegB:cicp:systems:auxiliary:alpha").ToArray();
            Require(alphaItems.Length <= 1, "HEIF primary image has ambiguous alpha auxiliaries.");
            if (alphaItems.Length == 1) {
                Layout alphaLayout = composer.ReadLayout(alphaItems[0], new HashSet<uint>(), 0);
                Require(alphaLayout.Width == image.Width && alphaLayout.Height == image.Height, "HEIF alpha dimensions do not match the primary image.");
                composer._retained = checked(composer._retained + image.PixelBuffer.LongLength);
                OfficeRasterImage alpha = composer.Decode(alphaLayout, codec, 0);
                composer.ApplyAlpha(image, alpha, composer.Targets(masterId, "prem").Contains(alphaItems[0].ItemId));
                composer._retained -= image.PixelBuffer.LongLength;
            }
            // EXIF is a display fallback for the complete image, including its auxiliary alpha.
            // Container properties were already applied independently to the referenced items.
            if (layout.Exif != 1) image = composer.Orient(image, layout.Exif);
            composer.ApplyColor(image, layout.Item.ItemType == "tmap" ? layout.Tiles[0] : layout.Item);
            options.CancellationToken.ThrowIfCancellationRequested();
            orientationNormalized = layout.Exif != 1;
            return true;
        } catch (Exception e) when (e is FormatException || e is OverflowException || e is ArgumentException || e is NotSupportedException ||
            e is System.IO.IOException || e is InvalidOperationException) {
            image = null;
            diagnostic = e.Message;
            return false;
        }
    }

    private OfficeRasterImage Decode(Layout layout, IOfficeRasterImageCodec codec, int depth) {
        _options.CancellationToken.ThrowIfCancellationRequested();
        Require(depth < 16, "Excessive HEIF derivation depth.");
        OfficeRasterImage image;
        if (layout.Item.ItemType == "hvc1") {
            _workPixels = checked(_workPixels + (long)layout.StoredWidth * layout.StoredHeight);
            Require(_workPixels <= _options.MaximumInspectionWorkPixels, "HEIF tile decoding exceeds the work-pixel limit.");
            int payloadLength = PayloadLength(layout.Item);
            // Account for synthesized container, native-provider input copies and RGBA copies before calling out.
            Budget(checked(payloadLength * 3L + layout.Item.CodecConfigurationBytes.Count * 3L +
                (long)layout.StoredWidth * layout.StoredHeight * 12 + 65536));
            byte[] tile = Synthesize(layout.Item, layout.StoredWidth, layout.StoredHeight);
            Require(codec.TryDecode(tile, "image/heic", out var decoded) && decoded != null,
                "The operating-system codec could not decode this HEIC still item.");
            _options.CancellationToken.ThrowIfCancellationRequested();
            Require(decoded!.Width == layout.StoredWidth && decoded.Height == layout.StoredHeight,
                "The operating-system codec returned unexpected HEIC tile dimensions.");
            image = decoded;
        } else if (layout.Item.ItemType is "iden" or "tmap") {
            image = Decode(ReadLayout(layout.Tiles[0], new HashSet<uint>(), depth + 1), codec, depth + 1);
        } else {
            Budget((long)layout.StoredWidth * layout.StoredHeight * 4);
            image = new OfficeRasterImage(layout.StoredWidth, layout.StoredHeight);
            _retained = checked(_retained + image.PixelBuffer.LongLength);
            try {
                int tileWidth = 0, tileHeight = 0;
                for (int i = 0; i < layout.Tiles.Length; i++) {
                    var child = ReadLayout(layout.Tiles[i], new HashSet<uint>(), depth + 1);
                    OfficeRasterImage tile = Decode(child, codec, depth + 1);
                    if (i == 0) { tileWidth = tile.Width; tileHeight = tile.Height; }
                    int x = i % layout.Columns * tileWidth, y = i / layout.Columns * tileHeight;
                    int copiedWidth = Math.Min(tile.Width, image.Width - x), copiedHeight = Math.Min(tile.Height, image.Height - y);
                    for (int row = 0; row < copiedHeight; row++) {
                        _options.CancellationToken.ThrowIfCancellationRequested();
                        Buffer.BlockCopy(tile.PixelBuffer, row * tile.Width * 4, image.PixelBuffer, ((y + row) * image.Width + x) * 4, copiedWidth * 4);
                    }
                }
            } finally { _retained -= image.PixelBuffer.LongLength; }
        }
        return ApplyTransforms(image, layout);
    }

    private OfficeRasterImage ApplyTransforms(OfficeRasterImage image, Layout layout) {
        foreach (var property in layout.Item.PropertyAssociations) {
            _options.CancellationToken.ThrowIfCancellationRequested();
            if (property.PropertyType == "clap") {
                Crop crop = ReadCrop(property, image.Width, image.Height);
                Budget(image.PixelBuffer.LongLength + (long)crop.Width * crop.Height * 4);
                var cropped = new OfficeRasterImage(crop.Width, crop.Height);
                for (int row = 0; row < crop.Height; row++) {
                    _options.CancellationToken.ThrowIfCancellationRequested();
                    Buffer.BlockCopy(image.PixelBuffer, ((crop.Top + row) * image.Width + crop.Left) * 4,
                        cropped.PixelBuffer, row * crop.Width * 4, crop.Width * 4);
                }
                image = cropped;
            } else if (property.PropertyType is "irot" or "imir") {
                int orientation = property.PropertyType == "imir" ? (_bytes[property.PayloadOffset] == 0 ? 4 : 2) :
                    new[] { 1, 8, 3, 6 }[_bytes[property.PayloadOffset]];
                image = Orient(image, orientation);
            }
        }
        return image;
    }

    private OfficeRasterImage Orient(OfficeRasterImage image, int orientation) {
        if (orientation == 1) return image;
        Budget(image.PixelBuffer.LongLength * 2);
        int width = image.Width, height = image.Height;
        byte[] rgba = OfficeRasterOrientation.Apply(image.PixelBuffer, ref width, ref height, orientation,
            _options.CancellationToken, "HEIF orientation exceeds the pixel limit.");
        return OfficeRasterImage.FromOwnedRgba32(width, height, rgba);
    }

    private Crop ReadCrop(OfficeHeifItemPropertyAssociationInfo property, int width, int height) {
        int p = property.PayloadOffset;
        decimal apertureWidth = Rational(p, false), apertureHeight = Rational(p + 8, false);
        decimal left = (width - apertureWidth) / 2 + Rational(p + 16, true);
        decimal top = (height - apertureHeight) / 2 + Rational(p + 24, true);
        Require(apertureWidth > 0 && apertureHeight > 0 && left >= 0 && top >= 0 && left + apertureWidth <= width && top + apertureHeight <= height,
            "HEIF clean aperture extends beyond its source image.");
        Require(decimal.Truncate(apertureWidth) == apertureWidth && decimal.Truncate(apertureHeight) == apertureHeight &&
            decimal.Truncate(left) == left && decimal.Truncate(top) == top, "HEIF clean aperture must align to the source pixel grid.");
        return new Crop((int)left, (int)top, (int)apertureWidth, (int)apertureHeight);
    }

    private decimal Rational(int p, bool signed) {
        uint denominator = U32(_bytes, p + 4);
        Require(denominator != 0, "HEIF clean aperture has a zero denominator.");
        return (signed ? (decimal)unchecked((int)U32(_bytes, p)) : U32(_bytes, p)) / denominator;
    }

    private void ApplyAlpha(OfficeRasterImage image, OfficeRasterImage alpha, bool premultiplied) {
        for (int i = 0; i < image.PixelBuffer.Length; i += 4) {
            if ((i & 16383) == 0) _options.CancellationToken.ThrowIfCancellationRequested();
            byte a = alpha.PixelBuffer[i];
            if (premultiplied) for (int c = 0; c < 3; c++)
                image.PixelBuffer[i + c] = a == 0 ? (byte)0 : (byte)Math.Min(255, (image.PixelBuffer[i + c] * 255 + a / 2) / a);
            image.PixelBuffer[i + 3] = a;
        }
    }

    private readonly struct Crop {
        internal Crop(int left, int top, int width, int height) { Left = left; Top = top; Width = width; Height = height; }
        internal int Left { get; }
        internal int Top { get; }
        internal int Width { get; }
        internal int Height { get; }
    }
}
