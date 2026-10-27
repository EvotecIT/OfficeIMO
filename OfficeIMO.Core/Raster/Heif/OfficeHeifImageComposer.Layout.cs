using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing;

/// <summary>Interprets HEIF items and presentation properties; compressed pixels remain an OS-codec boundary.</summary>
internal sealed partial class OfficeHeifImageComposer {
    private readonly byte[] _bytes;
    private readonly OfficeRasterDecodeOptions _options;
    private readonly OfficeHeifImageInfo _info;
    private readonly Dictionary<uint, OfficeHeifItemInfo> _items;
    private long _retained;
    private long _workPixels;
    private int _workItems;

    private OfficeHeifImageComposer(byte[] bytes, OfficeRasterDecodeOptions options, OfficeHeifImageInfo info) {
        _bytes = bytes;
        _options = options;
        _info = info;
        _items = info.Items.ToDictionary(i => i.ItemId);
        _retained = checked(options.RetainedManagedBytes + bytes.LongLength + info.ParserAllocatedBytes + 65536);
        Budget(0);
    }

    internal static bool TryIdentify(byte[] bytes, OfficeRasterDecodeOptions options, out int width, out int height) {
        width = height = 0;
        try {
            if (!TryCreate(bytes, options, out var composer)) return false;
            Layout layout = composer!.ReadLayout(composer._info.PrimaryItem!, new HashSet<uint>(), 0);
            width = layout.Width;
            height = layout.Height;
            return true;
        } catch (Exception e) when (e is FormatException || e is OverflowException || e is ArgumentException || e is NotSupportedException) {
            return false;
        }
    }

    private static bool TryCreate(byte[] bytes, OfficeRasterDecodeOptions options, out OfficeHeifImageComposer? composer) {
        composer = null;
        options.CancellationToken.ThrowIfCancellationRequested();
        if (bytes.Length > options.MaximumEncodedBytes || !OfficeImageReader.HasHeicSignature(bytes, options.CancellationToken) ||
            !OfficeHeifMetadataReader.TryReadCompositionInfo(bytes, options, out var info) ||
            info?.PrimaryItem == null || !IsHeifBrand(info.MajorBrand) && !info.CompatibleBrands.Any(IsHeifBrand)) return false;
        composer = new OfficeHeifImageComposer(bytes, options, info);
        return true;
    }

    private static bool IsHeifBrand(string brand) => brand is "heic" or "heix" or "hevc" or "heim" or "heis" or "mif1" or "msf1";

    private Layout ReadLayout(OfficeHeifItemInfo item, HashSet<uint> visiting, int depth) {
        _options.CancellationToken.ThrowIfCancellationRequested();
        Require(depth < 16 && ++_workItems <= 65536 && visiting.Add(item.ItemId), "Cyclic or excessive HEIF image derivation.");
        _retained = checked(_retained + 256);
        Budget(0);
        Require(item.ItemProtectionIndex == 0 && string.IsNullOrEmpty(item.ContentEncoding), "Protected HEIF items cannot be decoded.");
        int width = checked((int)(item.Width ?? 0)), height = checked((int)(item.Height ?? 0));
        Pixels(width, height);
        ValidateProperties(item);
        int rows = 1, columns = 1;
        OfficeHeifItemInfo[] tiles = Array.Empty<OfficeHeifItemInfo>();
        if (item.ItemType == "hvc1") {
            Require(item.CodecConfigurationType == "hvcC" && item.CodecConfigurationBytes.Count >= 23,
                "HEIC image has no complete hvcC decoder configuration.");
            Require(item.CodecConfigurationBytes[0] == 1, "Unsupported HEIC decoder configuration version.");
            PayloadLength(item);
        } else if (item.ItemType == "grid") {
            int length = PayloadLength(item);
            Require(length == 8 || length == 12, "Invalid HEIF grid descriptor.");
            byte[] grid = ReadPayload(item);
            Require(grid[0] == 0 && (grid[1] & ~1) == 0 && length == ((grid[1] & 1) == 0 ? 8 : 12), "Unsupported HEIF grid version or flags.");
            rows = grid[2] + 1; columns = grid[3] + 1;
            int gridWidth = length == 8 ? U16(grid, 4) : checked((int)U32(grid, 4));
            int gridHeight = length == 8 ? U16(grid, 6) : checked((int)U32(grid, 8));
            Require(width == gridWidth && height == gridHeight, "HEIF grid and spatial dimensions disagree.");
            uint[] ids = Targets(item.ItemId, "dimg");
            Require(ids.Length == checked(rows * columns), "HEIF grid tile count does not match its descriptor.");
            tiles = ids.Select(FindItem).ToArray();
            int tileWidth = 0, tileHeight = 0;
            foreach (OfficeHeifItemInfo tile in tiles) {
                Layout child = ReadLayout(tile, visiting, depth + 1);
                if (tileWidth == 0) { tileWidth = child.Width; tileHeight = child.Height; }
                Require(child.Width == tileWidth && child.Height == tileHeight, "HEIF grid tiles must have equal dimensions.");
            }
            Require((long)(columns - 1) * tileWidth < width && width <= (long)columns * tileWidth &&
                (long)(rows - 1) * tileHeight < height && height <= (long)rows * tileHeight, "HEIF grid tiles do not cover the output canvas.");
        } else if (item.ItemType is "iden" or "tmap") {
            uint[] ids = Targets(item.ItemId, "dimg");
            Require(ids.Length == (item.ItemType == "iden" ? 1 : 2), "Invalid HEIF identity or tone-map derivation.");
            // A tone-map's first input is its base image. Decode that image only; the gain map stays inert.
            tiles = new[] { FindItem(ids[0]) };
            Layout child = ReadLayout(tiles[0], visiting, depth + 1);
            Require(child.Width == width && child.Height == height, "HEIF identity dimensions disagree.");
        } else {
            throw new NotSupportedException("Only HEVC still images and grid/identity derivations with an SDR base are supported.");
        }
        int storedWidth = width, storedHeight = height;
        foreach (var property in item.PropertyAssociations) {
            if (property.PropertyType == "clap") {
                Crop crop = ReadCrop(property, width, height);
                width = crop.Width; height = crop.Height;
            } else if (property.PropertyType == "irot" && (_bytes[property.PayloadOffset] & 1) != 0) {
                int swap = width; width = height; height = swap;
            }
        }
        int exif = 1;
        if (item.IsPrimary && !HasOrientation(item) && _options.ApplyExifOrientation) {
            exif = ReadExifOrientation();
            if (exif >= 5) { int swap = width; width = height; height = swap; }
        }
        visiting.Remove(item.ItemId);
        return new Layout(item, storedWidth, storedHeight, width, height, rows, columns, tiles, exif);
    }

    private void ValidateProperties(OfficeHeifItemInfo item) {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var property in item.PropertyAssociations) {
            _options.CancellationToken.ThrowIfCancellationRequested();
            string type = property.PropertyType;
            Require(property.PayloadOffset >= 0 && property.PayloadLength >= 0 && property.PayloadOffset <= _bytes.Length - property.PayloadLength,
                "HEIF property extends beyond the input.");
            if (type is "ispe" or "hvcC" or "irot" or "imir" or "clap" or "pixi" or "auxC")
                Require(seen.Add(type), "Duplicate HEIF image property.");
            if (type is "irot" or "imir") {
                Require(property.PayloadLength == 1 && (_bytes[property.PayloadOffset] & (type == "irot" ? 252 : 254)) == 0,
                    "Invalid HEIF orientation property.");
            } else if (type == "clap") Require(property.PayloadLength == 32, "Invalid HEIF clean aperture.");
            else if (type == "ispe") Require(property.PayloadLength == 12, "Invalid HEIF spatial extent.");
            else if (type == "pasp") Require(item.PixelAspectRatioHorizontalSpacing == item.PixelAspectRatioVerticalSpacing,
                "Non-square HEIF pixels are not supported.");
            else if (property.IsEssential && type is not ("hvcC" or "pixi" or "colr" or "auxC" or "ccst"))
                throw new NotSupportedException("Unsupported essential HEIF property: " + type + ".");
        }
    }

    private int ReadExifOrientation() {
        OfficeHeifItemInfo? exif = _info.ExifItem;
        if (exif == null || exif.ItemProtectionIndex != 0) return 1;
        byte[] payload = ReadPayload(exif);
        if (payload.Length < 4) return 1;
        int start = checked(4 + (int)U32(payload, 0));
        return start < payload.Length && OfficeImageOrientationNormalizer.TryReadExifOrientationPayload(payload, start, payload.Length - start, out var orientation)
            ? (int)orientation : 1;
    }

    private static bool HasOrientation(OfficeHeifItemInfo item) => item.PropertyAssociations.Any(p => p.PropertyType is "irot" or "imir");
    private uint[] Targets(uint id, string type) => _info.References.Where(r => r.FromItemId == id && r.ReferenceType == type).SelectMany(r => r.ToItemIds).ToArray();
    private OfficeHeifItemInfo FindItem(uint id) => _items.TryGetValue(id, out var item) ? item : throw new FormatException("HEIF references an undeclared item.");

    private int PayloadLength(OfficeHeifItemInfo item) {
        var location = item.Location;
        Require(location != null && (location.IsFileBacked || location.IsItemDataBoxBacked) && location.Extents.Count > 0,
            "HEIF image has no supported local payload.");
        int length = 0;
        foreach (var extent in location!.Extents) {
            Require(extent.Offset >= 0 && extent.Length > 0 && extent.Offset <= _bytes.Length - extent.Length, "HEIF item extends beyond the input.");
            length = checked(length + extent.Length);
        }
        Require(length <= _options.MaximumEncodedBytes, "HEIF item exceeds the encoded-byte limit.");
        return length;
    }

    private byte[] ReadPayload(OfficeHeifItemInfo item) {
        int length = PayloadLength(item);
        Budget(length);
        var payload = new byte[length];
        int offset = 0;
        foreach (var extent in item.Location!.Extents) {
            _options.CancellationToken.ThrowIfCancellationRequested();
            Buffer.BlockCopy(_bytes, extent.Offset, payload, offset, extent.Length);
            offset += extent.Length;
        }
        return payload;
    }

    private void Budget(long extra) => Require(extra >= 0 && _retained <= OfficeRasterGuards.MaximumDecodedBytes - extra, "HEIF composition exceeds the managed-byte limit.");
    private void Pixels(int width, int height) => Require(OfficeRasterGuards.TryEnsurePixelCount(width, height, _options.MaximumDecodedPixels, out _), "HEIF image exceeds the pixel limit.");
    private static void Require(bool condition, string message) { if (!condition) throw new FormatException(message); }
    private static int U16(byte[] bytes, int p) => bytes[p] << 8 | bytes[p + 1];
    private static uint U32(byte[] bytes, int p) => (uint)bytes[p] << 24 | (uint)bytes[p + 1] << 16 | (uint)bytes[p + 2] << 8 | bytes[p + 3];

    private sealed class Layout {
        internal Layout(OfficeHeifItemInfo item, int storedWidth, int storedHeight, int width, int height, int rows, int columns, OfficeHeifItemInfo[] tiles, int exif) {
            Item = item; StoredWidth = storedWidth; StoredHeight = storedHeight; Width = width; Height = height;
            Rows = rows; Columns = columns; Tiles = tiles; Exif = exif;
        }
        internal OfficeHeifItemInfo Item { get; }
        internal int StoredWidth { get; }
        internal int StoredHeight { get; }
        internal int Width { get; }
        internal int Height { get; }
        internal int Rows { get; }
        internal int Columns { get; }
        internal OfficeHeifItemInfo[] Tiles { get; }
        internal int Exif { get; }
    }
}
