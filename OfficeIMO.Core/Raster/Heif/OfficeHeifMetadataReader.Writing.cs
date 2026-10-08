// Adapted from the Evotec ImagePlayground HEIF metadata implementation.
// Copyright (c) 2022 Evotec. MIT license; see Licenses/ImagePlayground-LICENSE.txt.
// OfficeIMO adaptation provides bounded byte/stream APIs, cancellation and output preflight.
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeHeifMetadataReader {
    private sealed partial class Parser {
        internal bool TryWriteExifProfile(byte[] data, OfficeImageMetadata? profile, out byte[]? result) {
            result = null;
            if (!TryFindMetaBox(data, out Box metaBox)) {
                return false;
            }

            int metaChildrenStart = metaBox.DataOffset + 4;
            if (metaChildrenStart > metaBox.EndOffset) {
                return false;
            }

            Box? itemInfoBox = null;
            Box? itemLocationBox = null;

            foreach (Box childBox in EnumerateBoxes(data, metaChildrenStart, metaBox.EndOffset)) {
                CheckWork();
                if (childBox.Type == "iinf") {
                    itemInfoBox = childBox;
                } else if (childBox.Type == "iloc") {
                    itemLocationBox = childBox;
                }
            }

            if (itemInfoBox is null || itemLocationBox is null) {
                return false;
            }

            if (!TryFindExifItemId(data, itemInfoBox.Value, out uint itemId, requireSupportedPayload: true)) {
                return false;
            }

            if (!TryFindItemExtents(data, itemLocationBox.Value, null, itemId, out IlocItem item) ||
                item.ConstructionMethod != 0 ||
                item.DataReferenceIndex != 0 ||
                item.Extents.Count != 1) {
                return false;
            }

            ItemExtent extent = item.Extents[0];
            if (!TryValidateWriterExtentOwnership(data, itemLocationBox.Value, item.ItemId, extent)) {
                return false;
            }
            byte[]? encodedExif = profile?.EncodeExifProfile(_cancellationToken, checked(_sourceBytes + _allocatedBytes));
            if (encodedExif != null) {
                ReserveBytes(encodedExif.LongLength + 24L);
            }
            byte[] exifItemData = encodedExif == null || encodedExif.Length == 0 ? Array.Empty<byte>() : CreateExifItemData(encodedExif);

            return TryWriteItemData(data, item, extent, out result, exifItemData);
        }

        internal bool TryWriteXmp(byte[] data, string? xmp, out byte[]? result) {
            result = null;
            if (!TryFindMetaBox(data, out Box metaBox)) {
                return false;
            }

            int metaChildrenStart = metaBox.DataOffset + 4;
            if (metaChildrenStart > metaBox.EndOffset) {
                return false;
            }

            Box? itemInfoBox = null;
            Box? itemLocationBox = null;

            foreach (Box childBox in EnumerateBoxes(data, metaChildrenStart, metaBox.EndOffset)) {
                CheckWork();
                if (childBox.Type == "iinf") {
                    itemInfoBox = childBox;
                } else if (childBox.Type == "iloc") {
                    itemLocationBox = childBox;
                }
            }

            if (itemInfoBox is null || itemLocationBox is null) {
                return false;
            }

            if (!TryFindXmpItemId(data, itemInfoBox.Value, out uint itemId, requireSupportedPayload: true)) {
                return false;
            }

            if (!TryFindItemExtents(data, itemLocationBox.Value, null, itemId, out IlocItem item) ||
                item.ConstructionMethod != 0 ||
                item.DataReferenceIndex != 0 ||
                item.Extents.Count != 1) {
                return false;
            }

            ItemExtent extent = item.Extents[0];
            if (!TryValidateWriterExtentOwnership(data, itemLocationBox.Value, item.ItemId, extent)) {
                return false;
            }
            if (xmp != null && System.Text.Encoding.UTF8.GetByteCount(xmp) > OfficeExifProfileCodec.MaximumProfileBytes) {
                return false;
            }
            byte[] xmpItemData = xmp is null
                ? Array.Empty<byte>()
                : System.Text.Encoding.UTF8.GetBytes(xmp);

            return TryWriteItemData(data, item, extent, out result, xmpItemData);
        }

        private bool TryWriteItemData(byte[] data, IlocItem item, ItemExtent extent, out byte[]? result, byte[] itemData) {
            result = null;
            bool isClearingItemData = itemData.Length == 0;
            int newOffset = isClearingItemData
                ? extent.Offset
                : checked(data.Length + 8);

            if ((ulong)newOffset < item.BaseOffset) {
                return false;
            }

            ulong extentOffset = (ulong)newOffset - item.BaseOffset;
            if (!CanWriteVariableUInt(extentOffset, extent.OffsetSize) ||
                !CanWriteVariableUInt((ulong)itemData.Length, extent.LengthSize)) {
                return false;
            }

            long outputLength = data.LongLength + (isClearingItemData ? 0L : 8L + itemData.LongLength);
            if (outputLength > OfficeRasterGuards.MaximumEncodedBytes ||
                itemData.Length > OfficeExifProfileCodec.MaximumProfileBytes + 10L ||
                checked(_sourceBytes + _allocatedBytes + itemData.LongLength + outputLength + 48L) > OfficeRasterGuards.MaximumDecodedBytes) {
                return false;
            }
            if (extent.Offset < 0 || extent.Length < 0 || extent.Offset > data.Length - extent.Length) {
                return false;
            }
            _cancellationToken.ThrowIfCancellationRequested();
            var output = new byte[(int)outputLength];
            CopyBytes(data, 0, output, 0, data.Length);
            if (extent.Length > 0) {
                for (int at = 0; at < extent.Length; at += 64 * 1024) {
                    _cancellationToken.ThrowIfCancellationRequested();
                    Array.Clear(output, extent.Offset + at, Math.Min(64 * 1024, extent.Length - at));
                }
            }

            WriteVariableUInt(output, extent.OffsetPosition, extent.OffsetSize, extentOffset);
            WriteVariableUInt(output, extent.LengthPosition, extent.LengthSize, (ulong)itemData.Length);

            if (itemData.Length > 0) {
                // A size-zero final box otherwise absorbs a new box appended after it.
                int position = 0;
                while (position < data.Length) {
                    CheckWork();
                    if (!TryReadBox(data, position, data.Length, out Box box)) {
                        return false;
                    }
                    if (ReadUInt32(data, position) == 0) {
                        WriteVariableUInt(output, position, 4, (ulong)(box.EndOffset - box.Offset));
                    }
                    position = box.EndOffset;
                }
                WriteVariableUInt(output, data.Length, 4, (ulong)(8 + itemData.Length));
                output[data.Length + 4] = (byte)'m';
                output[data.Length + 5] = (byte)'d';
                output[data.Length + 6] = (byte)'a';
                output[data.Length + 7] = (byte)'t';
                CopyBytes(itemData, 0, output, newOffset, itemData.Length);
            }

            _cancellationToken.ThrowIfCancellationRequested();
            result = output;
            return true;
        }

        // Clearing the old profile must not erase pixels or a different item's payload.
        // Scan structural locations without decoding any unrelated metadata family.
        private bool TryValidateWriterExtentOwnership(byte[] data, Box locationBox, uint itemId, ItemExtent target) {
            // Absolute offsets can address box structure as well as payloads. Only erase
            // nonempty extents within a media-data payload, never headers or metadata boxes.
            bool mediaDataContainsTarget = target.Length == 0;
            int topOffset = 0;
            while (topOffset < data.Length) {
                CheckWork();
                if (!TryReadBox(data, topOffset, data.Length, out Box box)) {
                    return false;
                }
                if (box.Type == "mdat" && target.Offset >= box.DataOffset &&
                    target.Offset <= box.EndOffset - target.Length) {
                    mediaDataContainsTarget = true;
                }
                topOffset = box.EndOffset;
            }
            if (!mediaDataContainsTarget) {
                return false;
            }
            int offset = locationBox.DataOffset;
            if (!TryReadFullBoxHeader(data, offset, locationBox.EndOffset, out byte version, out _, out offset) ||
                version > 2 || offset + 2 > locationBox.EndOffset) {
                return false;
            }
            int offsetSize = data[offset] >> 4;
            int lengthSize = data[offset++] & 0x0F;
            int baseOffsetSize = data[offset] >> 4;
            int indexSize = version == 1 || version == 2 ? data[offset] & 0x0F : 0;
            offset++;
            if (!IsSupportedFieldSize(offsetSize) || !IsSupportedFieldSize(lengthSize) ||
                !IsSupportedFieldSize(baseOffsetSize) || !IsSupportedFieldSize(indexSize)) {
                return false;
            }
            uint count;
            if (version < 2) {
                if (!TryReadUInt16(data, offset, locationBox.EndOffset, out ushort shortCount)) {
                    return false;
                }
                count = shortCount;
                offset += 2;
            } else {
                if (!TryReadUInt32(data, offset, locationBox.EndOffset, out count)) {
                    return false;
                }
                offset += 4;
            }
            if (count > 4096) {
                return false;
            }
            Box? idat = null;
            if (TryFindMetaBox(data, out Box meta)) {
                foreach (Box box in EnumerateBoxes(data, meta.DataOffset + 4, meta.EndOffset)) {
                    CheckWork();
                    if (box.Type == "idat") {
                        idat = box;
                    }
                }
            }
            var itemIds = new HashSet<uint>();
            bool foundTarget = false;
            for (uint index = 0; index < count; index++) {
                CheckWork();
                if (!TryReadIlocItem(data, locationBox.EndOffset, version, offsetSize, lengthSize,
                    baseOffsetSize, indexSize, ref offset, out IlocItem candidate) || !itemIds.Add(candidate.ItemId)) {
                    return false;
                }
                if (candidate.ItemId == itemId) {
                    foundTarget = true;
                    continue;
                }
                if (candidate.DataReferenceIndex != 0) {
                    continue;
                }
                if (candidate.ConstructionMethod > 1 || candidate.ConstructionMethod == 1 && idat == null) {
                    return false;
                }
                foreach (ItemExtent other in candidate.Extents) {
                    CheckWork();
                    long start = other.Offset + (candidate.ConstructionMethod == 1 ? idat!.Value.DataOffset : 0L);
                    long end = start + other.Length;
                    if (target.Length > 0 && other.Length > 0 && start < (long)target.Offset + target.Length && target.Offset < end) {
                        return false;
                    }
                }
            }
            return foundTarget;
        }

    }
}
