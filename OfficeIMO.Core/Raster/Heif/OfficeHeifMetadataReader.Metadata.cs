// Adapted from the Evotec ImagePlayground HEIF metadata implementation.
// Copyright (c) 2022 Evotec. MIT license; see Licenses/ImagePlayground-LICENSE.txt.
// OfficeIMO adaptation provides bounded byte/stream APIs, cancellation and output preflight.
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeHeifMetadataReader {
    private sealed partial class Parser {
        private static readonly UTF8Encoding StrictXmpUtf8 = new UTF8Encoding(false, true);

        internal bool TryReadExifProfile(byte[] data, out OfficeImageMetadata? profile) {
            profile = null;

            if (!TryFindMetaBox(data, out Box metaBox)) {
                return false;
            }

            int metaChildrenStart = metaBox.DataOffset + 4;
            if (metaChildrenStart > metaBox.EndOffset) {
                return false;
            }

            Box? itemInfoBox = null;
            Box? itemLocationBox = null;
            Box? itemDataBox = null;

            foreach (Box childBox in EnumerateBoxes(data, metaChildrenStart, metaBox.EndOffset)) {
                CheckWork();
                if (childBox.Type == "iinf") {
                    itemInfoBox = childBox;
                } else if (childBox.Type == "iloc") {
                    itemLocationBox = childBox;
                } else if (childBox.Type == "idat") {
                    itemDataBox = childBox;
                }
            }

            if (itemInfoBox is null || itemLocationBox is null) {
                return false;
            }

            if (!TryFindExifItemId(data, itemInfoBox.Value, out uint itemId, requireSupportedPayload: true)) {
                return false;
            }

            if (!TryFindItemExtents(data, itemLocationBox.Value, itemDataBox, itemId, out IlocItem item)) {
                return false;
            }

            if (!TryCopyExtents(data, item.Extents, out byte[]? exifItemData)) {
                return false;
            }

            if (exifItemData!.Length == 0) {
                profile = null;
                return true;
            }

            if (!TryGetTiffPayload(exifItemData, out byte[]? tiffPayload)) {
                return false;
            }

            try {
                profile = OfficeImageMetadata.ParseExifProfile(tiffPayload!, checked(_sourceBytes + _allocatedBytes), _cancellationToken);
                return true;
            } catch (OperationCanceledException) {
                throw;
            } catch (FormatException) {
                profile = null;
                return false;
            }
        }

        internal bool TryReadXmp(byte[] data, out string? xmp) {
            xmp = null;

            if (!TryFindMetaBox(data, out Box metaBox)) {
                return false;
            }

            int metaChildrenStart = metaBox.DataOffset + 4;
            if (metaChildrenStart > metaBox.EndOffset) {
                return false;
            }

            Box? itemInfoBox = null;
            Box? itemLocationBox = null;
            Box? itemDataBox = null;

            foreach (Box childBox in EnumerateBoxes(data, metaChildrenStart, metaBox.EndOffset)) {
                CheckWork();
                if (childBox.Type == "iinf") {
                    itemInfoBox = childBox;
                } else if (childBox.Type == "iloc") {
                    itemLocationBox = childBox;
                } else if (childBox.Type == "idat") {
                    itemDataBox = childBox;
                }
            }

            if (itemInfoBox is null || itemLocationBox is null) {
                return false;
            }

            if (!TryFindXmpItemId(data, itemInfoBox.Value, out uint itemId, requireSupportedPayload: true)) {
                return false;
            }

            if (!TryFindItemExtents(data, itemLocationBox.Value, itemDataBox, itemId, out IlocItem item)) {
                return false;
            }

            if (!TryCopyExtents(data, item.Extents, out byte[]? itemData)) {
                return false;
            }

            ReserveBytes(itemData!.LongLength * 2L + 24L);
            xmp = StrictXmpUtf8.GetString(itemData!);
            return true;
        }

        internal bool HasExifItem(byte[] data) {
            return TryFindItemInfoBox(data, out Box itemInfoBox) &&
                   TryFindExifItemId(data, itemInfoBox, out _);
        }

        internal bool HasXmpItem(byte[] data) {
            return TryFindItemInfoBox(data, out Box itemInfoBox) &&
                   TryFindXmpItemId(data, itemInfoBox, out _);
        }

        private bool TryGetTiffPayload(byte[] exifItemData, out byte[]? tiffPayload) {
            tiffPayload = null;

            if (LooksLikeTiffHeader(exifItemData, 0)) {
                tiffPayload = exifItemData;
                return true;
            }

            if (exifItemData.Length >= 10) {
                uint tiffHeaderOffset = ReadUInt32(exifItemData, 0);
                ulong tiffStart = 4UL + tiffHeaderOffset;
                if (tiffStart <= (ulong)exifItemData.Length && tiffStart <= int.MaxValue && LooksLikeTiffHeader(exifItemData, (int)tiffStart)) {
                    tiffPayload = CopyTail(exifItemData, (int)tiffStart);
                    return true;
                }
            }

            if (exifItemData.Length >= 12 && ReadAscii(exifItemData, 0, 6) == "Exif\0\0" && LooksLikeTiffHeader(exifItemData, 6)) {
                tiffPayload = CopyTail(exifItemData, 6);
                return true;
            }

            return false;
        }

    }
}
