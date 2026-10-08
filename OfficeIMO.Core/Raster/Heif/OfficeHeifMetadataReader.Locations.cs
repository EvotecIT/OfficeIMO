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
        private Dictionary<uint, OfficeHeifItemLocationInfo> ReadItemLocations(byte[] data, Box itemLocationBox, Box? itemDataBox) {
            var locations = new Dictionary<uint, OfficeHeifItemLocationInfo>();
            int offset = itemLocationBox.DataOffset;
            if (!TryReadFullBoxHeader(data, offset, itemLocationBox.EndOffset, out byte version, out _, out offset) || version > 2 ||
                offset + 2 > itemLocationBox.EndOffset) {
                throw new FormatException("Truncated or unsupported HEIF location collection.");
            }

            int offsetSize = data[offset] >> 4;
            int lengthSize = data[offset] & 0x0F;
            offset++;
            int baseOffsetSize = data[offset] >> 4;
            int indexSize = version == 1 || version == 2
                ? data[offset] & 0x0F
                : 0;
            offset++;

            if (!IsSupportedFieldSize(offsetSize) || !IsSupportedFieldSize(lengthSize) || !IsSupportedFieldSize(baseOffsetSize) || !IsSupportedFieldSize(indexSize)) {
                throw new FormatException("Truncated or unsupported HEIF location collection.");
            }

            uint itemCount;
            if (version < 2) {
                if (!TryReadUInt16(data, offset, itemLocationBox.EndOffset, out ushort shortItemCount)) {
                    throw new FormatException("Truncated or unsupported HEIF location collection.");
                }

                itemCount = shortItemCount;
                offset += 2;
            } else {
                if (!TryReadUInt32(data, offset, itemLocationBox.EndOffset, out itemCount)) {
                    throw new FormatException("Truncated or unsupported HEIF location collection.");
                }

                offset += 4;
            }

            if (itemCount > 4096) {
                throw new FormatException("HEIF item count exceeds structure limits.");
            }
            var declaredItemIds = new HashSet<uint>();
            for (uint index = 0; index < itemCount; index++) {
                CheckWork();
                if (!TryReadIlocItem(data, itemLocationBox.EndOffset, version, offsetSize, lengthSize, baseOffsetSize, indexSize, ref offset, out IlocItem item)) {
                    throw new FormatException("Truncated or unsupported HEIF location collection.");
                }
                if (!declaredItemIds.Add(item.ItemId)) {
                    throw new FormatException("Duplicate HEIF location item identifier.");
                }

                if (!TryCreateLocationExtents(item, itemDataBox, out List<OfficeHeifItemExtentInfo> extents)) {
                    continue;
                }

                locations[item.ItemId] = new OfficeHeifItemLocationInfo(item.ConstructionMethod, item.DataReferenceIndex, item.BaseOffset, extents);
            }
            if (offset != itemLocationBox.EndOffset) {
                throw new FormatException("HEIF location collection does not match its declared count.");
            }
            return locations;
        }

        private bool TryCreateLocationExtents(IlocItem item, Box? itemDataBox, out List<OfficeHeifItemExtentInfo> extents) {
            extents = new List<OfficeHeifItemExtentInfo>(item.Extents.Count);
            foreach (ItemExtent extent in item.Extents) {
                CheckWork();
                int resolvedOffset = extent.Offset;
                if (item.ConstructionMethod == 1 && itemDataBox is not null) {
                    if (extent.Offset > int.MaxValue - itemDataBox.Value.DataOffset) {
                        return false;
                    }

                    resolvedOffset = itemDataBox.Value.DataOffset + extent.Offset;
                }

                extents.Add(new OfficeHeifItemExtentInfo(resolvedOffset, extent.Length));
            }

            return true;
        }

        private bool TryFindItemExtents(byte[] data, Box itemLocationBox, Box? itemDataBox, uint itemId, out IlocItem item) {
            item = default;

            int offset = itemLocationBox.DataOffset;
            if (!TryReadFullBoxHeader(data, offset, itemLocationBox.EndOffset, out byte version, out _, out offset) || version > 2) {
                return false;
            }

            if (offset + 2 > itemLocationBox.EndOffset) {
                return false;
            }

            int offsetSize = data[offset] >> 4;
            int lengthSize = data[offset] & 0x0F;
            offset++;
            int baseOffsetSize = data[offset] >> 4;
            int indexSize = version == 1 || version == 2
                ? data[offset] & 0x0F
                : 0;
            offset++;

            if (!IsSupportedFieldSize(offsetSize) || !IsSupportedFieldSize(lengthSize) || !IsSupportedFieldSize(baseOffsetSize) || !IsSupportedFieldSize(indexSize)) {
                return false;
            }

            uint itemCount;
            if (version < 2) {
                if (!TryReadUInt16(data, offset, itemLocationBox.EndOffset, out ushort shortItemCount)) {
                    return false;
                }

                itemCount = shortItemCount;
                offset += 2;
            } else {
                if (!TryReadUInt32(data, offset, itemLocationBox.EndOffset, out itemCount)) {
                    return false;
                }

                offset += 4;
            }

            if (itemCount > 4096) {
                throw new FormatException("HEIF item count exceeds structure limits.");
            }
            bool found = false;
            var itemIds = new HashSet<uint>();
            for (uint index = 0; index < itemCount; index++) {
                CheckWork();
                if (!TryReadIlocItem(data, itemLocationBox.EndOffset, version, offsetSize, lengthSize, baseOffsetSize, indexSize, ref offset, out IlocItem candidateItem)) {
                    return false;
                }
                if (!itemIds.Add(candidateItem.ItemId)) {
                    return false;
                }

                if (candidateItem.ItemId == itemId) {
                    if (candidateItem.DataReferenceIndex != 0) {
                        return false;
                    }

                    if (candidateItem.ConstructionMethod == 0) {
                        item = candidateItem;
                        found = true;
                        continue;
                    }

                    if (candidateItem.ConstructionMethod == 1 && itemDataBox is not null) {
                        var extents = new List<ItemExtent>(candidateItem.Extents.Count);
                        foreach (ItemExtent extent in candidateItem.Extents) {
                            CheckWork();
                            if (extent.Offset > int.MaxValue - itemDataBox.Value.DataOffset) {
                                return false;
                            }

                            extents.Add(new ItemExtent(
                                itemDataBox.Value.DataOffset + extent.Offset,
                                extent.Length,
                                extent.OffsetPosition,
                                extent.OffsetSize,
                                extent.LengthPosition,
                                extent.LengthSize));
                        }

                        item = new IlocItem(candidateItem.ItemId, candidateItem.ConstructionMethod, candidateItem.DataReferenceIndex, 0, extents);
                        found = true;
                        continue;
                    }

                    return false;
                }
            }

            return found && offset == itemLocationBox.EndOffset;
        }

        private bool TryReadIlocItem(byte[] data, int endOffset, byte version, int offsetSize, int lengthSize, int baseOffsetSize, int indexSize, ref int offset, out IlocItem item) {
            item = default;

            uint itemId;
            if (version < 2) {
                if (!TryReadUInt16(data, offset, endOffset, out ushort shortItemId)) {
                    return false;
                }

                itemId = shortItemId;
                offset += 2;
            } else {
                if (!TryReadUInt32(data, offset, endOffset, out itemId)) {
                    return false;
                }

                offset += 4;
            }

            ushort constructionMethod = 0;
            if (version == 1 || version == 2) {
                if (!TryReadUInt16(data, offset, endOffset, out ushort constructionMethodRaw)) {
                    return false;
                }

                constructionMethod = (ushort)(constructionMethodRaw & 0x000F);
                offset += 2;
            }

            if (!TryReadUInt16(data, offset, endOffset, out ushort dataReferenceIndex)) {
                return false;
            }

            offset += 2;

            if (!TryReadVariableUInt(data, offset, endOffset, baseOffsetSize, out ulong baseOffset, out offset)) {
                return false;
            }

            if (!TryReadUInt16(data, offset, endOffset, out ushort extentCount)) {
                return false;
            }

            offset += 2;

            var extents = new List<ItemExtent>();
            for (ushort index = 0; index < extentCount; index++) {
                CheckWork();
                if (indexSize > 0 && !TryReadVariableUInt(data, offset, endOffset, indexSize, out _, out offset)) {
                    return false;
                }

                int extentOffsetPosition = offset;
                if (!TryReadVariableUInt(data, offset, endOffset, offsetSize, out ulong extentOffset, out offset)) {
                    return false;
                }

                int extentLengthPosition = offset;
                if (!TryReadVariableUInt(data, offset, endOffset, lengthSize, out ulong extentLength, out offset)) {
                    return false;
                }

                if (extentOffset > ulong.MaxValue - baseOffset || baseOffset + extentOffset > int.MaxValue || extentLength > int.MaxValue) {
                    return false;
                }

                extents.Add(new ItemExtent((int)(baseOffset + extentOffset), (int)extentLength, extentOffsetPosition, offsetSize, extentLengthPosition, lengthSize));
            }

            item = new IlocItem(itemId, constructionMethod, dataReferenceIndex, baseOffset, extents);
            return true;
        }

        private bool TryCopyExtents(byte[] data, List<ItemExtent> extents, out byte[]? itemData) {
            itemData = null;
            if (extents.Count == 0) {
                return false;
            }

            int totalLength = 0;
            foreach (ItemExtent extent in extents) {
                CheckWork();
                if (extent.Offset < 0 ||
                    extent.Length < 0 ||
                    extent.Offset > data.Length ||
                    extent.Length > data.Length - extent.Offset ||
                    extent.Length > int.MaxValue - totalLength) {
                    return false;
                }

                totalLength += extent.Length;
            }

            if (totalLength > OfficeExifProfileCodec.MaximumProfileBytes + 10) {
                return false;
            }
            ReserveBytes(totalLength + 24L);
            itemData = new byte[totalLength];
            int destinationOffset = 0;
            foreach (ItemExtent extent in extents) {
                CheckWork();
                Buffer.BlockCopy(data, extent.Offset, itemData, destinationOffset, extent.Length);
                destinationOffset += extent.Length;
            }

            return true;
        }

    }
}
