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
        internal bool TryReadInfo(byte[] data, out OfficeHeifImageInfo? info) {
            info = null;

            if (!TryReadFileType(data, out string majorBrand, out uint minorVersion, out List<string>? compatibleBrands)) {
                return false;
            }

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
            Box? itemPropertiesBox = null;
            Box? itemReferenceBox = null;
            uint? primaryItemId = null;

            foreach (Box childBox in EnumerateBoxes(data, metaChildrenStart, metaBox.EndOffset)) {
                CheckWork();
                if (childBox.Type == "iinf") {
                    itemInfoBox = childBox;
                } else if (childBox.Type == "iloc") {
                    itemLocationBox = childBox;
                } else if (childBox.Type == "idat") {
                    itemDataBox = childBox;
                } else if (childBox.Type == "iprp") {
                    itemPropertiesBox = childBox;
                } else if (childBox.Type == "iref") {
                    itemReferenceBox = childBox;
                } else if (childBox.Type == "pitm" && TryReadPrimaryItemId(data, childBox, out uint childPrimaryItemId)) {
                    primaryItemId = childPrimaryItemId;
                }
            }

            var itemBuilders = new List<HeifItemInfoBuilder>();
            if (itemInfoBox is not null && !TryReadItemInfos(data, itemInfoBox.Value, itemBuilders)) {
                return false;
            }

            if (itemPropertiesBox is not null) {
                TryApplyImageProperties(data, itemPropertiesBox.Value, itemBuilders);
            }

            Dictionary<uint, OfficeHeifItemLocationInfo> locations = itemLocationBox is not null
                ? ReadItemLocations(data, itemLocationBox.Value, itemDataBox)
                : new Dictionary<uint, OfficeHeifItemLocationInfo>();

            var items = itemBuilders
                .Select(item => new OfficeHeifItemInfo(
                    item.ItemId,
                    item.ItemType,
                    item.ItemName,
                    item.ItemProtectionIndex,
                    item.IsHidden,
                    item.MimeType,
                    item.ContentEncoding,
                    locations.TryGetValue(item.ItemId, out OfficeHeifItemLocationInfo? location) ? location : null,
                    item.PropertyAssociations,
                    primaryItemId.HasValue && item.ItemId == primaryItemId.Value,
                    item.ItemType == ExifItemType,
                    IsXmpMimeType(item.MimeType),
                    item.Width,
                    item.Height,
                    item.RotationDegrees,
                    item.IsMirrored,
                    item.PixelAspectRatioHorizontalSpacing,
                    item.PixelAspectRatioVerticalSpacing,
                    item.PixelBitDepths,
                    item.ColorType,
                    item.ColorPrimaries,
                    item.TransferCharacteristics,
                    item.MatrixCoefficients,
                    item.FullRangeFlag,
                    item.CodecConfigurationType,
                    item.CodecConfigurationBytes,
                    item.AuxiliaryType,
                    item.AuxiliarySubtypes))
                .ToList();
            List<OfficeHeifItemReference> references = itemReferenceBox is not null
                ? ReadItemReferences(data, itemReferenceBox.Value)
                : new List<OfficeHeifItemReference>();

            info = new OfficeHeifImageInfo(majorBrand, minorVersion, compatibleBrands!, primaryItemId, items, references);
            return true;
        }

    }
}
