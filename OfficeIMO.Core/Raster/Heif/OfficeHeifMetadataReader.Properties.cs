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
        private void TryApplyImageProperties(byte[] data, Box itemPropertiesBox, List<HeifItemInfoBuilder> itemBuilders) {
            Box? itemPropertyContainerBox = null;
            var associations = new Dictionary<uint, List<HeifItemPropertyAssociation>>();

            foreach (Box childBox in EnumerateBoxes(data, itemPropertiesBox.DataOffset, itemPropertiesBox.EndOffset)) {
                CheckWork();
                if (childBox.Type == "ipco") {
                    itemPropertyContainerBox = childBox;
                } else if (childBox.Type == "ipma") {
                    Dictionary<uint, List<HeifItemPropertyAssociation>> declared = ReadItemPropertyAssociations(data, childBox);
                    foreach (KeyValuePair<uint, List<HeifItemPropertyAssociation>> entry in declared) {
                        CheckWork();
                        // Multiple ipma boxes contribute entries in file order. Keep the first
                        // entry for a repeated item, while validating every later declaration.
                        if (!associations.ContainsKey(entry.Key)) {
                            associations.Add(entry.Key, entry.Value);
                        }
                    }
                }
            }

            if (itemPropertyContainerBox is null || associations.Count == 0) {
                return;
            }

            Dictionary<int, HeifImageProperty> imageProperties = ReadImageProperties(data, itemPropertyContainerBox.Value);

            Dictionary<uint, HeifItemInfoBuilder> buildersById = itemBuilders.ToDictionary(item => item.ItemId);
            foreach (KeyValuePair<uint, List<HeifItemPropertyAssociation>> association in associations) {
                CheckWork();
                if (!buildersById.TryGetValue(association.Key, out HeifItemInfoBuilder? itemBuilder)) {
                    continue;
                }

                var publicAssociations = new List<OfficeHeifItemPropertyAssociationInfo>(association.Value.Count);
                foreach (HeifItemPropertyAssociation propertyAssociation in association.Value) {
                    CheckWork();
                    if (!imageProperties.TryGetValue(propertyAssociation.PropertyIndex, out HeifImageProperty imageProperty)) {
                        continue;
                    }

                    publicAssociations.Add(new OfficeHeifItemPropertyAssociationInfo(propertyAssociation.PropertyIndex, imageProperty.PropertyType, propertyAssociation.IsEssential));
                    if (imageProperty.Width.HasValue) {
                        itemBuilder.Width = imageProperty.Width;
                    }

                    if (imageProperty.Height.HasValue) {
                        itemBuilder.Height = imageProperty.Height;
                    }

                    if (imageProperty.RotationDegrees.HasValue) {
                        itemBuilder.RotationDegrees = imageProperty.RotationDegrees;
                    }

                    if (imageProperty.IsMirrored) {
                        itemBuilder.IsMirrored = true;
                    }

                    if (imageProperty.PixelAspectRatioHorizontalSpacing.HasValue) {
                        itemBuilder.PixelAspectRatioHorizontalSpacing = imageProperty.PixelAspectRatioHorizontalSpacing;
                    }

                    if (imageProperty.PixelAspectRatioVerticalSpacing.HasValue) {
                        itemBuilder.PixelAspectRatioVerticalSpacing = imageProperty.PixelAspectRatioVerticalSpacing;
                    }

                    if (imageProperty.PixelBitDepths is not null) {
                        itemBuilder.PixelBitDepths = imageProperty.PixelBitDepths;
                    }

                    if (imageProperty.ColorType is not null) {
                        itemBuilder.ColorType = imageProperty.ColorType;
                        itemBuilder.ColorPrimaries = imageProperty.ColorPrimaries;
                        itemBuilder.TransferCharacteristics = imageProperty.TransferCharacteristics;
                        itemBuilder.MatrixCoefficients = imageProperty.MatrixCoefficients;
                        itemBuilder.FullRangeFlag = imageProperty.FullRangeFlag;
                    }

                    if (imageProperty.AuxiliaryType is not null) {
                        itemBuilder.AuxiliaryType = imageProperty.AuxiliaryType;
                        itemBuilder.AuxiliarySubtypes = imageProperty.AuxiliarySubtypes ?? Array.Empty<byte>();
                    }

                    if (imageProperty.CodecConfigurationType is not null) {
                        itemBuilder.CodecConfigurationType = imageProperty.CodecConfigurationType;
                        itemBuilder.CodecConfigurationBytes = imageProperty.CodecConfigurationBytes ?? Array.Empty<byte>();
                    }
                }

                itemBuilder.PropertyAssociations = publicAssociations;
            }
        }

        private Dictionary<int, HeifImageProperty> ReadImageProperties(byte[] data, Box itemPropertyContainerBox) {
            var imageProperties = new Dictionary<int, HeifImageProperty>();
            int propertyIndex = 1;
            foreach (Box propertyBox in EnumerateBoxes(data, itemPropertyContainerBox.DataOffset, itemPropertyContainerBox.EndOffset)) {
                CheckWork();
                imageProperties[propertyIndex] = HeifImageProperty.CreateUnknown(propertyBox.Type);
                if (propertyBox.Type == "ispe") {
                    int offset = propertyBox.DataOffset;
                    if (TryReadFullBoxHeader(data, offset, propertyBox.EndOffset, out _, out _, out offset) &&
                        TryReadUInt32(data, offset, propertyBox.EndOffset, out uint width) &&
                        TryReadUInt32(data, offset + 4, propertyBox.EndOffset, out uint height)) {
                        imageProperties[propertyIndex] = HeifImageProperty.CreateSpatialExtent(width, height);
                    }
                } else if (propertyBox.Type == "irot") {
                    if (propertyBox.DataOffset < propertyBox.EndOffset) {
                        imageProperties[propertyIndex] = HeifImageProperty.CreateRotation((data[propertyBox.DataOffset] & 0x03) * 90);
                    }
                } else if (propertyBox.Type == "imir") {
                    if (propertyBox.DataOffset < propertyBox.EndOffset) {
                        imageProperties[propertyIndex] = HeifImageProperty.CreateMirror();
                    }
                } else if (propertyBox.Type == "pasp") {
                    if (TryReadUInt32(data, propertyBox.DataOffset, propertyBox.EndOffset, out uint horizontalSpacing) &&
                        TryReadUInt32(data, propertyBox.DataOffset + 4, propertyBox.EndOffset, out uint verticalSpacing)) {
                        imageProperties[propertyIndex] = HeifImageProperty.CreatePixelAspectRatio(horizontalSpacing, verticalSpacing);
                    }
                } else if (propertyBox.Type == "pixi") {
                    if (TryReadPixelInformation(data, propertyBox, out List<byte>? pixelBitDepths)) {
                        imageProperties[propertyIndex] = HeifImageProperty.CreatePixelInformation(pixelBitDepths!);
                    }
                } else if (propertyBox.Type == "colr") {
                    if (TryReadColorInformation(data, propertyBox, out HeifImageProperty colorProperty)) {
                        imageProperties[propertyIndex] = colorProperty;
                    }
                } else if (propertyBox.Type == "auxC") {
                    if (TryReadAuxiliaryType(data, propertyBox, out string? auxiliaryType, out byte[]? auxiliarySubtypes)) {
                        imageProperties[propertyIndex] = HeifImageProperty.CreateAuxiliaryType(auxiliaryType!, auxiliarySubtypes!);
                    }
                } else if (propertyBox.Type == "hvcC" || propertyBox.Type == "av1C" || propertyBox.Type == "avcC") {
                    imageProperties[propertyIndex] = HeifImageProperty.CreateCodecConfiguration(
                        propertyBox.Type,
                        CopyRange(data, propertyBox.DataOffset, propertyBox.EndOffset - propertyBox.DataOffset));
                }

                propertyIndex++;
            }

            return imageProperties;
        }

        private bool TryReadPixelInformation(byte[] data, Box propertyBox, out List<byte>? pixelBitDepths) {
            pixelBitDepths = null;
            int offset = propertyBox.DataOffset;
            if (!TryReadFullBoxHeader(data, offset, propertyBox.EndOffset, out _, out _, out offset) ||
                offset + 1 > propertyBox.EndOffset) {
                return false;
            }

            int channelCount = data[offset++];
            if (channelCount < 0 || offset + channelCount > propertyBox.EndOffset) {
                return false;
            }

            pixelBitDepths = new List<byte>(channelCount);
            for (int index = 0; index < channelCount; index++) {
                CheckWork();
                pixelBitDepths.Add(data[offset++]);
            }

            return true;
        }

        private bool TryReadColorInformation(byte[] data, Box propertyBox, out HeifImageProperty colorProperty) {
            colorProperty = default;
            if (propertyBox.DataOffset + 4 > propertyBox.EndOffset) {
                return false;
            }

            string colorType = ReadAscii(data, propertyBox.DataOffset, 4);
            if (colorType == "nclx" || colorType == "nclc") {
                int offset = propertyBox.DataOffset + 4;
                if (!TryReadUInt16(data, offset, propertyBox.EndOffset, out ushort colorPrimaries) ||
                    !TryReadUInt16(data, offset + 2, propertyBox.EndOffset, out ushort transferCharacteristics) ||
                    !TryReadUInt16(data, offset + 4, propertyBox.EndOffset, out ushort matrixCoefficients)) {
                    return false;
                }

                bool? fullRangeFlag = null;
                if (colorType == "nclx" && offset + 7 <= propertyBox.EndOffset) {
                    fullRangeFlag = (data[offset + 6] & 0x80) != 0;
                }

                colorProperty = HeifImageProperty.CreateColorInformation(colorType, colorPrimaries, transferCharacteristics, matrixCoefficients, fullRangeFlag);
                return true;
            }

            colorProperty = HeifImageProperty.CreateColorInformation(colorType, null, null, null, null);
            return true;
        }

        private bool TryReadAuxiliaryType(byte[] data, Box propertyBox, out string? auxiliaryType, out byte[]? auxiliarySubtypes) {
            auxiliaryType = null;
            auxiliarySubtypes = null;

            int offset = propertyBox.DataOffset;
            if (!TryReadFullBoxHeader(data, offset, propertyBox.EndOffset, out _, out _, out offset) ||
                offset >= propertyBox.EndOffset) {
                return false;
            }

            auxiliaryType = ReadNullTerminatedString(data, offset, propertyBox.EndOffset, out offset);
            auxiliarySubtypes = offset < propertyBox.EndOffset
                ? CopyRange(data, offset, propertyBox.EndOffset - offset)
                : Array.Empty<byte>();
            return !string.IsNullOrEmpty(auxiliaryType);
        }

        private Dictionary<uint, List<HeifItemPropertyAssociation>> ReadItemPropertyAssociations(byte[] data, Box itemPropertyAssociationBox) {
            var associations = new Dictionary<uint, List<HeifItemPropertyAssociation>>();
            int offset = itemPropertyAssociationBox.DataOffset;
            if (!TryReadFullBoxHeader(data, offset, itemPropertyAssociationBox.EndOffset, out byte version, out uint flags, out offset) || version > 1) {
                throw new FormatException("Truncated HEIF property association collection.");
            }

            if (!TryReadUInt32(data, offset, itemPropertyAssociationBox.EndOffset, out uint entryCount)) {
                throw new FormatException("Truncated HEIF property association collection.");
            }

            offset += 4;
            bool associationUsesLargePropertyIndex = (flags & 1) == 1;
            if (entryCount > 4096) {
                throw new FormatException("HEIF item count exceeds structure limits.");
            }
            for (uint index = 0; index < entryCount; index++) {
                CheckWork();
                uint itemId;
                if (version < 1) {
                    if (!TryReadUInt16(data, offset, itemPropertyAssociationBox.EndOffset, out ushort shortItemId)) {
                        throw new FormatException("Truncated HEIF property association collection.");
                    }

                    itemId = shortItemId;
                    offset += 2;
                } else {
                    if (!TryReadUInt32(data, offset, itemPropertyAssociationBox.EndOffset, out itemId)) {
                        throw new FormatException("Truncated HEIF property association collection.");
                    }

                    offset += 4;
                }

                if (offset + 1 > itemPropertyAssociationBox.EndOffset) {
                    throw new FormatException("Truncated HEIF property association collection.");
                }

                int associationCount = data[offset++];
                var propertyAssociations = new List<HeifItemPropertyAssociation>();
                for (int associationIndex = 0; associationIndex < associationCount; associationIndex++) {
                    CheckWork();
                    if (associationUsesLargePropertyIndex) {
                        if (!TryReadUInt16(data, offset, itemPropertyAssociationBox.EndOffset, out ushort association)) {
                            throw new FormatException("Truncated HEIF property association collection.");
                        }

                        int propertyIndex = association & 0x7FFF;
                        if (propertyIndex > 0) {
                            propertyAssociations.Add(new HeifItemPropertyAssociation(propertyIndex, (association & 0x8000) != 0));
                        }

                        offset += 2;
                    } else {
                        if (offset + 1 > itemPropertyAssociationBox.EndOffset) {
                            throw new FormatException("Truncated HEIF property association collection.");
                        }

                        byte association = data[offset++];
                        int propertyIndex = association & 0x7F;
                        if (propertyIndex > 0) {
                            propertyAssociations.Add(new HeifItemPropertyAssociation(propertyIndex, (association & 0x80) != 0));
                        }
                    }
                }

                if (associations.ContainsKey(itemId)) {
                    throw new FormatException("Duplicate HEIF property association item identifier.");
                }
                associations[itemId] = propertyAssociations;
            }

            if (offset != itemPropertyAssociationBox.EndOffset) {
                throw new FormatException("HEIF property associations do not match the declared count.");
            }
            return associations;
        }

        private List<OfficeHeifItemReference> ReadItemReferences(byte[] data, Box itemReferenceBox) {
            var references = new List<OfficeHeifItemReference>();
            int offset = itemReferenceBox.DataOffset;
            if (!TryReadFullBoxHeader(data, offset, itemReferenceBox.EndOffset, out byte version, out _, out offset) || version > 1) {
                throw new FormatException("Invalid HEIF reference collection header.");
            }

            foreach (Box referenceTypeBox in EnumerateBoxes(data, offset, itemReferenceBox.EndOffset)) {
                CheckWork();
                int referenceOffset = referenceTypeBox.DataOffset;
                uint fromItemId;
                if (version == 0) {
                    if (!TryReadUInt16(data, referenceOffset, referenceTypeBox.EndOffset, out ushort shortFromItemId)) {
                        throw new FormatException("Truncated HEIF reference source identifier.");
                    }

                    fromItemId = shortFromItemId;
                    referenceOffset += 2;
                } else {
                    if (!TryReadUInt32(data, referenceOffset, referenceTypeBox.EndOffset, out fromItemId)) {
                        throw new FormatException("Truncated HEIF reference source identifier.");
                    }

                    referenceOffset += 4;
                }

                if (!TryReadUInt16(data, referenceOffset, referenceTypeBox.EndOffset, out ushort referenceCount)) {
                    throw new FormatException("Truncated HEIF reference count.");
                }

                referenceOffset += 2;
                int idBytes = version == 0 ? 2 : 4;
                if (referenceTypeBox.EndOffset - referenceOffset != referenceCount * idBytes) {
                    throw new FormatException("HEIF reference targets do not match the declared count.");
                }
                var toItemIds = new List<uint>();
                for (ushort index = 0; index < referenceCount; index++) {
                    CheckWork();
                    if (version == 0) {
                        if (!TryReadUInt16(data, referenceOffset, referenceTypeBox.EndOffset, out ushort shortToItemId)) {
                            throw new FormatException("Truncated HEIF reference target identifier.");
                        }

                        toItemIds.Add(shortToItemId);
                        referenceOffset += 2;
                    } else {
                        if (!TryReadUInt32(data, referenceOffset, referenceTypeBox.EndOffset, out uint toItemId)) {
                            throw new FormatException("Truncated HEIF reference target identifier.");
                        }

                        toItemIds.Add(toItemId);
                        referenceOffset += 4;
                    }
                }

                references.Add(new OfficeHeifItemReference(referenceTypeBox.Type, fromItemId, toItemIds));
            }

            return references;
        }

    }
}
