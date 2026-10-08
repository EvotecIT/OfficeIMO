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
        private readonly struct Box {
            public Box(int offset, int endOffset, int dataOffset, string type) {
                Offset = offset;
                EndOffset = endOffset;
                DataOffset = dataOffset;
                Type = type;
            }

            public int Offset { get; }

            public int EndOffset { get; }

            public int DataOffset { get; }

            public string Type { get; }
        }

        private readonly struct IlocItem {
            public IlocItem(uint itemId, ushort constructionMethod, ushort dataReferenceIndex, ulong baseOffset, List<ItemExtent> extents) {
                ItemId = itemId;
                ConstructionMethod = constructionMethod;
                DataReferenceIndex = dataReferenceIndex;
                BaseOffset = baseOffset;
                Extents = extents;
            }

            public uint ItemId { get; }

            public ushort ConstructionMethod { get; }

            public ushort DataReferenceIndex { get; }

            public ulong BaseOffset { get; }

            public List<ItemExtent> Extents { get; }
        }

        private readonly struct ItemExtent {
            public ItemExtent(int offset, int length, int offsetPosition, int offsetSize, int lengthPosition, int lengthSize) {
                Offset = offset;
                Length = length;
                OffsetPosition = offsetPosition;
                OffsetSize = offsetSize;
                LengthPosition = lengthPosition;
                LengthSize = lengthSize;
            }

            public int Offset { get; }

            public int Length { get; }

            public int OffsetPosition { get; }

            public int OffsetSize { get; }

            public int LengthPosition { get; }

            public int LengthSize { get; }
        }

        private sealed class HeifItemInfoBuilder {
            public HeifItemInfoBuilder(uint itemId, string itemType, string itemName, ushort itemProtectionIndex, bool isHidden, string? mimeType, string? contentEncoding) {
                ItemId = itemId;
                ItemType = itemType;
                ItemName = itemName;
                ItemProtectionIndex = itemProtectionIndex;
                IsHidden = isHidden;
                MimeType = mimeType;
                ContentEncoding = contentEncoding;
            }

            public uint ItemId { get; }

            public string ItemType { get; }

            public string ItemName { get; }

            public ushort ItemProtectionIndex { get; }

            public bool IsHidden { get; }

            public string? MimeType { get; }

            public string? ContentEncoding { get; }

            public IReadOnlyList<OfficeHeifItemPropertyAssociationInfo> PropertyAssociations { get; set; } = Array.Empty<OfficeHeifItemPropertyAssociationInfo>();

            public uint? Width { get; set; }

            public uint? Height { get; set; }

            public int? RotationDegrees { get; set; }

            public bool IsMirrored { get; set; }

            public uint? PixelAspectRatioHorizontalSpacing { get; set; }

            public uint? PixelAspectRatioVerticalSpacing { get; set; }

            public IReadOnlyList<byte> PixelBitDepths { get; set; } = Array.Empty<byte>();

            public string? ColorType { get; set; }

            public ushort? ColorPrimaries { get; set; }

            public ushort? TransferCharacteristics { get; set; }

            public ushort? MatrixCoefficients { get; set; }

            public bool? FullRangeFlag { get; set; }

            public string? CodecConfigurationType { get; set; }

            public IReadOnlyList<byte> CodecConfigurationBytes { get; set; } = Array.Empty<byte>();

            public string? AuxiliaryType { get; set; }

            public IReadOnlyList<byte> AuxiliarySubtypes { get; set; } = Array.Empty<byte>();
        }

        private readonly struct HeifItemPropertyAssociation {
            public HeifItemPropertyAssociation(int propertyIndex, bool isEssential) {
                PropertyIndex = propertyIndex;
                IsEssential = isEssential;
            }

            public int PropertyIndex { get; }

            public bool IsEssential { get; }
        }

        private readonly struct HeifImageProperty {
            private HeifImageProperty(string propertyType, uint? width, uint? height, int? rotationDegrees, bool isMirrored, uint? pixelAspectRatioHorizontalSpacing, uint? pixelAspectRatioVerticalSpacing, IReadOnlyList<byte>? pixelBitDepths, string? colorType, ushort? colorPrimaries, ushort? transferCharacteristics, ushort? matrixCoefficients, bool? fullRangeFlag, string? codecConfigurationType, IReadOnlyList<byte>? codecConfigurationBytes, string? auxiliaryType, IReadOnlyList<byte>? auxiliarySubtypes) {
                PropertyType = propertyType;
                Width = width;
                Height = height;
                RotationDegrees = rotationDegrees;
                IsMirrored = isMirrored;
                PixelAspectRatioHorizontalSpacing = pixelAspectRatioHorizontalSpacing;
                PixelAspectRatioVerticalSpacing = pixelAspectRatioVerticalSpacing;
                PixelBitDepths = pixelBitDepths;
                ColorType = colorType;
                ColorPrimaries = colorPrimaries;
                TransferCharacteristics = transferCharacteristics;
                MatrixCoefficients = matrixCoefficients;
                FullRangeFlag = fullRangeFlag;
                CodecConfigurationType = codecConfigurationType;
                CodecConfigurationBytes = codecConfigurationBytes;
                AuxiliaryType = auxiliaryType;
                AuxiliarySubtypes = auxiliarySubtypes;
            }

            public static HeifImageProperty CreateUnknown(string propertyType) =>
                new(propertyType, null, null, null, false, null, null, null, null, null, null, null, null, null, null, null, null);

            public static HeifImageProperty CreateSpatialExtent(uint width, uint height) =>
                new("ispe", width, height, null, false, null, null, null, null, null, null, null, null, null, null, null, null);

            public static HeifImageProperty CreateRotation(int rotationDegrees) =>
                new("irot", null, null, rotationDegrees, false, null, null, null, null, null, null, null, null, null, null, null, null);

            public static HeifImageProperty CreateMirror() =>
                new("imir", null, null, null, true, null, null, null, null, null, null, null, null, null, null, null, null);

            public static HeifImageProperty CreatePixelAspectRatio(uint horizontalSpacing, uint verticalSpacing) =>
                new("pasp", null, null, null, false, horizontalSpacing, verticalSpacing, null, null, null, null, null, null, null, null, null, null);

            public static HeifImageProperty CreatePixelInformation(IReadOnlyList<byte> pixelBitDepths) =>
                new("pixi", null, null, null, false, null, null, pixelBitDepths, null, null, null, null, null, null, null, null, null);

            public static HeifImageProperty CreateColorInformation(string colorType, ushort? colorPrimaries, ushort? transferCharacteristics, ushort? matrixCoefficients, bool? fullRangeFlag) =>
                new("colr", null, null, null, false, null, null, null, colorType, colorPrimaries, transferCharacteristics, matrixCoefficients, fullRangeFlag, null, null, null, null);

            public static HeifImageProperty CreateAuxiliaryType(string auxiliaryType, IReadOnlyList<byte> auxiliarySubtypes) =>
                new("auxC", null, null, null, false, null, null, null, null, null, null, null, null, null, null, auxiliaryType, auxiliarySubtypes);

            public static HeifImageProperty CreateCodecConfiguration(string codecConfigurationType, IReadOnlyList<byte> codecConfigurationBytes) =>
                new(codecConfigurationType, null, null, null, false, null, null, null, null, null, null, null, null, codecConfigurationType, codecConfigurationBytes, null, null);

            public string PropertyType { get; }

            public uint? Width { get; }

            public uint? Height { get; }

            public int? RotationDegrees { get; }

            public bool IsMirrored { get; }

            public uint? PixelAspectRatioHorizontalSpacing { get; }

            public uint? PixelAspectRatioVerticalSpacing { get; }

            public IReadOnlyList<byte>? PixelBitDepths { get; }

            public string? ColorType { get; }

            public ushort? ColorPrimaries { get; }

            public ushort? TransferCharacteristics { get; }

            public ushort? MatrixCoefficients { get; }

            public bool? FullRangeFlag { get; }

            public string? CodecConfigurationType { get; }

            public IReadOnlyList<byte>? CodecConfigurationBytes { get; }

            public string? AuxiliaryType { get; }

            public IReadOnlyList<byte>? AuxiliarySubtypes { get; }
        }    }
}
