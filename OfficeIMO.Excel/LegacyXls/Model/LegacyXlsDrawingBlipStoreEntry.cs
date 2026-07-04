namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes an OfficeArt FBSE image-store entry discovered in a legacy XLS drawing group.
    /// </summary>
    public sealed class LegacyXlsDrawingBlipStoreEntry {
        /// <summary>
        /// Creates preserve-only metadata for an OfficeArt FBSE image-store entry.
        /// </summary>
        public LegacyXlsDrawingBlipStoreEntry(
            ushort recordInstance,
            byte win32BlipType,
            byte macOsBlipType,
            string? uidHex,
            uint? sizeBytes,
            uint? referenceCount,
            ushort? embeddedBlipRecordType,
            uint? embeddedBlipPayloadLength,
            int? embeddedBlipPayloadAvailableLength,
            string? embeddedBlipPayloadSha256,
            byte[]? embeddedBlipPayloadBytes = null) {
            RecordInstance = recordInstance;
            RecordInstanceBlipTypeKind = TryGetBlipTypeKind(recordInstance);
            RecordInstanceBlipTypeName = GetBlipTypeName(recordInstance);
            Win32BlipType = win32BlipType;
            Win32BlipTypeKind = TryGetBlipTypeKind(win32BlipType);
            Win32BlipTypeName = GetBlipTypeName(win32BlipType);
            MacOsBlipType = macOsBlipType;
            MacOsBlipTypeKind = TryGetBlipTypeKind(macOsBlipType);
            MacOsBlipTypeName = GetBlipTypeName(macOsBlipType);
            UidHex = uidHex;
            SizeBytes = sizeBytes;
            ReferenceCount = referenceCount;
            EmbeddedBlipRecordType = embeddedBlipRecordType;
            EmbeddedBlipRecordTypeName = GetEmbeddedBlipRecordTypeName(embeddedBlipRecordType);
            EmbeddedBlipPayloadLength = embeddedBlipPayloadLength;
            EmbeddedBlipPayloadAvailableLength = embeddedBlipPayloadAvailableLength;
            EmbeddedBlipPayloadSha256 = embeddedBlipPayloadSha256;
            EmbeddedBlipPayloadBytes = embeddedBlipPayloadBytes?.ToArray() ?? Array.Empty<byte>();
            EmbeddedBlipContentType = GetEmbeddedBlipContentType(embeddedBlipRecordType, RecordInstanceBlipTypeKind, Win32BlipTypeKind, MacOsBlipTypeKind);
        }

        /// <summary>Gets the BLIP type value stored in the FBSE OfficeArt record instance field.</summary>
        public ushort RecordInstance { get; }

        /// <summary>Gets the typed BLIP value from <see cref="RecordInstance"/>, when known.</summary>
        public LegacyXlsDrawingBlipType? RecordInstanceBlipTypeKind { get; }

        /// <summary>Gets a stable display name for the FBSE record-instance BLIP type.</summary>
        public string RecordInstanceBlipTypeName { get; }

        /// <summary>Gets the Windows BLIP type byte.</summary>
        public byte Win32BlipType { get; }

        /// <summary>Gets the typed Windows BLIP value, when known.</summary>
        public LegacyXlsDrawingBlipType? Win32BlipTypeKind { get; }

        /// <summary>Gets a stable display name for the Windows BLIP type.</summary>
        public string Win32BlipTypeName { get; }

        /// <summary>Gets the MacOS BLIP type byte.</summary>
        public byte MacOsBlipType { get; }

        /// <summary>Gets the typed MacOS BLIP value, when known.</summary>
        public LegacyXlsDrawingBlipType? MacOsBlipTypeKind { get; }

        /// <summary>Gets a stable display name for the MacOS BLIP type.</summary>
        public string MacOsBlipTypeName { get; }

        /// <summary>Gets the FBSE image UID bytes as uppercase hexadecimal text, when available.</summary>
        public string? UidHex { get; }

        /// <summary>Gets the stored BLIP size from the FBSE entry, when available.</summary>
        public uint? SizeBytes { get; }

        /// <summary>Gets the FBSE reference count, when available.</summary>
        public uint? ReferenceCount { get; }

        /// <summary>Gets the embedded BLIP OfficeArt record type, when an embedded BLIP is present.</summary>
        public ushort? EmbeddedBlipRecordType { get; }

        /// <summary>Gets a stable display name for the embedded BLIP OfficeArt record type.</summary>
        public string? EmbeddedBlipRecordTypeName { get; }

        /// <summary>Gets the embedded BLIP payload length, when an embedded BLIP is present.</summary>
        public uint? EmbeddedBlipPayloadLength { get; }

        /// <summary>Gets the embedded BLIP payload bytes available in the drawing record.</summary>
        public int? EmbeddedBlipPayloadAvailableLength { get; }

        /// <summary>Gets the SHA-256 hash of the embedded BLIP payload bytes, when available.</summary>
        public string? EmbeddedBlipPayloadSha256 { get; }

        /// <summary>Gets a copy of importable embedded BLIP image bytes, when available.</summary>
        public byte[] EmbeddedBlipPayloadBytes { get; }

        /// <summary>Gets whether this entry carries importable embedded BLIP image bytes.</summary>
        public bool HasEmbeddedBlipPayloadBytes => EmbeddedBlipPayloadBytes.Length > 0;

        /// <summary>Gets the image content type inferred from the embedded BLIP record and BLIP type.</summary>
        public string? EmbeddedBlipContentType { get; }

        /// <summary>Gets whether this entry can be projected as an Open XML worksheet image.</summary>
        public bool HasImportableImagePayload => HasEmbeddedBlipPayloadBytes && !string.IsNullOrWhiteSpace(EmbeddedBlipContentType);

        private static LegacyXlsDrawingBlipType? TryGetBlipTypeKind(ushort value) {
            return value switch {
                0x00 => LegacyXlsDrawingBlipType.Error,
                0x01 => LegacyXlsDrawingBlipType.Unknown,
                0x02 => LegacyXlsDrawingBlipType.Emf,
                0x03 => LegacyXlsDrawingBlipType.Wmf,
                0x04 => LegacyXlsDrawingBlipType.Pict,
                0x05 => LegacyXlsDrawingBlipType.Jpeg,
                0x06 => LegacyXlsDrawingBlipType.Png,
                0x07 => LegacyXlsDrawingBlipType.Dib,
                0x11 => LegacyXlsDrawingBlipType.Tiff,
                0x12 => LegacyXlsDrawingBlipType.CmykJpeg,
                _ => null
            };
        }

        private static string GetBlipTypeName(ushort value) {
            return TryGetBlipTypeKind(value)?.ToString() ?? $"BlipType:0x{value:X2}";
        }

        private static string? GetEmbeddedBlipRecordTypeName(ushort? recordType) {
            if (!recordType.HasValue) {
                return null;
            }

            return recordType.Value switch {
                0xF01A => "OfficeArtBlipEMF",
                0xF01B => "OfficeArtBlipWMF",
                0xF01C => "OfficeArtBlipPICT",
                0xF01D => "OfficeArtBlipJPEG",
                0xF01E => "OfficeArtBlipPNG",
                0xF01F => "OfficeArtBlipDIB",
                0xF029 => "OfficeArtBlipTIFF",
                0xF02A => "OfficeArtBlipJPEG",
                _ => $"EmbeddedBlipRecordType:0x{recordType.Value:X4}"
            };
        }

        private static string? GetEmbeddedBlipContentType(
            ushort? embeddedRecordType,
            LegacyXlsDrawingBlipType? recordInstanceType,
            LegacyXlsDrawingBlipType? win32Type,
            LegacyXlsDrawingBlipType? macOsType) {
            LegacyXlsDrawingBlipType? type = GetBlipTypeFromEmbeddedRecord(embeddedRecordType)
                ?? recordInstanceType
                ?? win32Type
                ?? macOsType;

            return type switch {
                LegacyXlsDrawingBlipType.Png => "image/png",
                LegacyXlsDrawingBlipType.Jpeg => "image/jpeg",
                LegacyXlsDrawingBlipType.CmykJpeg => "image/jpeg",
                LegacyXlsDrawingBlipType.Dib => "image/bmp",
                LegacyXlsDrawingBlipType.Tiff => "image/tiff",
                LegacyXlsDrawingBlipType.Emf => "image/x-emf",
                LegacyXlsDrawingBlipType.Wmf => "image/x-wmf",
                _ => null
            };
        }

        private static LegacyXlsDrawingBlipType? GetBlipTypeFromEmbeddedRecord(ushort? recordType) {
            if (!recordType.HasValue) {
                return null;
            }

            return recordType.Value switch {
                0xF01A => LegacyXlsDrawingBlipType.Emf,
                0xF01B => LegacyXlsDrawingBlipType.Wmf,
                0xF01D => LegacyXlsDrawingBlipType.Jpeg,
                0xF01E => LegacyXlsDrawingBlipType.Png,
                0xF029 => LegacyXlsDrawingBlipType.Tiff,
                0xF02A => LegacyXlsDrawingBlipType.CmykJpeg,
                _ => null
            };
        }
    }
}
