namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Represents one picture bullet stored in a PPT9 document extension.</summary>
    public sealed class LegacyPptPictureBullet {
        private readonly byte[] _imageBytes;
        private readonly bool _isImagePayloadTruncated;

        internal LegacyPptPictureBullet(ushort index, byte preferredBlipType,
            byte unused, byte blipRecordVersion, ushort blipRecordInstance,
            ushort blipRecordType, uint blipPayloadLength,
            int blipPayloadAvailableLength, string? blipPayloadSha256,
            string? contentType, byte[] imageBytes,
            bool isImagePayloadTruncated) {
            Index = index;
            PreferredBlipType = preferredBlipType;
            Unused = unused;
            BlipRecordVersion = blipRecordVersion;
            BlipRecordInstance = blipRecordInstance;
            BlipRecordType = blipRecordType;
            BlipPayloadLength = blipPayloadLength;
            BlipPayloadAvailableLength = blipPayloadAvailableLength;
            BlipPayloadSha256 = blipPayloadSha256;
            ContentType = contentType;
            _imageBytes = imageBytes?.ToArray() ?? Array.Empty<byte>();
            _isImagePayloadTruncated = isImagePayloadTruncated;
        }

        /// <summary>Gets the zero-based index referenced by PPT9 paragraph properties.</summary>
        public ushort Index { get; }

        /// <summary>Gets the preferred Windows BLIP type byte.</summary>
        public byte PreferredBlipType { get; }

        /// <summary>Gets the ignored byte stored after the preferred BLIP type.</summary>
        public byte Unused { get; }

        /// <summary>Gets the embedded OfficeArt BLIP record version.</summary>
        public byte BlipRecordVersion { get; }

        /// <summary>Gets the embedded OfficeArt BLIP record instance.</summary>
        public ushort BlipRecordInstance { get; }

        /// <summary>Gets the embedded OfficeArt BLIP record type.</summary>
        public ushort BlipRecordType { get; }

        /// <summary>Gets the declared BLIP payload length.</summary>
        public uint BlipPayloadLength { get; }

        /// <summary>Gets the bounded BLIP payload length available in the source.</summary>
        public int BlipPayloadAvailableLength { get; }

        /// <summary>Gets the SHA-256 hash of the bounded raw BLIP payload.</summary>
        public string? BlipPayloadSha256 { get; }

        /// <summary>Gets the inferred image content type.</summary>
        public string? ContentType { get; }

        /// <summary>Gets a defensive copy of image bytes suitable for an Open XML image part.</summary>
        public byte[] ImageBytes => _imageBytes.ToArray();

        /// <summary>
        /// Gets whether the embedded BLIP or its declared metafile body is
        /// shorter than the bounded source.
        /// </summary>
        public bool IsPayloadTruncated => _isImagePayloadTruncated
            || BlipPayloadLength
            > unchecked((uint)BlipPayloadAvailableLength);

        /// <summary>Gets whether this picture bullet can be projected into DrawingML.</summary>
        public bool HasImportableImage => _imageBytes.Length > 0
            && !string.IsNullOrEmpty(ContentType) && !IsPayloadTruncated;
    }
}
