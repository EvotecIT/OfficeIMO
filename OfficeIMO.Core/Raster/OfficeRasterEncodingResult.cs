using System;

namespace OfficeIMO.Drawing {
    /// <summary>Encoded raster bytes together with completed metadata projection evidence.</summary>
    /// <remarks>Pixel encoding failures and cancellation throw instead of producing a partial result.</remarks>
    public sealed class OfficeRasterEncodingResult {
        internal OfficeRasterEncodingResult(byte[] bytes, OfficeImageMetadata metadata, OfficeImageMetadataProfileKinds omittedProfiles) {
            EncodedBytes = bytes;
            Metadata = metadata;
            OmittedProfiles = omittedProfiles;
        }

        /// <summary>The independently owned output array. Repeated access returns the same mutable array.</summary>
        public byte[] EncodedBytes { get; }

        /// <summary>The independently mutable requested metadata projection applied by this operation.</summary>
        /// <remarks>Editing this snapshot does not rewrite EncodedBytes. Re-read EncodedBytes for exact native density after quantization.
        /// This snapshot describes primary-image metadata and does not represent every TIFF page or animation frame.</remarks>
        public OfficeImageMetadata Metadata { get; }

        /// <summary>Supplied profile families omitted because the emitted container has no supported carrier.</summary>
        /// <remarks>This is evidence from the completed operation, not a promise to preserve arbitrary container-specific metadata.</remarks>
        public OfficeImageMetadataProfileKinds OmittedProfiles { get; }

        /// <summary>Whether any supplied profile family was omitted.</summary>
        /// <remarks>This does not describe pixel compression loss, density quantization, or metadata outside supported profile families.</remarks>
        public bool HasMetadataLoss => OmittedProfiles != OfficeImageMetadataProfileKinds.None;

        /// <summary>Returns the owned output array when no supplied profile family was omitted, otherwise throws.</summary>
        /// <remarks>This checks profile-family omissions only. JPEG or WebP pixel compression may still be lossy.</remarks>
        public byte[] RequireMetadataPreservation() {
            if (HasMetadataLoss) throw new InvalidOperationException($"The encoded container omitted metadata profile families: {OmittedProfiles}.");
            return EncodedBytes;
        }
    }
}
