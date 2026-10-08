using System;
using System.Threading;

namespace OfficeIMO.Drawing {
    public static partial class OfficeRasterImageEncoder {
        /// <summary>Encodes pixels and applies supported primary-image metadata, reporting supplied profile families omitted by the output container.</summary>
        /// <remarks>Metadata and options are snapshotted. A nullable Resolution override takes precedence over metadata density;
        /// otherwise metadata supplies native density. Pixel encoding observes the requested byte ceiling as output is produced,
        /// and the complete rewritten output must also fit that ceiling. Metadata rewriting uses the existing global allocation
        /// limits. Cancellation and failure throw without returning a partial result. RequireMetadataPreservation checks profile omissions,
        /// not pixel compression fidelity. No ICC color transform is applied to pixels.</remarks>
        public static OfficeRasterEncodingResult EncodeWithMetadata(OfficeRasterImage image, OfficeImageExportFormat format,
            OfficeImageMetadata metadata, OfficeRasterEncodingOptions? options = null, long maximumEncodedBytes = 134217728L,
            CancellationToken cancellationToken = default) {
            if (image == null) throw new ArgumentNullException(nameof(image));
            if (metadata == null) throw new ArgumentNullException(nameof(metadata));
            OfficeImageMetadata.ValidateEncodingCeiling(maximumEncodedBytes);
            cancellationToken.ThrowIfCancellationRequested();
            OfficeImageMetadata snapshot = metadata.Clone();
            OfficeRasterEncodingOptions appliedOptions = (options ?? new OfficeRasterEncodingOptions()).Clone();
            OfficeRasterEncodingOptions effective = MetadataEncodingOptions(format, snapshot, appliedOptions);
            // Project before encoding so TIFF-relative opaque fields fail before pixel compression.
            _ = snapshot.PrepareForEncoding(format.GetContainerFormat(), out _);
            byte[] bytes = Encode(image, format, effective, maximumEncodedBytes, cancellationToken, snapshot.RetainedEncodingBytes);
            return snapshot.ApplyForEncoding(bytes, appliedOptions, maximumEncodedBytes, cancellationToken, image.PixelBuffer.LongLength + 24L);
        }

        /// <summary>Encodes a single still frame, TIFF pages, or icon entries with primary-image metadata and omission evidence.</summary>
        /// <remarks>TIFF pages use one shared resolution setting. Profile editing is limited to the primary TIFF image;
        /// this method does not promise per-page metadata preservation. Icon entries omit supplied metadata profile families.
        /// Animation durations and playback counts are not encoded by still formats; animated encoding belongs to the animation engine.
        /// Multi-frame input requires TIFF or Icon output. Existing frame-count, pixel, encoded-byte, and managed-memory limits apply.</remarks>
        public static OfficeRasterEncodingResult EncodeWithMetadata(OfficeRasterFrames frames, OfficeImageExportFormat format,
            OfficeImageMetadata metadata, OfficeRasterEncodingOptions? options = null, long maximumEncodedBytes = 134217728L,
            CancellationToken cancellationToken = default) {
            if (frames == null) throw new ArgumentNullException(nameof(frames));
            if (metadata == null) throw new ArgumentNullException(nameof(metadata));
            if (frames.Count == 1) return EncodeWithMetadata(frames[0].Image, format, metadata, options, maximumEncodedBytes, cancellationToken);
            if (format != OfficeImageExportFormat.Tiff && format != OfficeImageExportFormat.Icon) throw new NotSupportedException("Multiple frames require TIFF pages, icon entries, or an animation encoder.");
            OfficeImageMetadata.ValidateEncodingCeiling(maximumEncodedBytes);
            cancellationToken.ThrowIfCancellationRequested();
            OfficeImageMetadata snapshot = metadata.Clone();
            OfficeRasterEncodingOptions appliedOptions = (options ?? new OfficeRasterEncodingOptions()).Clone();
            OfficeRasterEncodingOptions effective = MetadataEncodingOptions(format, snapshot, appliedOptions).Resolve(format);
            _ = snapshot.PrepareForEncoding(format.GetContainerFormat(), out _);
            var images = new OfficeRasterImage[frames.Count];
            long retained = checked(frames.Count * 64L);
            for (int index = 0; index < images.Length; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                images[index] = frames[index].Image;
                retained = checked(retained + images[index].PixelBuffer.LongLength + 24L);
            }
            byte[] bytes = format == OfficeImageExportFormat.Tiff
                ? OfficeTiffCodec.EncodePages(images, effective.Tiff, cancellationToken, maximumEncodedBytes, snapshot.RetainedEncodingBytes)
                : OfficeIconEncoder.Encode(images, effective, cancellationToken, maximumEncodedBytes, snapshot.RetainedEncodingBytes);
            return snapshot.ApplyForEncoding(bytes, appliedOptions, maximumEncodedBytes, cancellationToken, retained);
        }

        private static OfficeRasterEncodingOptions MetadataEncodingOptions(OfficeImageExportFormat format, OfficeImageMetadata snapshot, OfficeRasterEncodingOptions? options) {
            OfficeRasterEncodingOptions effective = (options ?? new OfficeRasterEncodingOptions()).Clone();
            if (!format.IsRaster()) throw new ArgumentException("Metadata-aware encoding requires a managed raster output format.", nameof(format));
            if (effective.Resolution != null) snapshot.Resolution = effective.Resolution;
            effective.Resolution = snapshot.Resolution;
            // Unitless container density is written by the metadata operation after pixels are encoded.
            if (snapshot.Resolution.Unit == OfficeImageResolutionUnit.AspectRatio && format != OfficeImageExportFormat.Tiff) effective.WriteResolutionMetadata = false;
            return effective;
        }
    }
}
