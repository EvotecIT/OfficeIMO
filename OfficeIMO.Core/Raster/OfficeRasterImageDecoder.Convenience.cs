using System;
using System.IO;

namespace OfficeIMO.Drawing {
    public static partial class OfficeRasterImageDecoder {
        /// <summary>Decodes the requested static frame, throwing InvalidDataException for unsupported input, rejection, or limits.</summary>
        /// <remarks>Use TryDecode with evidence when failure is an expected best-effort result. Cancellation propagates.</remarks>
        public static OfficeRasterImage Decode(byte[] bytes, OfficeRasterDecodeOptions? options = null) => Decode(bytes, options, out _);

        /// <summary>Decodes the requested static frame and returns source, loss, and orientation evidence.</summary>
        public static OfficeRasterImage Decode(byte[] bytes, OfficeRasterDecodeOptions? options, out OfficeRasterDecodeInfo info) {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
            if (!TryDecode(bytes, options, out OfficeRasterImage? image, out info) || image == null) throw DecodeException(info);
            return image;
        }

        /// <summary>Decodes a static frame from a stream, leaving it open and restoring its position when seekable.</summary>
        public static OfficeRasterImage Decode(Stream stream, OfficeRasterDecodeOptions? options = null) => Decode(stream, options, out _);

        /// <summary>Decodes a static frame from a stream with source, loss, and orientation evidence.</summary>
        public static OfficeRasterImage Decode(Stream stream, OfficeRasterDecodeOptions? options, out OfficeRasterDecodeInfo info) {
            if (!TryDecode(stream, options, out OfficeRasterImage? image, out info) || image == null) throw DecodeException(info);
            return image;
        }

        /// <summary>Decodes every rendered frame or independent page within aggregate limits; cancellation propagates.</summary>
        public static OfficeRasterFrames DecodeFrames(byte[] bytes, OfficeRasterDecodeOptions? options = null, int maximumFrames = 256) => DecodeFrames(bytes, options, out _, maximumFrames);

        /// <summary>Decodes every frame or page and retains source descriptors and complete-sequence evidence.</summary>
        public static OfficeRasterFrames DecodeFrames(byte[] bytes, OfficeRasterDecodeOptions? options, out OfficeRasterDecodeInfo info, int maximumFrames = 256) {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
            if (!TryDecodeFrames(bytes, options, out OfficeRasterFrames? frames, out info, maximumFrames) || frames == null) throw DecodeException(info);
            return frames;
        }

        /// <summary>Decodes every frame or page from a stream, leaving it open and restoring its position when seekable.</summary>
        public static OfficeRasterFrames DecodeFrames(Stream stream, OfficeRasterDecodeOptions? options = null, int maximumFrames = 256) => DecodeFrames(stream, options, out _, maximumFrames);

        /// <summary>Decodes every stream frame or page with complete-sequence and source evidence.</summary>
        public static OfficeRasterFrames DecodeFrames(Stream stream, OfficeRasterDecodeOptions? options, out OfficeRasterDecodeInfo info, int maximumFrames = 256) {
            if (!TryDecodeFrames(stream, options, out OfficeRasterFrames? frames, out info, maximumFrames) || frames == null) throw DecodeException(info);
            return frames;
        }

        /// <summary>Decodes every stream frame or page within limits, retaining failure evidence and preserving seekable position.</summary>
        public static bool TryDecodeFrames(Stream stream, OfficeRasterDecodeOptions? options, out OfficeRasterFrames? frames,
            out OfficeRasterDecodeInfo info, int maximumFrames = 256) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (maximumFrames < 1 || maximumFrames > 4096) throw new ArgumentOutOfRangeException(nameof(maximumFrames));
            OfficeRasterDecodeOptions effective = options ?? new OfficeRasterDecodeOptions();
            effective.Validate();
            if (effective.FrameIndex != 0) throw new ArgumentException("All-frame decoding requires FrameIndex zero.", nameof(options));
            long position = stream.CanSeek ? stream.Position : 0L;
            try {
                if (!OfficeBoundedStreamReader.TryRead(stream, effective.MaximumEncodedBytes, effective.CancellationToken,
                        out byte[] bytes, out long retained)) {
                    frames = null;
                    info = new OfficeRasterDecodeInfo(OfficeImageFormat.Unknown, 0, 0, false, "The stream is empty, incomplete, or exceeds the encoded limit.");
                    return false;
                }
                return TryDecodeFrames(bytes, effective.WithAdditionalRetainedManagedBytes(retained), out frames, out info, maximumFrames);
            } finally {
                if (stream.CanSeek) stream.Position = position;
            }
        }

        private static InvalidDataException DecodeException(OfficeRasterDecodeInfo info) =>
            new InvalidDataException($"{info.Diagnostic ?? "Raster decoding failed."} [{info.Failure}; {info.Format}]");
    }
}
