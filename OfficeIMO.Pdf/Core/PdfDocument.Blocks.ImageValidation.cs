using OfficeIMO.Drawing;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Threading;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    private const string SupportedImageMessage =
        "PdfDocument.Image accepts JPEG and the raster formats decoded by OfficeIMO.Drawing. JPEG and writer-safe PNG payloads are embedded directly; other supported raster payloads are normalized to PNG once before PDF serialization.";
    private static readonly ConditionalWeakTable<byte[], PreparedImageCacheEntry> PreparedImageCache = new();

    // Prepared images by source content: the array-keyed cache above misses whenever the same image arrives in a
    // new array (an OfficeDrawingImage or page background copies its source; the next document decodes its logo
    // afresh), and preparing a PNG inflates, unfilters and splits its alpha - large-object-heap churn on every
    // render. A prepared image is immutable once built, so equal content can share one.
    private static readonly PreparedImageContentCache PreparedImagesByContent = new();

    private sealed class PreparedImageContentCache {
        private const int MaximumEntries = 32;
        private const long MaximumBytes = 32L * 1024L * 1024L;
        private readonly object _sync = new();
        private readonly System.Collections.Generic.Dictionary<string, PreparedImage> _entries = new(System.StringComparer.Ordinal);
        private long _bytes;

        internal bool TryGet(byte[] sourceHash, out PreparedImage prepared) {
            lock (_sync) {
                return _entries.TryGetValue(System.Convert.ToBase64String(sourceHash), out prepared);
            }
        }

        internal void Add(byte[] sourceHash, PreparedImage prepared) {
            long size = prepared.Data.LongLength + (prepared.PreparedStream?.Data.LongLength ?? 0L) + (prepared.PreparedStream?.SoftMask?.Data.LongLength ?? 0L);
            if (size > MaximumBytes / 4) {
                return;
            }

            string key = System.Convert.ToBase64String(sourceHash);
            lock (_sync) {
                if (_entries.ContainsKey(key)) {
                    return;
                }

                // Clear-all eviction keeps this simple; an LRU would suit a process that cycles through more
                // distinct images than the cache holds.
                if (_entries.Count >= MaximumEntries || _bytes + size > MaximumBytes) {
                    _entries.Clear();
                    _bytes = 0;
                }

                _entries[key] = prepared;
                _bytes += size;
            }
        }
    }

    private sealed class PreparedImageCacheEntry {
        internal object Gate { get; } = new();
        internal byte[]? SourceHash { get; set; }
        internal PreparedImage Prepared { get; set; }
        internal bool HasPrepared { get; set; }
        internal bool DataIsSourceCopy { get; set; }
    }

    internal readonly struct PreparedImage {
        internal PreparedImage(
            byte[] data,
            OfficeImageInfo info,
            OfficeImageFormat sourceFormat,
            bool wasTranscoded,
            PdfWriter.PdfImageStream? preparedStream = null) {
            Data = data;
            Info = info;
            SourceFormat = sourceFormat;
            WasTranscoded = wasTranscoded;
            PreparedStream = preparedStream;
        }

        internal byte[] Data { get; }
        internal OfficeImageInfo Info { get; }
        internal OfficeImageFormat SourceFormat { get; }
        internal bool WasTranscoded { get; }
        internal PdfWriter.PdfImageStream? PreparedStream { get; }
    }

    /// <summary>
    /// Checks whether image bytes can be embedded by the first-party PDF writer.
    /// </summary>
    public static bool TryValidateImageBytes(byte[] data, out OfficeImageInfo? imageInfo, out string? unsupportedReason) {
        bool prepared = TryPrepareImageBytes(data, out _, out imageInfo, out _, out unsupportedReason);
        return prepared;
    }

    /// <summary>
    /// Prepares source image bytes for first-party PDF embedding. Writer-safe JPEG and PNG data is retained;
    /// other raster formats supported by <see cref="OfficeRasterImageDecoder"/> are normalized to PNG.
    /// </summary>
    public static bool TryPrepareImageBytes(
        byte[] data,
        out byte[] preparedBytes,
        out OfficeImageInfo? imageInfo,
        out bool wasTranscoded,
        out string? unsupportedReason) {
        preparedBytes = System.Array.Empty<byte>();
        imageInfo = null;
        wasTranscoded = false;
        unsupportedReason = null;
        try {
            PreparedImage prepared = PrepareImageBytes(data);
            preparedBytes = (byte[])prepared.Data.Clone();
            imageInfo = prepared.Info;
            wasTranscoded = prepared.WasTranscoded;
            return true;
        } catch (NotSupportedException ex) {
            unsupportedReason = ex.Message;
            return false;
        } catch (ArgumentException ex) {
            unsupportedReason = ex.Message;
            return false;
        }
    }

    internal static OfficeImageInfo ValidateImageBytes(byte[] data) => PrepareImageBytes(data).Info;

    internal static PreparedImage PrepareImageBytes(byte[] data) =>
        PrepareImageBytes(data, CancellationToken.None);

    internal static PreparedImage PrepareImageBytes(byte[] data, CancellationToken cancellationToken) {
        Guard.NotNullOrEmpty(data, nameof(data));
        cancellationToken.ThrowIfCancellationRequested();
        if (cancellationToken.CanBeCanceled) {
            return PrepareImageBytesCore(data, cancellationToken);
        }

        PreparedImageCacheEntry entry = PreparedImageCache.GetValue(data, static _ => new PreparedImageCacheEntry());
        lock (entry.Gate) {
            // Most prepared images keep a byte-for-byte copy of their source, and comparing against it is far
            // cheaper than hashing the source again to prove the caller has not changed the array since.
            if (entry.HasPrepared && entry.DataIsSourceCopy && BytesEqual(data, entry.Prepared.Data)) {
                return entry.Prepared;
            }
        }

        byte[] sourceHash = ComputeImageSourceHash(data);

        lock (entry.Gate) {
            if (entry.HasPrepared &&
                entry.SourceHash != null &&
                System.Linq.Enumerable.SequenceEqual(entry.SourceHash, sourceHash)) {
                return entry.Prepared;
            }

            if (!PreparedImagesByContent.TryGet(sourceHash, out PreparedImage prepared)) {
                prepared = PrepareImageBytesCore(data, cancellationToken);
                PreparedImagesByContent.Add(sourceHash, prepared);
            }
            entry.SourceHash = sourceHash;
            entry.Prepared = prepared;
            entry.DataIsSourceCopy = BytesEqual(data, prepared.Data);
            entry.HasPrepared = true;
            return prepared;
        }
    }

    private static PreparedImage PrepareImageBytesCore(byte[] data, CancellationToken cancellationToken) {
        if (!OfficeImageReader.TryIdentify(data, null, cancellationToken, out OfficeImageInfo sourceInfo)) {
            // Keep the established pass-through contract for JPEG streams whose dimensions are not
            // understood by the managed header reader. The PDF writer embeds JPEG data without
            // decoding it, and layout deliberately falls back to the requested/page box in this case.
            if (LooksLikeJpeg(data)) {
                OfficeImageMetadataSnapshot jpegMetadata = OfficeImageMetadataInspector.Inspect(
                    data,
                    OfficeImageFormat.Jpeg,
                    retainedManagedBytes: 0L,
                    cancellationToken: cancellationToken);
                bool hasJpegIcc = (jpegMetadata.Kinds & OfficeImageMetadataKinds.Icc) != 0;
                if (hasJpegIcc && jpegMetadata.Icc == null) {
                    throw new NotSupportedException(
                        SupportedImageMessage + " The embedded JPEG ICC profile cannot be retained or normalized safely.");
                }
                if (hasJpegIcc) {
                    if (!PdfWriter.TryGetJpegComponentCount(data, cancellationToken, out int jpegComponentCount)) {
                        throw new NotSupportedException(
                            SupportedImageMessage + " The tagged JPEG component count cannot be verified; four-component JPEG data cannot be normalized safely.");
                    }
                    if (jpegComponentCount == 4) {
                        throw new NotSupportedException(
                            SupportedImageMessage + " A four-component JPEG with an embedded ICC profile cannot be normalized safely.");
                    }
                }
                return new PreparedImage(
                    CloneWithCancellation(data, cancellationToken),
                    new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0),
                    OfficeImageFormat.Jpeg,
                    wasTranscoded: false);
            }

            // Metadata identification deliberately rejects PNG dimensions that exceed the
            // shared raster budget. Preserve the PDF writer's more specific validation
            // diagnostic instead of collapsing those payloads into an unknown header.
            if (LooksLikePng(data)) {
                if (PdfWriter.TryGetPngImageData(data, cancellationToken, out PdfWriter.PdfImageStream pngImage, out string? pngReason)) {
                    return new PreparedImage(
                        CloneWithCancellation(data, cancellationToken),
                        new OfficeImageInfo(OfficeImageFormat.Png, pngImage.PixelWidth, pngImage.PixelHeight),
                        OfficeImageFormat.Png,
                        wasTranscoded: false,
                        pngImage);
                }

                string suffix = string.IsNullOrWhiteSpace(pngReason) ? string.Empty : " " + pngReason;
                throw new NotSupportedException(SupportedImageMessage + suffix);
            }

            throw new NotSupportedException(SupportedImageMessage + " The source image header is not recognized.");
        }

        OfficeImageMetadataSnapshot sourceMetadata = OfficeImageMetadataInspector.Inspect(
            data,
            sourceInfo.Format,
            retainedManagedBytes: 0L,
            cancellationToken: cancellationToken);
        bool hasEmbeddedIccProfile = (sourceMetadata.Kinds & OfficeImageMetadataKinds.Icc) != 0;
        if (hasEmbeddedIccProfile && sourceMetadata.Icc == null) {
            throw new NotSupportedException(
                SupportedImageMessage + " The embedded ICC profile cannot be retained or normalized safely.");
        }

        if (sourceInfo.Format == OfficeImageFormat.Jpeg) {
            bool hasComponentCount = PdfWriter.TryGetJpegComponentCount(data, cancellationToken, out int componentCount);
            if (hasEmbeddedIccProfile && !hasComponentCount) {
                throw new NotSupportedException(
                    SupportedImageMessage + " The tagged JPEG component count cannot be verified; four-component JPEG data cannot be normalized safely.");
            }
            if (hasComponentCount && componentCount == 4) {
                if (hasEmbeddedIccProfile) {
                    throw new NotSupportedException(
                        SupportedImageMessage + " A four-component JPEG with an embedded ICC profile cannot be normalized safely.");
                }
                if (!OfficeImagePdfCompatibility.TryValidateTranscodeDimensions(
                        sourceInfo,
                        OfficeImagePdfCompatibility.DefaultMaximumTranscodePixels,
                        out string? jpegTranscodeLimitReason)) {
                    throw new NotSupportedException(SupportedImageMessage + " " + jpegTranscodeLimitReason);
                }
                if (!OfficeImagePngConverter.TryConvertToPng(data, cancellationToken, out byte[] normalizedJpegPng) ||
                    !OfficeImageReader.TryIdentify(normalizedJpegPng, null, cancellationToken, out OfficeImageInfo normalizedJpegInfo)) {
                    throw new NotSupportedException(SupportedImageMessage + " Four-component JPEG data could not be normalized safely for PDF embedding.");
                }
                if (!PdfWriter.TryGetPngImageData(
                        normalizedJpegPng,
                        cancellationToken,
                        out PdfWriter.PdfImageStream normalizedJpegStream,
                        out string? normalizedJpegReason)) {
                    string suffix = string.IsNullOrWhiteSpace(normalizedJpegReason) ? string.Empty : " " + normalizedJpegReason;
                    throw new NotSupportedException(SupportedImageMessage + " Four-component JPEG data could not be normalized safely for PDF embedding." + suffix);
                }
                return new PreparedImage(
                    normalizedJpegPng,
                    normalizedJpegInfo,
                    sourceInfo.Format,
                    wasTranscoded: true,
                    normalizedJpegStream);
            }
            if (componentCount != 0 && componentCount != 1 && componentCount != 3) {
                throw new NotSupportedException(SupportedImageMessage + " JPEG component count is not supported for PDF embedding.");
            }
            byte[] jpegSnapshot = CloneWithCancellation(data, cancellationToken);
            return new PreparedImage(
                jpegSnapshot,
                sourceInfo,
                sourceInfo.Format,
                wasTranscoded: false,
                PdfWriter.CreatePreparedJpegStream(jpegSnapshot, sourceInfo, hasComponentCount ? componentCount : 0));
        }

        if (sourceInfo.Format == OfficeImageFormat.Png) {
            if (PdfWriter.TryGetPngImageData(data, cancellationToken, out PdfWriter.PdfImageStream sourcePngImage, out string? sourcePngReason)) {
                return new PreparedImage(
                    CloneWithCancellation(data, cancellationToken),
                    sourceInfo,
                    sourceInfo.Format,
                    wasTranscoded: false,
                    sourcePngImage);
            }

            string suffix = string.IsNullOrWhiteSpace(sourcePngReason) ? string.Empty : " " + sourcePngReason;
            throw new NotSupportedException(SupportedImageMessage + suffix);
        }

        if (hasEmbeddedIccProfile) {
            throw new NotSupportedException(
                SupportedImageMessage + $" The embedded {sourceInfo.Format} ICC profile cannot be retained through PNG normalization.");
        }

        if (!OfficeImagePdfCompatibility.TryValidateTranscodeDimensions(
                sourceInfo,
                OfficeImagePdfCompatibility.DefaultMaximumTranscodePixels,
                out string? transcodeLimitReason)) {
            throw new NotSupportedException(SupportedImageMessage + " " + transcodeLimitReason);
        }

        if (!OfficeImagePngConverter.TryConvertToPng(data, cancellationToken, out byte[] normalizedPng)) {
            throw new NotSupportedException(
                $"{SupportedImageMessage} Detected {sourceInfo.Format} ({sourceInfo.MimeType}), but it could not be normalized.");
        }

        if (!PdfWriter.TryGetPngImageData(normalizedPng, cancellationToken, out PdfWriter.PdfImageStream normalizedImage, out string? normalizedReason)) {
            string suffix = string.IsNullOrWhiteSpace(normalizedReason) ? string.Empty : " " + normalizedReason;
            throw new NotSupportedException(
                $"{SupportedImageMessage} Detected {sourceInfo.Format} ({sourceInfo.MimeType}), but it could not be normalized.{suffix}");
        }

        OfficeImageInfo normalizedInfo = OfficeImageReader.TryIdentify(
            normalizedPng,
            null,
            cancellationToken,
            out OfficeImageInfo identifiedNormalized)
                ? identifiedNormalized
                : new OfficeImageInfo(
                    OfficeImageFormat.Png,
                    normalizedImage.PixelWidth,
                    normalizedImage.PixelHeight,
                    sourceInfo.DpiX,
                    sourceInfo.DpiY);
        return new PreparedImage(
            normalizedPng,
            normalizedInfo,
            sourceInfo.Format,
            wasTranscoded: true,
            normalizedImage);
    }

    private static bool BytesEqual(byte[] left, byte[] right) {
#if NET8_0_OR_GREATER
        return left.AsSpan().SequenceEqual(right);
#else
        if (left.Length != right.Length) return false;
        for (int i = 0; i < left.Length; i++) {
            if (left[i] != right[i]) return false;
        }
        return true;
#endif
    }

    private static byte[] ComputeImageSourceHash(byte[] data) {
#if NET6_0_OR_GREATER
        return SHA256.HashData(data);
#else
        using (SHA256 sha256 = SHA256.Create()) {
            return sha256.ComputeHash(data);
        }
#endif
    }

    private static byte[] CloneWithCancellation(byte[] source, CancellationToken cancellationToken) {
        var copy = new byte[source.Length];
        const int chunkSize = 64 * 1024;
        for (int offset = 0; offset < source.Length; offset += chunkSize) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = Math.Min(chunkSize, source.Length - offset);
            Buffer.BlockCopy(source, offset, copy, offset, count);
        }
        return copy;
    }

    private static bool LooksLikeJpeg(byte[] data) =>
        data.Length >= 4 &&
        data[0] == 0xFF &&
        data[1] == 0xD8 &&
        data[data.Length - 2] == 0xFF &&
        data[data.Length - 1] == 0xD9;

    private static bool LooksLikePng(byte[] data) =>
        data.Length >= 8 &&
        data[0] == 137 && data[1] == 80 && data[2] == 78 && data[3] == 71 &&
        data[4] == 13 && data[5] == 10 && data[6] == 26 && data[7] == 10;
}
