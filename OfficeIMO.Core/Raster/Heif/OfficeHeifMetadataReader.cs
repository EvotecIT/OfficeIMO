// Adapted from the Evotec ImagePlayground HEIF metadata implementation.
// Copyright (c) 2022 Evotec. MIT license; see Licenses/ImagePlayground-LICENSE.txt.
// OfficeIMO adaptation provides bounded byte/stream APIs, cancellation and output preflight.
using System;
using System.IO;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Reads and replaces HEIF/HEIC item metadata without decoding image pixels.</summary>
/// <remarks>
/// Encoded input and output are bounded to 128 MiB, individual metadata/property payloads to
/// 16 MiB, declared item collections to 4,096 entries, and parser work to 65,536 records.
/// Rewriting replaces or clears an existing single file-backed extent. Item-data-box,
/// external, derived and multiple-extent items are readable where supported but not writable.
/// Each metadata family is inspected independently; unchanged siblings are preserved as bytes.
/// Reads and writes select a unique metadata item associated with the primary image through
/// cdsc references. A sole unassociated metadata item remains supported for older containers;
/// ambiguous or explicitly unrelated items are not selected. Information retains all declared
/// items, while incomplete declared collections are rejected without returning partial results.
/// Protected items and XMP items declaring a content encoding remain discoverable, but their
/// payload reads and writes (including clearing) are unsupported. Stream reads include known
/// caller-owned memory backing in the operation's managed working-set budget.
/// XMP writes also account for the supplied UTF-16 string in that budget.
/// XMP reads reject malformed UTF-8, and writes reject strings containing unpaired UTF-16
/// surrogates. A valid new packet can replace an existing malformed requested payload.
/// </remarks>
public static partial class OfficeHeifMetadataReader {
    /// <summary>Reads container brands, primary image properties, items, locations and references.</summary>
    public static bool TryReadInfo(byte[] data, out OfficeHeifImageInfo? info, CancellationToken cancellationToken = default) {
        OfficeHeifImageInfo? value = null;
        bool success = TryRun(data, cancellationToken, parser => parser.TryReadInfo(data, out value));
        info = success ? value : null;
        return success;
    }

    /// <summary>Reads the existing EXIF item through the shared typed EXIF codec.</summary>
    public static bool TryReadExifProfile(byte[] data, out OfficeImageMetadata? profile, CancellationToken cancellationToken = default) {
        OfficeImageMetadata? value = null;
        bool success = TryRun(data, cancellationToken, parser => parser.TryReadExifProfile(data, out value));
        profile = success ? value : null;
        return success;
    }

    /// <summary>Reads an existing XMP item's UTF-8 packet without parsing its EXIF sibling.</summary>
    public static bool TryReadXmp(byte[] data, out string? xmp, CancellationToken cancellationToken = default) {
        string? value = null;
        bool success = TryRun(data, cancellationToken, parser => parser.TryReadXmp(data, out value));
        xmp = success ? value : null;
        return success;
    }

    /// <summary>Returns whether the container declares an EXIF item, even if its payload is unlocated.</summary>
    public static bool HasExifItem(byte[] data, CancellationToken cancellationToken = default) =>
        TryRun(data, cancellationToken, parser => parser.HasExifItem(data));

    /// <summary>Returns whether the container declares an XMP item, even if its payload is unlocated.</summary>
    public static bool HasXmpItem(byte[] data, CancellationToken cancellationToken = default) =>
        TryRun(data, cancellationToken, parser => parser.HasXmpItem(data));

    /// <summary>Replaces or clears an existing writable EXIF item and returns independently owned container bytes.</summary>
    /// <remarks>A null profile clears the item. Rejection leaves input unchanged and returns null output.</remarks>
    public static bool TryWriteExifProfile(byte[] data, OfficeImageMetadata? profile, out byte[]? output, CancellationToken cancellationToken = default) {
        byte[]? value = null;
        bool success = TryRun(data, cancellationToken, parser => parser.TryWriteExifProfile(data, profile, out value));
        output = success ? value : null;
        return success;
    }

    /// <summary>Replaces or clears an existing writable XMP item and returns independently owned container bytes.</summary>
    /// <remarks>A null packet clears the item. Rejection leaves input unchanged and returns null output.</remarks>
    public static bool TryWriteXmp(byte[] data, string? xmp, out byte[]? output, CancellationToken cancellationToken = default) {
        byte[]? value = null;
        bool success = TryRun(data, cancellationToken, parser => parser.TryWriteXmp(data, xmp, out value));
        output = success ? value : null;
        return success;
    }

    /// <summary>Reads container information from the current stream position, restoring seekable streams and leaving them open.</summary>
    public static bool TryReadInfo(Stream source, out OfficeHeifImageInfo? info, CancellationToken cancellationToken = default) {
        byte[] data = ReadBytes(source, cancellationToken, out long retainedBytes);
        OfficeHeifImageInfo? value = null;
        bool success = TryRun(data, cancellationToken, parser => parser.TryReadInfo(data, out value), retainedBytes);
        info = success ? value : null;
        return success;
    }

    /// <summary>Reads EXIF from the current stream position, restoring seekable streams and leaving them open.</summary>
    public static bool TryReadExifProfile(Stream source, out OfficeImageMetadata? profile, CancellationToken cancellationToken = default) {
        byte[] data = ReadBytes(source, cancellationToken, out long retainedBytes);
        OfficeImageMetadata? value = null;
        bool success = TryRun(data, cancellationToken, parser => parser.TryReadExifProfile(data, out value), retainedBytes);
        profile = success ? value : null;
        return success;
    }

    /// <summary>Reads XMP from the current stream position, restoring seekable streams and leaving them open.</summary>
    public static bool TryReadXmp(Stream source, out string? xmp, CancellationToken cancellationToken = default) {
        byte[] data = ReadBytes(source, cancellationToken, out long retainedBytes);
        string? value = null;
        bool success = TryRun(data, cancellationToken, parser => parser.TryReadXmp(data, out value), retainedBytes);
        xmp = success ? value : null;
        return success;
    }

    /// <summary>Reads bounded container information from a file.</summary>
    public static bool TryReadInfo(string filePath, out OfficeHeifImageInfo? info, CancellationToken cancellationToken = default) =>
        TryReadInfo(ReadFile(filePath, cancellationToken), out info, cancellationToken);

    /// <summary>Reads an existing EXIF item from a bounded file.</summary>
    public static bool TryReadExifProfile(string filePath, out OfficeImageMetadata? profile, CancellationToken cancellationToken = default) =>
        TryReadExifProfile(ReadFile(filePath, cancellationToken), out profile, cancellationToken);

    /// <summary>Reads an existing XMP item from a bounded file.</summary>
    public static bool TryReadXmp(string filePath, out string? xmp, CancellationToken cancellationToken = default) =>
        TryReadXmp(ReadFile(filePath, cancellationToken), out xmp, cancellationToken);

    /// <summary>Returns whether a bounded file declares an EXIF item.</summary>
    public static bool HasExifItem(string filePath, CancellationToken cancellationToken = default) =>
        HasExifItem(ReadFile(filePath, cancellationToken), cancellationToken);

    /// <summary>Returns whether a bounded file declares an XMP item.</summary>
    public static bool HasXmpItem(string filePath, CancellationToken cancellationToken = default) =>
        HasXmpItem(ReadFile(filePath, cancellationToken), cancellationToken);

    /// <summary>Replaces or clears an existing writable EXIF item in a file.</summary>
    /// <remarks>Stages complete output beside the destination and atomically commits it. Rejection, staging failure, and cancellation before commit preserve the destination. Unsupported atomic replacement fails explicitly.</remarks>
    public static bool TryWriteExifProfile(string filePath, string outputPath, OfficeImageMetadata? profile, CancellationToken cancellationToken = default) {
        if (!TryWriteExifProfile(ReadFile(filePath, cancellationToken), profile, out byte[]? output, cancellationToken)) {
            return false;
        }
        WriteFile(outputPath, output!, cancellationToken);
        return true;
    }

    /// <summary>Replaces or clears an existing writable XMP item in a file.</summary>
    /// <remarks>Stages complete output beside the destination and atomically commits it. Rejection, staging failure, and cancellation before commit preserve the destination. Unsupported atomic replacement fails explicitly.</remarks>
    public static bool TryWriteXmp(string filePath, string outputPath, string? xmp, CancellationToken cancellationToken = default) {
        if (!TryWriteXmp(ReadFile(filePath, cancellationToken), xmp, out byte[]? output, cancellationToken)) {
            return false;
        }
        WriteFile(outputPath, output!, cancellationToken);
        return true;
    }

    private static bool TryRun(byte[] data, CancellationToken token, Func<Parser, bool> operation, long additionallyRetainedBytes = 0L) {
        if (data == null) {
            throw new ArgumentNullException(nameof(data));
        }
        token.ThrowIfCancellationRequested();
        if (!OfficeRasterGuards.IsEncodedPayloadWithinLimits(data.Length)) {
            return false;
        }
        try {
            long retainedBytes = checked(data.LongLength + 24L + additionallyRetainedBytes + (additionallyRetainedBytes > 0 ? 24L : 0L));
            bool result = operation(new Parser(retainedBytes, token));
            token.ThrowIfCancellationRequested();
            return result;
        } catch (FormatException) {
            return false;
        } catch (OverflowException) {
            return false;
        } catch (ArgumentException) {
            return false;
        }
    }

    private static byte[] ReadBytes(Stream source, CancellationToken token, out long additionallyRetainedBytes) =>
        OfficeRasterImageDecoder.ReadEncodedBytes(source, out additionallyRetainedBytes,
            new OfficeRasterDecodeOptions { CancellationToken = token });

    private static byte[] ReadFile(string path, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        using var source = File.OpenRead(path);
        return ReadBytes(source, token, out _);
    }

    private static void WriteFile(string path, byte[] data, CancellationToken token) {
        OfficeImageFileWriter.WriteAllBytes(path, data, token);
    }

    private sealed partial class Parser {
        private const string ExifItemType = "Exif";
        private const string XmpMimeType = "application/rdf+xml";
        private readonly CancellationToken _cancellationToken;
        private readonly long _sourceBytes;
        private long _allocatedBytes;
        private int _workRecords;

        internal Parser(long sourceBytes, CancellationToken token) {
            _sourceBytes = sourceBytes;
            _cancellationToken = token;
        }

        private void CheckWork() {
            _cancellationToken.ThrowIfCancellationRequested();
            if (++_workRecords > 65_536) {
                throw new FormatException("HEIF structure exceeds the parser work limit.");
            }
            ReserveBytes(256L);
        }

        private void ReserveBytes(long bytes) {
            _allocatedBytes = checked(_allocatedBytes + bytes);
            if (checked(_sourceBytes + _allocatedBytes) > OfficeRasterGuards.MaximumDecodedBytes) {
                throw new FormatException("HEIF metadata exceeds the managed working-set limit.");
            }
        }

        private void CopyBytes(byte[] source, int sourceOffset, byte[] destination, int destinationOffset, int length) {
            const int chunkSize = 64 * 1024;
            for (int offset = 0; offset < length; offset += chunkSize) {
                _cancellationToken.ThrowIfCancellationRequested();
                Buffer.BlockCopy(source, sourceOffset + offset, destination, destinationOffset + offset,
                    Math.Min(chunkSize, length - offset));
            }
        }
    }
}
