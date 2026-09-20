using System;
#if NET8_0_OR_GREATER
using System.Buffers;
#endif
using System.IO;
using System.IO.Compression;
using System.Threading;

namespace OfficeIMO.Pdf.Filters;

internal static class FlateDecoder {
    private const int CopyBufferSize = 81920;
#if !NET8_0_OR_GREATER
    [ThreadStatic]
    private static byte[]? _validationBuffer;
#endif

    public static byte[] Decode(byte[] data) {
        // Try zlib (RFC1950) first when available in this target
#if NET6_0_OR_GREATER
        if (TryZlib(data, maxOutputBytes: null, out var result, out _)) return result!;
#endif
        // Try raw Deflate
        if (TryInflate(data, maxOutputBytes: null, out var result2, out _)) return result2!;
        // Try skip zlib header (2 bytes) with raw Deflate
        if (data.Length > 2 && IsLikelyZlib(data)) {
            var sliced = new byte[data.Length - 2];
            Buffer.BlockCopy(data, 2, sliced, 0, sliced.Length);
            if (TryInflate(sliced, maxOutputBytes: null, out var result3, out _)) return result3!;
        }
        // Fallback to original
        return data;
    }

    public static bool TryDecode(byte[] data, int maxOutputBytes, out byte[] output) {
        return TryDecode(data, maxOutputBytes, out output, out _);
    }

    public static bool TryDecode(
        byte[] data,
        int maxOutputBytes,
        out byte[] output,
        out bool limitExceeded,
        CancellationToken cancellationToken = default) {
        return TryDecode(data, 0, data.Length, maxOutputBytes, out output, out limitExceeded, cancellationToken);
    }

    internal static bool TryDecode(
        byte[] data,
        int offset,
        int length,
        int maxOutputBytes,
        out byte[] output,
        out bool limitExceeded,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        limitExceeded = false;
        if (offset < 0 || length < 0 || offset > data.Length - length) {
            output = Array.Empty<byte>();
            return false;
        }
        if (maxOutputBytes < 0) {
            output = Array.Empty<byte>();
            return false;
        }

#if NET6_0_OR_GREATER
        if (TryZlib(data, offset, length, maxOutputBytes, out var result, out bool zlibLimitExceeded, cancellationToken)) {
            output = result!;
            return true;
        }

        if (zlibLimitExceeded) {
            limitExceeded = true;
            output = Array.Empty<byte>();
            return false;
        }
#endif

        if (TryInflate(data, offset, length, maxOutputBytes, out var result2, out bool inflateLimitExceeded, cancellationToken)) {
            output = result2!;
            return true;
        }

        if (inflateLimitExceeded) {
            limitExceeded = true;
            output = Array.Empty<byte>();
            return false;
        }

        if (length > 2 && IsLikelyZlib(data, offset, length)) {
            if (TryInflate(data, offset + 2, length - 2, maxOutputBytes, out var result3, out bool slicedLimitExceeded, cancellationToken)) {
                output = result3!;
                return true;
            }

            if (slicedLimitExceeded) {
                limitExceeded = true;
                output = Array.Empty<byte>();
                return false;
            }
        }

        output = Array.Empty<byte>();
        return false;
    }

    internal static bool TryValidate(
        byte[] data,
        int offset,
        int length,
        int maxOutputBytes,
        out int decodedLength,
        out bool limitExceeded,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        decodedLength = 0;
        limitExceeded = false;
        if (offset < 0 || length < 0 || offset > data.Length - length || maxOutputBytes < 0) {
            return false;
        }

#if NET6_0_OR_GREATER
        if (TryValidateZlib(data, offset, length, maxOutputBytes, out decodedLength, out bool zlibLimitExceeded, cancellationToken)) {
            return true;
        }
        if (zlibLimitExceeded) {
            limitExceeded = true;
            return false;
        }
#endif

        if (TryValidateInflate(data, offset, length, maxOutputBytes, out decodedLength, out bool inflateLimitExceeded, cancellationToken)) {
            return true;
        }
        if (inflateLimitExceeded) {
            limitExceeded = true;
            return false;
        }

        if (length > 2 && IsLikelyZlib(data, offset, length)) {
            if (TryValidateInflate(
                    data,
                    offset + 2,
                    length - 2,
                    maxOutputBytes,
                    out decodedLength,
                    out bool slicedLimitExceeded,
                    cancellationToken)) {
                return true;
            }
            limitExceeded = slicedLimitExceeded;
        }
        return false;
    }

    private static bool TryInflate(
        byte[] input,
        int? maxOutputBytes,
        out byte[]? output,
        out bool limitExceeded,
        CancellationToken cancellationToken = default) =>
        TryInflate(input, 0, input.Length, maxOutputBytes, out output, out limitExceeded, cancellationToken);

    private static bool TryInflate(
        byte[] input,
        int offset,
        int length,
        int? maxOutputBytes,
        out byte[]? output,
        out bool limitExceeded,
        CancellationToken cancellationToken = default) {
        try {
            using var msIn = new MemoryStream(input, offset, length, writable: false, publiclyVisible: false);
            using var ds = new DeflateStream(msIn, CompressionMode.Decompress, leaveOpen: true);
            return TryCopyToByteArray(ds, maxOutputBytes, out output, out limitExceeded, cancellationToken);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            throw;
        } catch {
            output = null;
            limitExceeded = false;
            return false;
        }
    }

    private static bool TryValidateInflate(
        byte[] input,
        int offset,
        int length,
        int maxOutputBytes,
        out int decodedLength,
        out bool limitExceeded,
        CancellationToken cancellationToken) {
        try {
            using var inputStream = new MemoryStream(input, offset, length, writable: false, publiclyVisible: false);
            using var decoder = new DeflateStream(inputStream, CompressionMode.Decompress, leaveOpen: true);
            return TryCountDecodedBytes(decoder, maxOutputBytes, out decodedLength, out limitExceeded, cancellationToken);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            throw;
        } catch {
            decodedLength = 0;
            limitExceeded = false;
            return false;
        }
    }

#if NET6_0_OR_GREATER
    private static bool TryZlib(
        byte[] input,
        int? maxOutputBytes,
        out byte[]? output,
        out bool limitExceeded,
        CancellationToken cancellationToken = default) =>
        TryZlib(input, 0, input.Length, maxOutputBytes, out output, out limitExceeded, cancellationToken);

    private static bool TryZlib(
        byte[] input,
        int offset,
        int length,
        int? maxOutputBytes,
        out byte[]? output,
        out bool limitExceeded,
        CancellationToken cancellationToken = default) {
        try {
            using var msIn = new MemoryStream(input, offset, length, writable: false, publiclyVisible: false);
            using var zs = new ZLibStream(msIn, CompressionMode.Decompress, leaveOpen: true);
            return TryCopyToByteArray(zs, maxOutputBytes, out output, out limitExceeded, cancellationToken);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            throw;
        } catch {
            output = null;
            limitExceeded = false;
            return false;
        }
    }

    private static bool TryValidateZlib(
        byte[] input,
        int offset,
        int length,
        int maxOutputBytes,
        out int decodedLength,
        out bool limitExceeded,
        CancellationToken cancellationToken) {
        try {
            using var inputStream = new MemoryStream(input, offset, length, writable: false, publiclyVisible: false);
            using var decoder = new ZLibStream(inputStream, CompressionMode.Decompress, leaveOpen: true);
            return TryCountDecodedBytes(decoder, maxOutputBytes, out decodedLength, out limitExceeded, cancellationToken);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            throw;
        } catch {
            decodedLength = 0;
            limitExceeded = false;
            return false;
        }
    }
#endif

    private static bool TryCountDecodedBytes(
        Stream source,
        int maxOutputBytes,
        out int decodedLength,
        out bool limitExceeded,
        CancellationToken cancellationToken) {
        decodedLength = 0;
        limitExceeded = false;
#if NET8_0_OR_GREATER
        byte[] buffer = ArrayPool<byte>.Shared.Rent(CopyBufferSize);
#else
        byte[] buffer = _validationBuffer ??= new byte[CopyBufferSize];
#endif
        try {
            while (true) {
                cancellationToken.ThrowIfCancellationRequested();
                int read = source.Read(buffer, 0, buffer.Length);
                if (read <= 0) return true;
                if (decodedLength > maxOutputBytes - read) {
                    decodedLength = 0;
                    limitExceeded = true;
                    return false;
                }
                decodedLength += read;
            }
        } finally {
#if NET8_0_OR_GREATER
            ArrayPool<byte>.Shared.Return(buffer);
#endif
        }
    }

    private static bool TryCopyToByteArray(
        Stream source,
        int? maxOutputBytes,
        out byte[]? output,
        out bool limitExceeded,
        CancellationToken cancellationToken = default) {
#if NET8_0_OR_GREATER
        return TryCopyToPooledByteArray(source, maxOutputBytes, out output, out limitExceeded, cancellationToken);
#else
        limitExceeded = false;
        using var msOut = new MemoryStream();
        var buffer = new byte[CopyBufferSize];
        int read;
        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            read = source.Read(buffer, 0, buffer.Length);
            if (read <= 0) break;
            if (maxOutputBytes.HasValue && msOut.Length + read > maxOutputBytes.Value) {
                output = null;
                limitExceeded = true;
                return false;
            }

            msOut.Write(buffer, 0, read);
        }

        output = msOut.ToArray();
        return true;
#endif
    }

#if NET8_0_OR_GREATER
    private static bool TryCopyToPooledByteArray(
        Stream source,
        int? maxOutputBytes,
        out byte[]? output,
        out bool limitExceeded,
        CancellationToken cancellationToken) {
        limitExceeded = false;
        byte[] readBuffer = ArrayPool<byte>.Shared.Rent(CopyBufferSize);
        byte[]? accumulated = null;
        int accumulatedLength = 0;
        try {
            while (true) {
                cancellationToken.ThrowIfCancellationRequested();
                int read = source.Read(readBuffer, 0, CopyBufferSize);
                if (read <= 0) break;
                long requiredLength = (long)accumulatedLength + read;
                if (requiredLength > int.MaxValue ||
                    maxOutputBytes.HasValue && requiredLength > maxOutputBytes.Value) {
                    output = null;
                    limitExceeded = maxOutputBytes.HasValue && requiredLength > maxOutputBytes.Value;
                    return false;
                }

                EnsurePooledCapacity(ref accumulated, (int)requiredLength, accumulatedLength);
                Buffer.BlockCopy(readBuffer, 0, accumulated!, accumulatedLength, read);
                accumulatedLength += read;
            }

            if (accumulatedLength == 0) {
                output = Array.Empty<byte>();
                return true;
            }

            output = new byte[accumulatedLength];
            Buffer.BlockCopy(accumulated!, 0, output, 0, accumulatedLength);
            return true;
        } finally {
            ArrayPool<byte>.Shared.Return(readBuffer);
            if (accumulated != null) ArrayPool<byte>.Shared.Return(accumulated);
        }
    }

    private static void EnsurePooledCapacity(ref byte[]? buffer, int requiredLength, int dataLength) {
        if (buffer != null && buffer.Length >= requiredLength) return;
        int requestedLength = buffer == null
            ? Math.Max(CopyBufferSize, requiredLength)
            : Math.Max(requiredLength, buffer.Length <= int.MaxValue / 2 ? buffer.Length * 2 : requiredLength);
        byte[] expanded = ArrayPool<byte>.Shared.Rent(requestedLength);
        if (buffer != null) {
            if (dataLength > 0) Buffer.BlockCopy(buffer, 0, expanded, 0, dataLength);
            ArrayPool<byte>.Shared.Return(buffer);
        }
        buffer = expanded;
    }
#endif

    private static bool IsLikelyZlib(byte[] d) {
        return IsLikelyZlib(d, 0, d.Length);
    }

    private static bool IsLikelyZlib(byte[] d, int offset, int length) {
        // RFC1950: first byte CMF low 4 bits = 8 for deflate; checksum of first two bytes mod 31 == 0
        if (length < 2) return false;
        bool deflate = (d[offset] & 0x0F) == 8;
        int cmfcm = (d[offset] << 8) + d[offset + 1];
        bool mod = (cmfcm % 31) == 0;
        return deflate && mod;
    }
}
