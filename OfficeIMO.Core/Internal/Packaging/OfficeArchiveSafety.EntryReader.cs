using System;
using System.IO;
using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Core.Internal;

internal static partial class OfficeArchiveSafety {
    /// <summary>Opens a validated compressed range without ZipArchive's declared-size output clamp.
    /// The caller must enforce decoded length, cancellation and CRC on the returned stream.</summary>
    internal static Stream OpenEntryPayload(byte[] bytes, ZipArchiveEntry entry,
        Provenance.OfficeProvenanceZip.OfficeProvenanceZipEntryMetadata metadata) => OpenEntryPayload(bytes, entry, metadata, CancellationToken.None);

    internal static Stream OpenEntryPayload(byte[] bytes, ZipArchiveEntry entry,
        Provenance.OfficeProvenanceZip.OfficeProvenanceZipEntryMetadata metadata, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        long compressedLength = entry.CompressedLength;
        int start = metadata.PayloadOffset;
        if (compressedLength < 0 || compressedLength > int.MaxValue || start < 0
            || start > metadata.PayloadUpperBound || metadata.PayloadUpperBound > bytes.Length
            || compressedLength > metadata.PayloadUpperBound - start) {
            throw new InvalidDataException("ZIP compressed payload is outside its package bounds.");
        }
        ushort localFlags = Provenance.OfficeProvenanceBinary.ReadUInt16(bytes, metadata.LocalHeaderOffset + 6, littleEndian: true);
        ushort localMethod = Provenance.OfficeProvenanceBinary.ReadUInt16(bytes, metadata.LocalHeaderOffset + 8, littleEndian: true);
        if (localFlags != metadata.GeneralPurposeFlags || localMethod != metadata.CompressionMethod
            || (localFlags & (0x0001 | 0x0020 | 0x0040 | 0x2000)) != 0 || localMethod is not (0 or 8)) {
            throw new InvalidDataException("ZIP entry uses conflicting, encrypted or unsupported compression metadata.");
        }
        using (var validation = new MemoryStream(bytes, writable: false)) {
            var local = new byte[30];
            ReadExactlyAt(validation, metadata.LocalHeaderOffset, local, local.Length);
            if (!TryFindZip64Extra(metadata.CentralExtraField, 0, metadata.CentralExtraField.Length, out int zip64Offset, out _))
                throw new InvalidDataException("ZIP central extra fields are malformed or repeat ZIP64 metadata.");
            ValidateLocalIntegrity(validation, local, metadata.LocalHeaderOffset, start, metadata.PayloadUpperBound,
                metadata.Checksum, compressedLength, entry.Length, zip64Offset >= 0);
        }
        if (localMethod == 8 && !OfficeDeflateStreamValidator.TryValidateExact(bytes, start, checked((int)compressedLength), entry.Length, cancellationToken))
            throw new InvalidDataException("ZIP member has invalid or incomplete Deflate framing.");
        var compressed = new MemoryStream(bytes, start, checked((int)compressedLength), writable: false);
        return localMethod == 0 ? compressed : new DeflateStream(compressed, CompressionMode.Decompress);
    }

    /// <summary>Checks a fully decoded snapshot against its ZIP CRC. The reader must also prove
    /// actual decoded EOF; CRC detects corruption and does not authenticate input.</summary>
    internal static void ValidateEntryChecksum(byte[] bytes, uint expectedChecksum,
        CancellationToken cancellationToken = default) {
        uint checksum = Drawing.OfficePngCrc32.Begin();
        for (int offset = 0; offset < bytes.Length; offset += Math.Min(81920, bytes.Length - offset)) {
            cancellationToken.ThrowIfCancellationRequested();
            checksum = Drawing.OfficePngCrc32.Append(checksum, bytes, offset, Math.Min(81920, bytes.Length - offset));
        }
        cancellationToken.ThrowIfCancellationRequested();
        if (Drawing.OfficePngCrc32.Complete(checksum) != expectedChecksum) {
            throw new InvalidDataException("The decoded archive entry does not match its ZIP checksum.");
        }
    }

    /// <summary>Reads a declared ZIP payload with bounded allocation and cooperative cancellation.</summary>
    internal static async Task<byte[]> ReadEntryBytesAsync(Stream source, long declaredLength,
        long maximumLength, CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        cancellationToken.ThrowIfCancellationRequested();
        if (declaredLength < 0 || declaredLength > maximumLength
            || maximumLength < 0 || declaredLength > int.MaxValue) {
            throw new InvalidDataException("The archive entry has an invalid declared length.");
        }

        var bytes = new byte[checked((int)declaredLength)];
        int offset = 0;
        while (offset < bytes.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            int requested = Math.Min(81920, bytes.Length - offset);
            int read = await source.ReadAsync(bytes, offset, requested, cancellationToken).ConfigureAwait(false);
            if (read <= 0 || read > requested) {
                throw new InvalidDataException("The archive entry is shorter than its declared length.");
            }
            offset += read;
        }
        cancellationToken.ThrowIfCancellationRequested();
        var trailing = new byte[1];
        if (await source.ReadAsync(trailing, 0, 1, cancellationToken).ConfigureAwait(false) != 0) {
            throw new InvalidDataException("The archive entry exceeds its declared expansion length.");
        }
        return bytes;
    }
}
