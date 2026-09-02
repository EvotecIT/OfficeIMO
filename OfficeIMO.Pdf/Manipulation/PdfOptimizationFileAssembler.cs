using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Threading;

namespace OfficeIMO.Pdf;

internal static class PdfOptimizationFileAssembler {
    private const int ObjectStreamChunkSize = 100;
    private const string OutputLimitMessage = "The optimized PDF exceeded the configured output limit while it was being serialized.";
    private static readonly byte[] ObjectStreamSeparator = { (byte)'\n' };

    internal static byte[] Assemble(IReadOnlyList<byte[]> bodies, IReadOnlyList<bool> objectStreamEligibility, int catalogId, int infoId, PdfFileVersion fileVersion, PdfOptimizationOptions options, string trailerIdEntry) {
        if (bodies.Count != objectStreamEligibility.Count) throw new ArgumentException("Object body and eligibility counts must match.", nameof(objectStreamEligibility));
        using var output = new MemoryStream();
        using (var boundedOutput = new PdfBoundedWriteStream(
            output,
            options.MaximumOutputBytes,
            OutputLimitMessage)) {
            if (!options.UseObjectStreams && options.XrefFormat == PdfOptimizationXrefFormat.ClassicTable) {
                using var objects = new PdfObjectStore();
                for (int i = 0; i < bodies.Count; i++) {
                    options.CancellationToken.ThrowIfCancellationRequested();
                    objects.AddSegments(
                        PdfEncoding.Latin1GetBytes((i + 1).ToString(CultureInfo.InvariantCulture) + " 0 obj\n"),
                        bodies[i],
                        PdfEncoding.Latin1GetBytes("endobj\n"));
                }
                PdfFileAssembler.Assemble(
                    boundedOutput,
                    objects,
                    catalogId,
                    infoId,
                    fileVersion,
                    trailerIdEntry: trailerIdEntry,
                    cancellationToken: options.CancellationToken);
            } else {
                AssembleXrefStream(boundedOutput, bodies, objectStreamEligibility, catalogId, infoId, fileVersion, options, trailerIdEntry);
            }
        }
        options.CancellationToken.ThrowIfCancellationRequested();
        return output.ToArray();
    }

    private static void AssembleXrefStream(Stream output, IReadOnlyList<byte[]> bodies, IReadOnlyList<bool> eligibility, int catalogId, int infoId, PdfFileVersion fileVersion, PdfOptimizationOptions options, string trailerIdEntry) {
        fileVersion = PdfFileAssembler.RequireAtLeast(fileVersion, PdfFileVersion.Pdf15);
        var packs = BuildObjectStreamPacks(bodies, eligibility, options.UseObjectStreams);
        int baseCount = bodies.Count;
        for (int i = 0; i < packs.Count; i++) packs[i].ObjectNumber = baseCount + i + 1;
        int xrefObjectNumber = baseCount + packs.Count + 1;
        int size = xrefObjectNumber + 1;
        EnsureTemporaryBufferWithinTotalLimit(checked((long)size * 13L), options.MaximumOutputBytes);
        var types = new byte[size]; var field2 = new long[size]; var field3 = new int[size];
        field3[0] = 65535;
        foreach (ObjectStreamPack pack in packs) for (int i = 0; i < pack.ObjectIds.Count; i++) { int id = pack.ObjectIds[i]; types[id] = 2; field2[id] = pack.ObjectNumber; field3[id] = i; }

        byte[] header = PdfEncoding.Latin1GetBytes("%PDF-" + PdfFileAssembler.GetHeaderVersion(fileVersion) + "\n%\u00e2\u00e3\u00cf\u00d3\n"); output.Write(header, 0, header.Length);
        for (int id = 1; id <= baseCount; id++) {
            options.CancellationToken.ThrowIfCancellationRequested();
            if (types[id] == 2) continue;
            types[id] = 1; field2[id] = output.Position;
            WriteIndirectObject(output, id, bodies[id - 1]);
        }
        foreach (ObjectStreamPack pack in packs) {
            options.CancellationToken.ThrowIfCancellationRequested();
            types[pack.ObjectNumber] = 1; field2[pack.ObjectNumber] = output.Position;
            WriteObjectStream(output, pack, bodies, options);
        }
        long xrefOffset = output.Position; types[xrefObjectNumber] = 1; field2[xrefObjectNumber] = xrefOffset;
        EnsureTemporaryBufferWithinRemainingLimit(output, checked((long)types.Length * 13L), options.MaximumOutputBytes);
        int xrefLength = checked(types.Length * 13);
        string xrefDictionary = "<< /Type /XRef /Size " + size.ToString(CultureInfo.InvariantCulture) + " /W [1 8 4] /Root " + PdfSyntaxEscaper.IndirectReference(catalogId) + (infoId > 0 ? " /Info " + PdfSyntaxEscaper.IndirectReference(infoId) : string.Empty) + trailerIdEntry + " /Length " + xrefLength.ToString(CultureInfo.InvariantCulture) + " >>";
        WriteXrefStream(output, xrefObjectNumber, xrefDictionary, types, field2, field3, options.CancellationToken);
        Write(output, PdfEncoding.Latin1GetBytes("startxref\n" + xrefOffset.ToString(CultureInfo.InvariantCulture) + "\n%%EOF\n"));
    }

    private static List<ObjectStreamPack> BuildObjectStreamPacks(IReadOnlyList<byte[]> bodies, IReadOnlyList<bool> eligibility, bool enabled) {
        var packs = new List<ObjectStreamPack>(); if (!enabled) return packs;
        ObjectStreamPack? current = null;
        for (int i = 0; i < bodies.Count; i++) {
            if (!eligibility[i]) continue;
            if (current is null || current.ObjectIds.Count == ObjectStreamChunkSize) { current = new ObjectStreamPack(); packs.Add(current); }
            current.ObjectIds.Add(i + 1);
        }
        return packs;
    }

    private static void WriteObjectStream(
        Stream output,
        ObjectStreamPack pack,
        IReadOnlyList<byte[]> bodies,
        PdfOptimizationOptions options) {
        byte[] header = BuildObjectStreamHeader(pack, bodies, options.CancellationToken);
        long? maximumCompressedBytes = GetRemainingOutputBytes(output, options.MaximumOutputBytes);
        using MemoryStream compressed = CompressObjectStreamContent(
            pack,
            bodies,
            header,
            maximumCompressedBytes,
            options.CancellationToken);
        int compressedLength = checked((int)compressed.Length);
        string dictionary = "<< /Type /ObjStm /N " + pack.ObjectIds.Count.ToString(CultureInfo.InvariantCulture) +
            " /First " + header.Length.ToString(CultureInfo.InvariantCulture) +
            " /Filter /FlateDecode /Length " + compressedLength.ToString(CultureInfo.InvariantCulture) + " >>";
        WriteStreamObject(output, pack.ObjectNumber, dictionary, compressed.GetBuffer(), compressedLength);
    }

    private static byte[] BuildObjectStreamHeader(
        ObjectStreamPack pack,
        IReadOnlyList<byte[]> bodies,
        CancellationToken cancellationToken) {
        var header = new StringBuilder();
        long offset = 0L;
        for (int i = 0; i < pack.ObjectIds.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            int id = pack.ObjectIds[i];
            header.Append(id.ToString(CultureInfo.InvariantCulture))
                .Append(' ')
                .Append(offset.ToString(CultureInfo.InvariantCulture))
                .Append(' ');
            offset = checked(offset + bodies[id - 1].LongLength + 1L);
        }
        header.Append('\n');
        return PdfEncoding.Latin1GetBytes(header.ToString());
    }

    private static MemoryStream CompressObjectStreamContent(
        ObjectStreamPack pack,
        IReadOnlyList<byte[]> bodies,
        byte[] header,
        long? maximumCompressedBytes,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        var compressed = new MemoryStream();
        try {
            using (var bounded = new PdfBoundedWriteStream(compressed, maximumCompressedBytes, OutputLimitMessage)) {
                bounded.WriteByte(0x78);
                bounded.WriteByte(0x9C);
                uint adlerA = 1U;
                uint adlerB = 0U;
                using (var deflate = new DeflateStream(bounded, CompressionLevel.Optimal, leaveOpen: true)) {
                    WriteCompressedSegment(deflate, header, cancellationToken, ref adlerA, ref adlerB);
                    for (int i = 0; i < pack.ObjectIds.Count; i++) {
                        cancellationToken.ThrowIfCancellationRequested();
                        WriteCompressedSegment(
                            deflate,
                            bodies[pack.ObjectIds[i] - 1],
                            cancellationToken,
                            ref adlerA,
                            ref adlerB);
                        WriteCompressedSegment(
                            deflate,
                            ObjectStreamSeparator,
                            cancellationToken,
                            ref adlerA,
                            ref adlerB);
                    }
                }
                uint adler = (adlerB << 16) | adlerA;
                bounded.WriteByte((byte)(adler >> 24));
                bounded.WriteByte((byte)(adler >> 16));
                bounded.WriteByte((byte)(adler >> 8));
                bounded.WriteByte((byte)adler);
            }
            cancellationToken.ThrowIfCancellationRequested();
            return compressed;
        } catch {
            compressed.Dispose();
            throw;
        }
    }

    private static void WriteCompressedSegment(
        DeflateStream destination,
        byte[] bytes,
        CancellationToken cancellationToken,
        ref uint adlerA,
        ref uint adlerB) {
        const int chunkSize = 64 * 1024;
        const uint modulus = 65521U;
        for (int offset = 0; offset < bytes.Length; offset += chunkSize) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = Math.Min(chunkSize, bytes.Length - offset);
            destination.Write(bytes, offset, count);
            int end = offset + count;
            for (int index = offset; index < end; index++) {
                adlerA = (adlerA + bytes[index]) % modulus;
                adlerB = (adlerB + adlerA) % modulus;
            }
        }
    }

    private static void WriteStreamObject(
        Stream output,
        int objectNumber,
        string dictionary,
        byte[] content,
        int? contentLength = null) {
        byte[][] segments = PdfObjectBytes.CreateStreamObjectSegments(objectNumber, dictionary, content);
        Write(output, segments[0]);
        int length = contentLength ?? content.Length;
        output.Write(content, 0, length);
        Write(output, segments[2]);
    }

    private static void WriteIndirectObject(Stream output, int objectNumber, byte[] body) {
        Write(output, PdfEncoding.Latin1GetBytes(objectNumber.ToString(CultureInfo.InvariantCulture) + " 0 obj\n"));
        Write(output, body);
        Write(output, PdfEncoding.Latin1GetBytes("endobj\n"));
    }

    private static void WriteXrefStream(
        Stream output,
        int objectNumber,
        string dictionary,
        byte[] types,
        long[] field2,
        int[] field3,
        CancellationToken cancellationToken) {
        byte[][] segments = PdfObjectBytes.CreateStreamObjectSegments(objectNumber, dictionary, Array.Empty<byte>());
        Write(output, segments[0]);
        var entry = new byte[13];
        for (int index = 0; index < types.Length; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            entry[0] = types[index];
            WriteBigEndian(entry, 1, field2[index], 8);
            WriteBigEndian(entry, 9, field3[index], 4);
            Write(output, entry);
        }
        Write(output, segments[2]);
    }

    private static long? GetRemainingOutputBytes(Stream output, long? maximumOutputBytes) {
        if (!maximumOutputBytes.HasValue) return null;
        long remaining = maximumOutputBytes.Value - output.Position;
        if (remaining < 1L) throw PdfOutputLimitErrors.Create(OutputLimitMessage);
        return remaining;
    }

    private static void EnsureTemporaryBufferWithinTotalLimit(long length, long? maximumOutputBytes) {
        if (length > int.MaxValue) {
            throw new InvalidOperationException("The optimized PDF exceeds the supported in-memory result size.");
        }
        if (maximumOutputBytes.HasValue && length > maximumOutputBytes.Value) {
            throw PdfOutputLimitErrors.Create(OutputLimitMessage);
        }
    }

    private static void EnsureTemporaryBufferWithinRemainingLimit(
        Stream output,
        long length,
        long? maximumOutputBytes) {
        EnsureTemporaryBufferWithinTotalLimit(length, maximumOutputBytes);
        if (maximumOutputBytes.HasValue && length > maximumOutputBytes.Value - output.Position) {
            throw PdfOutputLimitErrors.Create(OutputLimitMessage);
        }
    }

    private static void WriteBigEndian(byte[] destination, int offset, long value, int length) { for (int i = length - 1; i >= 0; i--) { destination[offset + i] = (byte)(value & 0xFF); value >>= 8; } }
    private static void Write(Stream output, byte[] bytes) => output.Write(bytes, 0, bytes.Length);
    internal static byte[] CompressFlate(byte[] data, CancellationToken cancellationToken) {
        const int chunkSize = 64 * 1024;
        cancellationToken.ThrowIfCancellationRequested();
        using var output = new MemoryStream();
        output.WriteByte(0x78);
        output.WriteByte(0x9C);
        using (var deflate = new DeflateStream(output, CompressionLevel.Optimal, true)) {
            for (int offset = 0; offset < data.Length; offset += chunkSize) {
                cancellationToken.ThrowIfCancellationRequested();
                deflate.Write(data, offset, Math.Min(chunkSize, data.Length - offset));
            }
        }
        uint adler = Adler32(data, cancellationToken);
        output.WriteByte((byte)(adler >> 24));
        output.WriteByte((byte)(adler >> 16));
        output.WriteByte((byte)(adler >> 8));
        output.WriteByte((byte)adler);
        return output.ToArray();
    }

    private static uint Adler32(byte[] data, CancellationToken cancellationToken) {
        const int chunkSize = 64 * 1024;
        const uint mod = 65521;
        uint a = 1, b = 0;
        for (int i = 0; i < data.Length; i++) {
            if (i % chunkSize == 0) cancellationToken.ThrowIfCancellationRequested();
            a = (a + data[i]) % mod;
            b = (b + a) % mod;
        }
        return (b << 16) | a;
    }

    private sealed class ObjectStreamPack { internal int ObjectNumber { get; set; } internal List<int> ObjectIds { get; } = new List<int>(); }
}
