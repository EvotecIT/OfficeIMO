using System.Globalization;
using System.Security.Cryptography;

namespace OfficeIMO.Pdf;

internal static class PdfFileAssembler {
    internal static byte[] Assemble(
        IReadOnlyList<byte[]> objects,
        int catalogId,
        int infoId,
        PdfFileVersion fileVersion = PdfFileVersion.Pdf14,
        PdfStandardEncryptionOptions? encryption = null,
        long objectMemoryLimitBytes = PdfObjectStore.DefaultMemoryLimitBytes,
        string? trailerIdEntry = null) {
        using var stream = new MemoryStream();
        Assemble(stream, objects, catalogId, infoId, fileVersion, encryption, objectMemoryLimitBytes, trailerIdEntry);
        return stream.ToArray();
    }

    internal static long Assemble(
        Stream destination,
        IReadOnlyList<byte[]> objects,
        int catalogId,
        int infoId,
        PdfFileVersion fileVersion = PdfFileVersion.Pdf14,
        PdfStandardEncryptionOptions? encryption = null,
        long objectMemoryLimitBytes = PdfObjectStore.DefaultMemoryLimitBytes,
        string? trailerIdEntry = null) =>
        AssembleWithEvidence(
            destination,
            objects,
            catalogId,
            infoId,
            fileVersion,
            encryption,
            objectMemoryLimitBytes,
            trailerIdEntry,
            out _);

    internal static byte[] AssembleWithEvidence(
        IReadOnlyList<byte[]> objects,
        int catalogId,
        int infoId,
        PdfFileVersion fileVersion,
        PdfStandardEncryptionOptions? encryption,
        long objectMemoryLimitBytes,
        out PdfFileAssemblyBufferEvidence bufferEvidence) {
        using var stream = new MemoryStream();
        AssembleWithEvidence(
            stream,
            objects,
            catalogId,
            infoId,
            fileVersion,
            encryption,
            objectMemoryLimitBytes,
            trailerIdEntry: null,
            out bufferEvidence);
        return stream.ToArray();
    }

    internal static long AssembleWithEvidence(
        Stream destination,
        IReadOnlyList<byte[]> objects,
        int catalogId,
        int infoId,
        PdfFileVersion fileVersion,
        PdfStandardEncryptionOptions? encryption,
        long objectMemoryLimitBytes,
        string? trailerIdEntry,
        out PdfFileAssemblyBufferEvidence bufferEvidence) {
        Guard.FileVersion(fileVersion, nameof(fileVersion));
        Guard.NotNull(destination, nameof(destination));
        Guard.NotNull(objects, nameof(objects));
        if (!destination.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(destination));
        if (objectMemoryLimitBytes < 0L) throw new ArgumentOutOfRangeException(nameof(objectMemoryLimitBytes), objectMemoryLimitBytes, "PDF object-buffer memory limit cannot be negative.");

        long sourceRetainedBytes = GetRetainedMemoryBytes(objects);
        long sourcePeakRetainedBytes = GetPeakRetainedMemoryBytes(objects);
        bool sourceSpilled = objects is PdfObjectStore sourceStore && sourceStore.IsSpilled;
        using PdfEncryptionAssembly? encryptionAssembly = encryption == null
            ? null
            : PdfStandardSecurityWriter.Encrypt(objects, encryption, objectMemoryLimitBytes);
        if (encryptionAssembly != null) {
            fileVersion = RequireAtLeast(fileVersion, GetMinimumEncryptionVersion(encryption!.Algorithm));
            objects = encryptionAssembly.Objects;
        }

        long assemblyPeakRetainedBytes = encryptionAssembly == null
            ? sourcePeakRetainedBytes
            : AddWithoutOverflow(sourceRetainedBytes, GetPeakRetainedMemoryBytes(objects));
        bool assemblySpilled = sourceSpilled || objects is PdfObjectStore assemblyStore && assemblyStore.IsSpilled;
        bufferEvidence = new PdfFileAssemblyBufferEvidence(assemblyPeakRetainedBytes, assemblySpilled);

        byte[] header = PdfEncoding.Latin1GetBytes("%PDF-" + GetHeaderVersion(fileVersion) + "\n%\u00e2\u00e3\u00cf\u00d3\n");
        using HashAlgorithm? fileIdHash = encryptionAssembly == null ? SHA256.Create() : null;
        fileIdHash?.TransformBlock(header, 0, header.Length, header, 0);
        destination.Write(header, 0, header.Length);
        long written = header.LongLength;

        var offsets = new List<long> { 0L };
        for (int i = 0; i < objects.Count; i++) {
            offsets.Add(written);
            if (objects is PdfObjectStore objectStore) {
                objectStore.CopyTo(i, destination, fileIdHash);
                written += objectStore.GetLength(i);
            } else {
                byte[] obj = objects[i];
                fileIdHash?.TransformBlock(obj, 0, obj.Length, obj, 0);
                destination.Write(obj, 0, obj.Length);
                written += obj.LongLength;
            }
        }

        byte[] fileId = encryptionAssembly?.FileId ?? FinalizeFileId(fileIdHash!);

        long xrefPos = written;
        var trailer = new StringBuilder();
        trailer.Append("xref\n");
        trailer.Append("0 ").Append((objects.Count + 1).ToString(CultureInfo.InvariantCulture)).Append('\n');
        trailer.Append("0000000000 65535 f \n");
        for (int i = 1; i <= objects.Count; i++) {
            trailer.Append(offsets[i].ToString("0000000000", CultureInfo.InvariantCulture)).Append(" 00000 n \n");
        }

        trailer.Append("trailer\n");
        trailer.Append("<< /Size ").Append((objects.Count + 1).ToString(CultureInfo.InvariantCulture))
            .Append(" /Root ").Append(PdfSyntaxEscaper.IndirectReference(catalogId))
            .Append(infoId > 0 ? " /Info " + PdfSyntaxEscaper.IndirectReference(infoId) : string.Empty)
            .Append(BuildTrailerEntries(encryptionAssembly, fileId, trailerIdEntry)).Append(" >>\n");
        trailer.Append("startxref\n").Append(xrefPos.ToString(CultureInfo.InvariantCulture)).Append("\n%%EOF\n");
        byte[] trailerBytes = Encoding.ASCII.GetBytes(trailer.ToString());
        destination.Write(trailerBytes, 0, trailerBytes.Length);
        return written + trailerBytes.LongLength;
    }

    private static long GetRetainedMemoryBytes(IReadOnlyList<byte[]> objects) {
        if (objects is PdfObjectStore store) return store.RetainedMemoryBytes;
        long total = 0L;
        for (int index = 0; index < objects.Count; index++) total = AddWithoutOverflow(total, objects[index].LongLength);
        return total;
    }

    private static long GetPeakRetainedMemoryBytes(IReadOnlyList<byte[]> objects) =>
        objects is PdfObjectStore store ? store.PeakRetainedMemoryBytes : GetRetainedMemoryBytes(objects);

    private static long AddWithoutOverflow(long left, long right) =>
        left > long.MaxValue - right ? long.MaxValue : left + right;

    private static string BuildTrailerEntries(PdfEncryptionAssembly? encryptionAssembly, byte[] fileId, string? trailerIdEntry) {
        if (encryptionAssembly == null && !string.IsNullOrWhiteSpace(trailerIdEntry)) {
            return trailerIdEntry!;
        }

        string id = PdfSyntaxEscaper.HexString(fileId);
        string encryptionEntry = encryptionAssembly == null
            ? string.Empty
            : " /Encrypt " + PdfSyntaxEscaper.IndirectReference(encryptionAssembly.EncryptionObjectNumber);
        return encryptionEntry + " /ID [" + id + " " + id + "]";
    }

    private static byte[] FinalizeFileId(HashAlgorithm hash) {
        hash.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
        byte[] fullHash = hash.Hash ?? throw new InvalidOperationException("Unable to calculate the PDF trailer file identifier.");
        var fileId = new byte[16];
        Buffer.BlockCopy(fullHash, 0, fileId, 0, fileId.Length);
        return fileId;
    }

    private static PdfFileVersion GetMinimumEncryptionVersion(PdfStandardEncryptionAlgorithm algorithm) {
        switch (algorithm) {
            case PdfStandardEncryptionAlgorithm.Aes256:
                return PdfFileVersion.Pdf20;
            case PdfStandardEncryptionAlgorithm.Aes128:
                return PdfFileVersion.Pdf16;
            case PdfStandardEncryptionAlgorithm.LegacyRc4:
                return PdfFileVersion.Pdf14;
            default:
                throw new ArgumentOutOfRangeException(nameof(algorithm), algorithm, "Unsupported PDF Standard encryption algorithm.");
        }
    }

    internal static string GetHeaderVersion(PdfFileVersion fileVersion) {
        switch (fileVersion) {
            case PdfFileVersion.Pdf14:
                return "1.4";
            case PdfFileVersion.Pdf15:
                return "1.5";
            case PdfFileVersion.Pdf16:
                return "1.6";
            case PdfFileVersion.Pdf17:
                return "1.7";
            case PdfFileVersion.Pdf20:
                return "2.0";
            default:
                Guard.FileVersion(fileVersion, nameof(fileVersion));
                return "1.4";
        }
    }

    internal static PdfFileVersion RequireAtLeast(PdfFileVersion fileVersion, PdfFileVersion minimumVersion) {
        Guard.FileVersion(fileVersion, nameof(fileVersion));
        Guard.FileVersion(minimumVersion, nameof(minimumVersion));
        return fileVersion < minimumVersion ? minimumVersion : fileVersion;
    }

    internal static PdfFileVersion ParseHeaderVersionOrDefault(string? headerVersion) {
        switch (headerVersion) {
            case "1.5":
                return PdfFileVersion.Pdf15;
            case "1.6":
                return PdfFileVersion.Pdf16;
            case "1.7":
                return PdfFileVersion.Pdf17;
            case "2.0":
                return PdfFileVersion.Pdf20;
            default:
                return PdfFileVersion.Pdf14;
        }
    }
}

internal readonly struct PdfFileAssemblyBufferEvidence {
    internal PdfFileAssemblyBufferEvidence(long peakRetainedObjectBytes, bool objectBufferSpilled) {
        PeakRetainedObjectBytes = peakRetainedObjectBytes;
        ObjectBufferSpilled = objectBufferSpilled;
    }

    internal long PeakRetainedObjectBytes { get; }
    internal bool ObjectBufferSpilled { get; }
}
