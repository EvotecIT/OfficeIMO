using System.Globalization;

namespace OfficeIMO.Pdf;

internal static class PdfFileAssembler {
    internal static byte[] Assemble(IReadOnlyList<byte[]> objects, int catalogId, int infoId, PdfFileVersion fileVersion = PdfFileVersion.Pdf14, PdfStandardEncryptionOptions? encryption = null) {
        using var stream = new MemoryStream();
        Assemble(stream, objects, catalogId, infoId, fileVersion, encryption);
        return stream.ToArray();
    }

    internal static long Assemble(Stream destination, IReadOnlyList<byte[]> objects, int catalogId, int infoId, PdfFileVersion fileVersion = PdfFileVersion.Pdf14, PdfStandardEncryptionOptions? encryption = null) {
        Guard.FileVersion(fileVersion, nameof(fileVersion));
        Guard.NotNull(destination, nameof(destination));
        Guard.NotNull(objects, nameof(objects));
        if (!destination.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(destination));

        PdfEncryptionAssembly? encryptionAssembly = null;
        if (encryption != null) {
            fileVersion = RequireAtLeast(fileVersion, GetMinimumEncryptionVersion(encryption.Algorithm));
            encryptionAssembly = PdfStandardSecurityWriter.Encrypt(objects, encryption);
            objects = encryptionAssembly.Objects;
        }

        byte[] header = PdfEncoding.Latin1GetBytes("%PDF-" + GetHeaderVersion(fileVersion) + "\n%\u00e2\u00e3\u00cf\u00d3\n");
        destination.Write(header, 0, header.Length);
        long written = header.LongLength;

        var offsets = new List<long> { 0L };
        for (int i = 0; i < objects.Count; i++) {
            offsets.Add(written);
            byte[] obj = objects[i];
            destination.Write(obj, 0, obj.Length);
            written += obj.LongLength;
        }

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
            .Append(BuildTrailerSecurityEntries(encryptionAssembly)).Append(" >>\n");
        trailer.Append("startxref\n").Append(xrefPos.ToString(CultureInfo.InvariantCulture)).Append("\n%%EOF\n");
        byte[] trailerBytes = Encoding.ASCII.GetBytes(trailer.ToString());
        destination.Write(trailerBytes, 0, trailerBytes.Length);
        return written + trailerBytes.LongLength;
    }

    private static string BuildTrailerSecurityEntries(PdfEncryptionAssembly? encryptionAssembly) {
        if (encryptionAssembly == null) {
            return string.Empty;
        }

        string id = PdfSyntaxEscaper.HexString(encryptionAssembly.FileId);
        return " /Encrypt " + PdfSyntaxEscaper.IndirectReference(encryptionAssembly.EncryptionObjectNumber) +
            " /ID [" + id + " " + id + "]";
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
