using System.Globalization;

namespace OfficeIMO.Pdf;

internal static class PdfFileAssembler {
    internal static byte[] Assemble(IReadOnlyList<byte[]> objects, int catalogId, int infoId, PdfFileVersion fileVersion = PdfFileVersion.Pdf14) {
        Guard.FileVersion(fileVersion, nameof(fileVersion));

        using var ms = new MemoryStream();
        byte[] header = PdfEncoding.Latin1GetBytes("%PDF-" + GetHeaderVersion(fileVersion) + "\n%\u00e2\u00e3\u00cf\u00d3\n");
        ms.Write(header, 0, header.Length);

        var offsets = new List<long> { 0L };
        for (int i = 0; i < objects.Count; i++) {
            offsets.Add(ms.Position);
            byte[] obj = objects[i];
            ms.Write(obj, 0, obj.Length);
        }

        long xrefPos = ms.Position;
        using var writer = new StreamWriter(ms, Encoding.ASCII, 1024, leaveOpen: true) { NewLine = "\n" };
        writer.WriteLine("xref");
        writer.WriteLine("0 " + (objects.Count + 1).ToString(CultureInfo.InvariantCulture));
        writer.WriteLine("0000000000 65535 f ");
        for (int i = 1; i <= objects.Count; i++) {
            writer.WriteLine(offsets[i].ToString("0000000000", CultureInfo.InvariantCulture) + " 00000 n ");
        }

        writer.WriteLine("trailer");
        writer.WriteLine("<< /Size " + (objects.Count + 1).ToString(CultureInfo.InvariantCulture) + " /Root " + PdfSyntaxEscaper.IndirectReference(catalogId) + " /Info " + PdfSyntaxEscaper.IndirectReference(infoId) + " >>");
        writer.WriteLine("startxref");
        writer.WriteLine(xrefPos.ToString(CultureInfo.InvariantCulture));
        writer.WriteLine("%%EOF");
        writer.Flush();

        return ms.ToArray();
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
