using System.Globalization;

namespace OfficeIMO.Pdf;

internal static class PdfFileAssembler {
    internal static byte[] Assemble(IReadOnlyList<byte[]> objects, int catalogId, int infoId) {
        using var ms = new MemoryStream();
        byte[] header = PdfEncoding.Latin1GetBytes("%PDF-1.4\n%\u00e2\u00e3\u00cf\u00d3\n");
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
}
