using System.Globalization;
using System.Text;

namespace OfficeIMO.Pdf;

internal static class PdfLinearizedFileAssembler {
    private const int PlaceholderWidth = 20;

    internal static byte[] Assemble(IReadOnlyList<byte[]> bodies, int catalogId, int infoId, PdfFileVersion fileVersion) {
        var wrapped = new List<byte[]>(bodies.Count);
        for (int i = 0; i < bodies.Count; i++) wrapped.Add(PdfObjectBytes.WrapIndirectObject(i + 1, bodies[i]));
        byte[] probePdf = PdfFileAssembler.Assemble(wrapped, catalogId, infoId, fileVersion);
        PdfReadDocument probe = PdfReadDocument.Load(probePdf);
        if (probe.Pages.Count == 0) throw new NotSupportedException("Linearization requires at least one readable page.");
        int firstPageObjectNumber = probe.Pages[0].ObjectNumber;
        int linearizationObjectNumber = bodies.Count + 1;
        int size = linearizationObjectNumber + 1;

        using var output = new MemoryStream();
        Write(output, "%PDF-" + PdfFileAssembler.GetHeaderVersion(fileVersion) + "\n%\u00e2\u00e3\u00cf\u00d3\n");
        long linearizationOffset = output.Position;
        string placeholder = new string('0', PlaceholderWidth);
        string linearization = "<< /Linearized 1 /L " + placeholder + " /H [0 0] /O " + firstPageObjectNumber.ToString(CultureInfo.InvariantCulture) + " /E " + placeholder + " /N " + probe.Pages.Count.ToString(CultureInfo.InvariantCulture) + " /T " + placeholder + " >>\n";
        Write(output, PdfObjectBytes.WrapIndirectObject(linearizationObjectNumber, PdfEncoding.Latin1GetBytes(linearization)));

        var offsets = new long[size]; offsets[linearizationObjectNumber] = linearizationOffset;
        offsets[firstPageObjectNumber] = output.Position; Write(output, PdfObjectBytes.WrapIndirectObject(firstPageObjectNumber, bodies[firstPageObjectNumber - 1]));
        for (int id = 1; id <= bodies.Count; id++) {
            if (id == firstPageObjectNumber) continue;
            offsets[id] = output.Position; Write(output, PdfObjectBytes.WrapIndirectObject(id, bodies[id - 1]));
        }
        long xrefOffset = output.Position;
        using (var writer = new StreamWriter(output, Encoding.ASCII, 1024, true) { NewLine = "\n" }) {
            writer.WriteLine("xref"); writer.WriteLine("0 " + size.ToString(CultureInfo.InvariantCulture)); writer.WriteLine("0000000000 65535 f ");
            for (int id = 1; id < size; id++) writer.WriteLine(offsets[id].ToString("0000000000", CultureInfo.InvariantCulture) + " 00000 n ");
            writer.WriteLine("trailer"); writer.WriteLine("<< /Size " + size.ToString(CultureInfo.InvariantCulture) + " /Root " + PdfSyntaxEscaper.IndirectReference(catalogId) + (infoId > 0 ? " /Info " + PdfSyntaxEscaper.IndirectReference(infoId) : string.Empty) + " >>");
            writer.WriteLine("startxref"); writer.WriteLine(xrefOffset.ToString(CultureInfo.InvariantCulture)); writer.WriteLine("%%EOF"); writer.Flush();
        }
        byte[] result = output.ToArray();
        PatchPlaceholder(result, "/L ", result.LongLength);
        PatchPlaceholder(result, "/E ", xrefOffset);
        PatchPlaceholder(result, "/T ", xrefOffset);
        return result;
    }

    private static void PatchPlaceholder(byte[] pdf, string marker, long value) {
        byte[] markerBytes = Encoding.ASCII.GetBytes(marker); int index = IndexOf(pdf, markerBytes) + markerBytes.Length;
        byte[] replacement = Encoding.ASCII.GetBytes(value.ToString("D" + PlaceholderWidth.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture));
        Buffer.BlockCopy(replacement, 0, pdf, index, replacement.Length);
    }

    private static int IndexOf(byte[] source, byte[] pattern) {
        for (int i = 0; i <= source.Length - pattern.Length; i++) { bool match = true; for (int j = 0; j < pattern.Length; j++) if (source[i + j] != pattern[j]) { match = false; break; } if (match) return i; }
        throw new InvalidOperationException("Linearization placeholder was not found.");
    }

    private static void Write(Stream output, string value) => Write(output, PdfEncoding.Latin1GetBytes(value));
    private static void Write(Stream output, byte[] bytes) => output.Write(bytes, 0, bytes.Length);
}
