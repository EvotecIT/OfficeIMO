using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.TestAssets;

/// <summary>Direct PDF syntax with native text, a nested Form XObject, an image and an AcroForm widget.</summary>
internal static class PdfInspectionProjectionFixture {
    internal static byte[] Build(int imageAliases = 1) {
        byte[] image = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(1, 1, OfficeColor.Blue), OfficeImageExportFormat.Jpeg);
        string resources = string.Join(" ", Enumerable.Range(0, imageAliases).Select(index => "/Im" + index.ToString(CultureInfo.InvariantCulture) + " 6 0 R"));
        byte[][] objects = {
            Ascii("<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [7 0 R] >> >>"),
            Ascii("<< /Type /Pages /Count 1 /Kids [3 0 R] >>"),
            Ascii("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 200] /Resources << /Font << /F1 4 0 R >> /XObject << " + resources + " /Fm 8 0 R >> >> /Contents 5 0 R /Annots [7 0 R] >>"),
            Ascii("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>"),
            Stream(Ascii("BT /F1 12 Tf 20 100 Td (Native page text) Tj ET q 40 0 0 40 10 10 cm /Im0 Do Q /Fm Do")),
            Stream(image, "/Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode"),
            Ascii("<< /Type /Annot /Subtype /Widget /FT /Tx /T (Entry) /V (Source value) /RV (<body>Rich source value</body>) /Rect [20 20 90 40] /P 3 0 R >>"),
            Stream(Ascii("BT /F1 12 Tf 20 70 Td (Nested form text) Tj ET"), "/Type /XObject /Subtype /Form /BBox [0 0 300 200] /Resources << /Font << /F1 4 0 R >> >>")
        };
        using var output = new MemoryStream(); Write(output, Ascii("%PDF-1.7\n"));
        var offsets = new List<long>();
        for (int index = 0; index < objects.Length; index++) {
            offsets.Add(output.Position); Write(output, Ascii((index + 1).ToString(CultureInfo.InvariantCulture) + " 0 obj\n"));
            Write(output, objects[index]); Write(output, Ascii("\nendobj\n"));
        }
        long xref = output.Position;
        Write(output, Ascii("xref\n0 " + (objects.Length + 1).ToString(CultureInfo.InvariantCulture) + "\n0000000000 65535 f \n"));
        foreach (long offset in offsets) Write(output, Ascii(offset.ToString("D10", CultureInfo.InvariantCulture) + " 00000 n \n"));
        Write(output, Ascii("trailer\n<< /Root 1 0 R /Size " + (objects.Length + 1).ToString(CultureInfo.InvariantCulture) + " >>\nstartxref\n" + xref.ToString(CultureInfo.InvariantCulture) + "\n%%EOF\n"));
        return output.ToArray();
    }
    private static byte[] Ascii(string value) => Encoding.ASCII.GetBytes(value);
    private static void Write(Stream output, byte[] bytes) => output.Write(bytes, 0, bytes.Length);
    private static byte[] Stream(byte[] bytes, string dictionary = "") => Ascii("<< " + dictionary + " /Length " + bytes.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n")
        .Concat(bytes).Concat(Ascii("\nendstream")).ToArray();
}
