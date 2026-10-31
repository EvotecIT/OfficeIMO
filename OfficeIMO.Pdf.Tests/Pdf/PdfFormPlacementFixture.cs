using System.Globalization;
using System.Text;
using OfficeIMO.Pdf;

namespace OfficeIMO.Tests.Pdf;

internal static class PdfFormPlacementFixture {
    internal static byte[] CreateAuthored(string kind) => PdfDocument.Create(pdf => pdf.Content(content => {
        if (kind == "choice") {
            content.ChoiceField("Value", new[] { "HgjQ", "QgjH" }, "HgjQ", width: global::OfficeIMO.OfficeLength.Points(180), height: global::OfficeIMO.OfficeLength.Points(18), fontSize: global::OfficeIMO.OfficeLength.Points(14));
        } else {
            content.TextField("Value", width: global::OfficeIMO.OfficeLength.Points(180), height: global::OfficeIMO.OfficeLength.Points(18), value: "HgjQ", fontSize: global::OfficeIMO.OfficeLength.Points(14),
                style: new PdfFormFieldStyle { IsComb = kind == "comb", MaxLength = kind == "comb" ? 4 : null });
        }
    })).ToBytes();

    internal static byte[] CreateWithoutAppearance(int? catalogAlignment, int? parentAlignment, int? widgetAlignment, bool nested = false) {
        static string Alignment(int? alignment) => alignment.HasValue ? " /Q " + alignment.Value.ToString(CultureInfo.InvariantCulture) : string.Empty;
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 120] /Contents 4 0 R /Resources << >> /Annots [8 0 R] >>",
            "<< /Length 0 >>\nstream\n\nendstream",
            "<< /Fields [" + (nested ? "9" : "7") + " 0 R] /DA (/Helv 12 Tf 0 g) /DR << /Font << /Helv 6 0 R >> >>" + Alignment(catalogAlignment) + " >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "<< /FT /Tx /T (Value) /V (HgjQ) /Kids [8 0 R]" + Alignment(nested ? null : parentAlignment) + " >>",
            "<< /Type /Annot /Subtype /Widget /Parent 7 0 R /P 3 0 R /Rect [20 40 200 58] /F 4" + Alignment(widgetAlignment) + " >>"
        };
        if (nested) objects.Add("<< /T (Group) /Kids [7 0 R]" + Alignment(parentAlignment) + " >>");
        var pdf = new StringBuilder("%PDF-1.4\n");
        var offsets = new List<int> { 0 };
        for (int index = 0; index < objects.Count; index++) {
            offsets.Add(Encoding.ASCII.GetByteCount(pdf.ToString()));
            pdf.Append(index + 1).Append(" 0 obj\n").Append(objects[index]).Append("\nendobj\n");
        }
        int xref = Encoding.ASCII.GetByteCount(pdf.ToString());
        pdf.Append("xref\n0 ").Append(offsets.Count).Append("\n0000000000 65535 f \n");
        for (int index = 1; index < offsets.Count; index++) pdf.Append(offsets[index].ToString("D10", CultureInfo.InvariantCulture)).Append(" 00000 n \n");
        pdf.Append("trailer\n<< /Root 1 0 R /Size ").Append(offsets.Count).Append(" >>\nstartxref\n").Append(xref).Append("\n%%EOF\n");
        return Encoding.ASCII.GetBytes(pdf.ToString());
    }
}
