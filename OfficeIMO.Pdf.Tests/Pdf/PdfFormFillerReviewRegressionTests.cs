using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfFormFillerTests {
    private const string WidgetDefaults = "/DR << /Font << /Helv 4 0 R >> >> /DA (/Helv 12 Tf 0 g)";

    private static string TextWidget(int number, string name, string rect, string entries) =>
        number + " 0 obj\n<< /Type /Annot /Subtype /Widget /FT /Tx /T (" + name + ") /Rect [" + rect + "] " + entries + " >>\nendobj";

    private static string Stream(int number, string dictionary, string content) =>
        number + " 0 obj\n<< " + dictionary + " /Length " + content.Length + " >>\nstream\n" + content + "\nendstream\nendobj";

    private static string TextAppearance(int number, string text) =>
        Stream(number, "/Type /XObject /Subtype /Form /BBox [0 0 200 20] /Resources << /Font << /Helv 4 0 R >> >>", "BT /Helv 12 Tf 2 5 Td (" + text + ") Tj ET");

    private static byte[] BuildReviewFormPdf(string acroForm, string page, string catalog, params string[] objects) {
        var all = new List<string> {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /AcroForm << " + acroForm + " " + WidgetDefaults + " >> " + catalog + " >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R " + page + " >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            Stream(5, string.Empty, "q Q")
        };
        all.AddRange(objects);
        var pdf = new StringBuilder("%PDF-1.7\n");
        var offsets = new SortedDictionary<int, int>();
        foreach (string item in all) {
            offsets[int.Parse(item.Substring(0, item.IndexOf(' ')))] = pdf.Length;
            pdf.Append(item).Append('\n');
        }
        int xref = pdf.Length, size = offsets.Keys.Last() + 1;
        pdf.Append("xref\n0 ").Append(size).Append("\n0000000000 65535 f \n");
        for (int number = 1; number < size; number++) {
            pdf.Append(offsets.TryGetValue(number, out int offset) ? offset.ToString("D10") + " 00000 n \n" : "0000000000 00000 f \n");
        }
        pdf.Append("trailer\n<< /Root 1 0 R /Size ").Append(size).Append(" >>\nstartxref\n").Append(xref).Append("\n%%EOF\n");
        return Encoding.ASCII.GetBytes(pdf.ToString());
    }

    private static string Latin1(byte[] pdf) => Encoding.GetEncoding("ISO-8859-1").GetString(pdf);

    [Theory]
    [InlineData(2)]
    [InlineData(32)]
    public void FlattenFields_DoesNotPaintWidgetsViewersNeverShow(int annotationFlags) {
        byte[] input = BuildReviewFormPdf("/Fields [10 0 R 11 0 R]", "/Annots [10 0 R 11 0 R]", string.Empty,
            TextWidget(10, "Hidden", "72 700 272 720", "/F " + annotationFlags + " /V (HIDDENVALUE) /AP << /N 12 0 R >>"),
            TextWidget(11, "Shown", "72 650 272 670", "/F 4 /V (shown) /AP << /N 13 0 R >>"),
            TextAppearance(12, "HIDDENVALUE"), TextAppearance(13, "shown"));

        byte[] flattened = PdfFormFiller.FlattenFields(input);

        Assert.DoesNotContain("HIDDENVALUE", Latin1(flattened), StringComparison.Ordinal);
        string text = PdfDocument.Load(flattened).ExtractText();
        Assert.Contains("shown", text, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Widget", Latin1(flattened), StringComparison.Ordinal);
    }
}
