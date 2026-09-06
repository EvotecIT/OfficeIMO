using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OfficeIMO.TestAssets;

internal static class SparseAnnotationTestSource {
    // Removing annotation 4 causes the retained annotation 20 to be rewritten as object 4.
    internal static byte[] Create() {
        var objects = new SortedDictionary<int, string> {
            [1] = "<< /Type /Catalog /Pages 2 0 R >>",
            [2] = "<< /Type /Pages /Kids [10 0 R] /Count 1 >>",
            [4] = "<< /Type /Annot /Subtype /Text /Rect [36 36 54 54] /Contents (Removed comment) /P 10 0 R >>",
            [10] = "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 400] /Resources << >> /Contents 30 0 R /Annots [4 0 R 20 0 R] >>",
            [20] = "<< /Type /Annot /Subtype /Text /Rect [80 80 98 98] /Contents (Retained comment) /P 10 0 R >>",
            [30] = "<< /Length 0 >>\nstream\n\nendstream"
        };
        var text = new StringBuilder("%PDF-1.4\n");
        var offsets = new Dictionary<int, int>();
        foreach (var item in objects) {
            offsets[item.Key] = Encoding.ASCII.GetByteCount(text.ToString());
            text.Append(item.Key).Append(" 0 obj\n").Append(item.Value).Append("\nendobj\n");
        }
        int xref = Encoding.ASCII.GetByteCount(text.ToString());
        text.Append("xref\n0 31\n0000000000 65535 f \n");
        for (int number = 1; number <= 30; number++) {
            text.Append(offsets.TryGetValue(number, out int offset)
                ? offset.ToString("D10", CultureInfo.InvariantCulture) + " 00000 n \n"
                : "0000000000 00000 f \n");
        }
        text.Append("trailer\n<< /Root 1 0 R /Size 31 >>\nstartxref\n").Append(xref).Append("\n%%EOF\n");
        return Encoding.ASCII.GetBytes(text.ToString());
    }
}
