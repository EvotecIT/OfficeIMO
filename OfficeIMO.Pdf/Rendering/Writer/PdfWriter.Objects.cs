using System.Globalization;
using System.Text;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static int AddObject(System.Collections.Generic.List<byte[]> list, string body) {
        int id = list.Count + 1;
        var bytes = PdfEncoding.Latin1GetBytes(id.ToString(CultureInfo.InvariantCulture) + " 0 obj\n" + body + "endobj\n");
        list.Add(bytes);
        return id;
    }

    private static byte[] Merge(params byte[][] arrays) {
        int len = arrays.Sum(a => a.Length);
        var buf = new byte[len];
        int pos = 0;
        foreach (var a in arrays) { Buffer.BlockCopy(a, 0, buf, pos, a.Length); pos += a.Length; }
        return buf;
    }

    private static string PdfString(string s) {
        if (RequiresUnicodeLiteral(s)) {
            return "<" + EncodeUtf16Hex(s) + ">";
        }
        // Literal string in parentheses with robust escaping (incl. control chars via octal)
        return "(" + EscapeLiteral(s) + ")";
    }

    private static bool RequiresUnicodeLiteral(string s) {
        for (int i = 0; i < s.Length; i++) {
            if (s[i] > 0xFF) return true;
        }
        return false;
    }

    private static string EncodeUtf16Hex(string s) {
        if (s.Length == 0) return string.Empty;
        var bytes = Encoding.BigEndianUnicode.GetBytes(s);
        var sb = new StringBuilder((bytes.Length + 2) * 2);
        sb.Append("FEFF");
        for (int i = 0; i < bytes.Length; i++) {
            sb.Append(bytes[i].ToString("X2", CultureInfo.InvariantCulture));
        }
        return sb.ToString();
    }

    private sealed class LayoutResult {
        public System.Collections.Generic.List<Page> Pages { get; } = new();
        public bool UsedBold { get; set; }
        public bool UsedItalic { get; set; }
        public bool UsedBoldItalic { get; set; }
        public sealed class Page {
            public PdfOptions Options { get; set; } = null!;
            public string Content { get; set; } = string.Empty;
            public System.Collections.Generic.List<LinkAnnotation> Annotations { get; } = new();
            public System.Collections.Generic.List<PageImage> Images { get; } = new();
            public bool UsedBold { get; set; }
            public bool UsedItalic { get; set; }
            public bool UsedBoldItalic { get; set; }
        }
    }

    private sealed class LinkAnnotation {
        public double X1 { get; init; }
        public double Y1 { get; init; }
        public double X2 { get; init; }
        public double Y2 { get; init; }
        public string Uri { get; init; } = string.Empty;
    }

    private sealed class PageImage {
        public byte[] Data { get; init; } = System.Array.Empty<byte>();
        public double X { get; init; }
        public double Y { get; init; }
        public double W { get; init; }
        public double H { get; init; }
        public string Name { get; set; } = string.Empty;
        public int ObjectId { get; set; }
    }
}

