using System.Globalization;
using System.IO.Compression;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static int AddObject(System.Collections.Generic.List<byte[]> list, string body) {
        int id = list.Count + 1;
        list.Add(PdfObjectBytes.WrapIndirectObject(id, body));
        return id;
    }

    private static int ReserveObject(System.Collections.Generic.List<byte[]> list) {
        return AddObject(list, "<< >>\n");
    }

    private static void ReplaceObject(System.Collections.Generic.List<byte[]> list, int id, string body) {
        Guard.NotNull(list, nameof(list));
        if (id < 1 || id > list.Count) {
            throw new ArgumentOutOfRangeException(nameof(id), "PDF object id is outside the current object table.");
        }

        list[id - 1] = PdfObjectBytes.WrapIndirectObject(id, body);
    }

    private static int AddStreamObject(System.Collections.Generic.List<byte[]> list, byte[] content) {
        Guard.NotNull(content, nameof(content));
        return AddStreamObject(
            list,
            "<< /Length " + content.Length.ToString(CultureInfo.InvariantCulture) + " >>",
            content);
    }

    private static int AddFlateStreamObject(System.Collections.Generic.List<byte[]> list, byte[] content) {
        Guard.NotNull(content, nameof(content));
        byte[] compressed = DeflateZlib(content);
        return AddStreamObject(
            list,
            "<< /Length " + compressed.Length.ToString(CultureInfo.InvariantCulture) + " /Filter /FlateDecode >>",
            compressed);
    }

    private static int AddFlateStreamObject(System.Collections.Generic.List<byte[]> list, byte[] content, string extraDictionaryEntries) {
        Guard.NotNull(content, nameof(content));
        Guard.NotNull(extraDictionaryEntries, nameof(extraDictionaryEntries));
        byte[] compressed = DeflateZlib(content);
        string trimmedEntries = extraDictionaryEntries.Trim();
        string entries = trimmedEntries.Length == 0 ? string.Empty : " " + trimmedEntries;
        return AddStreamObject(
            list,
            "<< /Length " + compressed.Length.ToString(CultureInfo.InvariantCulture) + entries + " /Filter /FlateDecode >>",
            compressed);
    }

    private static int AddStreamObject(System.Collections.Generic.List<byte[]> list, string dictionary, byte[] content) {
        Guard.NotNull(content, nameof(content));
        Guard.NotNullOrWhiteSpace(dictionary, nameof(dictionary));

        int id = list.Count + 1;
        list.Add(PdfObjectBytes.WrapStreamObject(id, dictionary, content));
        return id;
    }

    private static byte[] DeflateZlib(byte[] data) {
        using var output = new MemoryStream();
        output.WriteByte(0x78);
        output.WriteByte(0x9C);
        using (var deflate = new DeflateStream(output, CompressionLevel.Optimal, leaveOpen: true)) {
            deflate.Write(data, 0, data.Length);
        }

        uint adler = Adler32(data);
        output.WriteByte((byte)((adler >> 24) & 0xFF));
        output.WriteByte((byte)((adler >> 16) & 0xFF));
        output.WriteByte((byte)((adler >> 8) & 0xFF));
        output.WriteByte((byte)(adler & 0xFF));
        return output.ToArray();
    }

    private static string PdfString(string s) {
        return PdfSyntaxEscaper.LiteralString(s);
    }

    private sealed class LayoutResult {
        public System.Collections.Generic.List<Page> Pages { get; } = new();
        public bool UsedBold { get; set; }
        public bool UsedItalic { get; set; }
        public bool UsedBoldItalic { get; set; }
        public sealed class Page {
            public PdfOptions Options { get; set; } = null!;
            public int PageGroupId { get; set; }
            public string Content { get; set; } = string.Empty;
            public System.Collections.Generic.List<LinkAnnotation> Annotations { get; } = new();
            public System.Collections.Generic.List<FormFieldAnnotation> FormFields { get; } = new();
            public System.Collections.Generic.List<PageImage> Images { get; } = new();
            public System.Collections.Generic.List<PageGraphicsState> GraphicsStates { get; } = new();
            public System.Collections.Generic.List<PageShading> Shadings { get; } = new();
            public System.Collections.Generic.List<PageBookmark> Bookmarks { get; } = new();
            public System.Collections.Generic.List<PageNamedDestination> NamedDestinations { get; } = new();
            public System.Collections.Generic.HashSet<PdfStandardFont> UsedFonts { get; } = new();
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
        public string? Uri { get; init; }
        public string? DestinationName { get; init; }
        public string? Contents { get; init; }
    }

    private sealed class FormFieldAnnotation {
        public double X1 { get; init; }
        public double Y1 { get; init; }
        public double X2 { get; init; }
        public double Y2 { get; init; }
        public FormFieldAnnotationKind Kind { get; init; }
        public string Name { get; init; } = string.Empty;
        public string Value { get; init; } = string.Empty;
        public IReadOnlyList<string> Values { get; init; } = Array.Empty<string>();
        public double FontSize { get; init; }
        public bool IsChecked { get; init; }
        public string CheckedValueName { get; init; } = "Yes";
        public IReadOnlyList<string> Options { get; init; } = Array.Empty<string>();
        public double ButtonSize { get; init; }
        public double ButtonGap { get; init; }
        public PdfFormFieldStyle Style { get; init; } = new PdfFormFieldStyle();
        public bool IsComboBox { get; init; }
        public bool AllowsMultipleSelection { get; init; }
    }

    private enum FormFieldAnnotationKind {
        Text,
        CheckBox,
        Choice,
        RadioButtonGroup
    }

    private sealed class PageBookmark {
        public int Level { get; init; }
        public string Title { get; init; } = string.Empty;
        public double Y { get; init; }
    }

    private sealed class PageNamedDestination {
        public string Name { get; init; } = string.Empty;
        public double Y { get; init; }
    }

    private sealed class PageNumberInfo {
        public int VariantPageNumber { get; }
        public int PageNumber { get; }
        public int TotalPages { get; }

        public PageNumberInfo(int variantPageNumber, int pageNumber, int totalPages) {
            VariantPageNumber = variantPageNumber;
            PageNumber = pageNumber;
            TotalPages = totalPages;
        }
    }

    private sealed class PageGraphicsState {
        public string Name { get; set; } = string.Empty;
        public double FillOpacity { get; set; } = 1D;
        public double StrokeOpacity { get; set; } = 1D;
    }

    private sealed class PageShading {
        public string Name { get; set; } = string.Empty;
        public OfficeColor StartColor { get; set; }
        public OfficeColor EndColor { get; set; }
        public double X0 { get; set; }
        public double Y0 { get; set; }
        public double X1 { get; set; }
        public double Y1 { get; set; }
    }

    private sealed class OutlineNode {
        public int Id { get; set; }
        public int Level { get; init; }
        public int PageIndex { get; init; }
        public double Y { get; init; }
        public string Title { get; init; } = string.Empty;
        public OutlineNode? Parent { get; set; }
        public System.Collections.Generic.List<OutlineNode> Children { get; } = new();
    }

    private sealed class PageImage {
        public byte[] Data { get; init; } = System.Array.Empty<byte>();
        public OfficeImageInfo Info { get; init; } = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        public double X { get; init; }
        public double Y { get; init; }
        public double W { get; init; }
        public double H { get; init; }
        public OfficeClipPath? ClipPath { get; init; }
        public double ClipX { get; init; }
        public double ClipY { get; init; }
        public double ClipHeight { get; init; }
        public bool IsBackgroundDecoration { get; init; }
        public double Opacity { get; init; } = 1D;
        public double RotationAngle { get; init; }
        public string? GraphicsStateName { get; init; }
        public string Name { get; set; } = string.Empty;
        public int ObjectId { get; set; }
    }
}

