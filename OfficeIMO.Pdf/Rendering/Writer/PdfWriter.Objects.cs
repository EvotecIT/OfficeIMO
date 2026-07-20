using System.Globalization;
using System.IO.Compression;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static int AddObject(System.Collections.Generic.IList<byte[]> list, string body) {
        int id = list.Count + 1;
        list.Add(PdfObjectBytes.WrapIndirectObject(id, body));
        return id;
    }

    private static int ReserveObject(System.Collections.Generic.IList<byte[]> list) {
        return AddObject(list, "<< >>\n");
    }

    private static void ReplaceObject(System.Collections.Generic.IList<byte[]> list, int id, string body) {
        Guard.NotNull(list, nameof(list));
        if (id < 1 || id > list.Count) {
            throw new ArgumentOutOfRangeException(nameof(id), "PDF object id is outside the current object table.");
        }

        list[id - 1] = PdfObjectBytes.WrapIndirectObject(id, body);
    }

    private static int AddStreamObject(System.Collections.Generic.IList<byte[]> list, byte[] content) {
        Guard.NotNull(content, nameof(content));
        return AddStreamObject(
            list,
            "<< /Length " + content.Length.ToString(CultureInfo.InvariantCulture) + " >>",
            content);
    }

    private static int AddFlateStreamObject(System.Collections.Generic.IList<byte[]> list, byte[] content) {
        Guard.NotNull(content, nameof(content));
        byte[] compressed = DeflateZlib(content);
        return AddStreamObject(
            list,
            "<< /Length " + compressed.Length.ToString(CultureInfo.InvariantCulture) + " /Filter /FlateDecode >>",
            compressed);
    }

    private static int AddFlateStreamObject(System.Collections.Generic.IList<byte[]> list, byte[] content, string extraDictionaryEntries) {
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

    private static int AddStreamObject(System.Collections.Generic.IList<byte[]> list, string dictionary, byte[] content) {
        Guard.NotNull(content, nameof(content));
        Guard.NotNullOrWhiteSpace(dictionary, nameof(dictionary));

        int id = list.Count + 1;
        if (list is PdfObjectStore objectStore) {
            objectStore.AddSegments(PdfObjectBytes.CreateStreamObjectSegments(id, dictionary, content));
        } else {
            list.Add(PdfObjectBytes.WrapStreamObject(id, dictionary, content));
        }
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

    private sealed class LayoutResult : IDisposable {
        private readonly PdfPageContentStore _contentStore;

        public LayoutResult(PdfPageContentStore contentStore) {
            _contentStore = contentStore;
        }

        public System.Collections.Generic.List<Page> Pages { get; } = new();
        public bool UsedBold { get; set; }
        public bool UsedItalic { get; set; }
        public bool UsedBoldItalic { get; set; }
        public bool HasTableOfContents { get; set; }
        public System.Collections.Generic.List<SectionBlock> SectionDefinitions { get; } = new();
        public sealed class Page {
            public PdfOptions Options { get; set; } = null!;
            public int PageGroupId { get; set; }
            public PdfPageContentHandle Content { get; set; }
            public System.Collections.Generic.List<LinkAnnotation> Annotations { get; } = new();
            public System.Collections.Generic.List<TextAnnotation> TextAnnotations { get; } = new();
            public System.Collections.Generic.List<FreeTextAnnotation> FreeTextAnnotations { get; } = new();
            public System.Collections.Generic.List<HighlightAnnotation> HighlightAnnotations { get; } = new();
            public System.Collections.Generic.List<FormFieldAnnotation> FormFields { get; } = new();
            public System.Collections.Generic.List<PageImage> Images { get; } = new();
            public System.Collections.Generic.List<PageGraphicsState> GraphicsStates { get; } = new();
            public System.Collections.Generic.List<PageShading> Shadings { get; } = new();
            public System.Collections.Generic.List<PageEffectGroup> EffectGroups { get; } = new();
            public System.Collections.Generic.List<PageBookmark> Bookmarks { get; } = new();
            public System.Collections.Generic.List<PageNamedDestination> NamedDestinations { get; } = new();
            public System.Collections.Generic.List<PageSection> Sections { get; } = new();
            public System.Collections.Generic.List<PdfLayerDefinition> Layers { get; } = new();
            public System.Collections.Generic.List<PageStructElement> StructElements { get; } = new();
            public System.Collections.Generic.List<PdfGeneratedDrawingAccessibilityEvidence> Drawings { get; } = new();
            public System.Collections.Generic.HashSet<PdfStandardFont> UsedFonts { get; } = new();
            public System.Collections.Generic.HashSet<PdfNamedFontFace> UsedNamedFonts { get; } = new();
            public int? StructParentIndex { get; set; }
            public int NextMarkedContentId { get; set; }
            public bool UsedBold { get; set; }
            public bool UsedItalic { get; set; }
            public bool UsedBoldItalic { get; set; }
        }

        public string ReadContent(PdfPageContentHandle handle) => _contentStore.Read(handle);

        public void Dispose() => _contentStore.Dispose();
    }

    private sealed class LinkAnnotation {
        public double X1 { get; set; }
        public double Y1 { get; set; }
        public double X2 { get; set; }
        public double Y2 { get; set; }
        public string? Uri { get; set; }
        public string? DestinationName { get; set; }
        public string? Contents { get; set; }
        public int? StructParentIndex { get; set; }
        public int? StructElementIndex { get; set; }
        public PageImage? LinkedImage { get; set; }
        public int ObjectId { get; set; }
    }

    private sealed class TextAnnotation {
        public double X1 { get; set; }
        public double Y1 { get; set; }
        public double X2 { get; set; }
        public double Y2 { get; set; }
        public string Contents { get; set; } = string.Empty;
        public PdfTextAnnotationIcon Icon { get; set; }
        public PdfColor? Color { get; set; }
        public bool Open { get; set; }
        public int ObjectId { get; set; }
    }

    private sealed class FreeTextAnnotation {
        public double X1 { get; set; }
        public double Y1 { get; set; }
        public double X2 { get; set; }
        public double Y2 { get; set; }
        public string Contents { get; set; } = string.Empty;
        public double FontSize { get; set; }
        public PdfColor TextColor { get; set; }
        public PdfColor? BorderColor { get; set; }
        public double BorderWidth { get; set; }
        public PdfColor? FillColor { get; set; }
        public PdfAlign TextAlign { get; set; }
        public double Padding { get; set; }
        public double? LineHeight { get; set; }
        public int ObjectId { get; set; }
    }

    private sealed class HighlightAnnotation {
        public double X1 { get; set; }
        public double Y1 { get; set; }
        public double X2 { get; set; }
        public double Y2 { get; set; }
        public string Contents { get; set; } = string.Empty;
        public PdfColor Color { get; set; }
        public int ObjectId { get; set; }
    }

    private sealed class FormFieldAnnotation {
        public double X1 { get; set; }
        public double Y1 { get; set; }
        public double X2 { get; set; }
        public double Y2 { get; set; }
        public FormFieldAnnotationKind Kind { get; set; }
        public string Name { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;
        public IReadOnlyList<string> Values { get; set; } = Array.Empty<string>();
        public double FontSize { get; set; }
        public bool IsChecked { get; set; }
        public string CheckedValueName { get; set; } = "Yes";
        public IReadOnlyList<string> Options { get; set; } = Array.Empty<string>();
        public double ButtonSize { get; set; }
        public double ButtonGap { get; set; }
        public PdfFormFieldStyle Style { get; set; } = new PdfFormFieldStyle();
        public bool IsComboBox { get; set; }
        public bool AllowsMultipleSelection { get; set; }
    }

    private sealed class AnnotationStructureReference {
        public int StructParentIndex { get; set; }
        public int StructElementIndex { get; set; }
        public int ObjectId { get; set; }
    }

    private enum FormFieldAnnotationKind {
        Text,
        CheckBox,
        Choice,
        RadioButtonGroup
    }

    private sealed class PageBookmark {
        public int Level { get; set; }
        public string Title { get; set; } = string.Empty;
        public double Y { get; set; }
    }

    private sealed class PageNamedDestination {
        public string Name { get; set; } = string.Empty;
        public double Y { get; set; }
    }

    private sealed class PageSection {
        public string DestinationName { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public int Level { get; set; }
        public double Y { get; set; }
        public PdfSectionReference? Reference { get; set; }
    }

    private sealed class PageStructElement {
        public int? MarkedContentId { get; set; }
        public System.Collections.Generic.List<int>? AdditionalMarkedContentIds { get; set; }
        public string StructureType { get; set; } = "P";
        public string AlternativeText { get; set; } = string.Empty;
        public string TableHeaderScope { get; set; } = string.Empty;
        public int TableColumnSpan { get; set; } = 1;
        public int TableRowSpan { get; set; } = 1;
        public int? ParentElementIndex { get; set; }
        public PageStructElement? ParentElement { get; set; }
        public int? AnnotationObjectId { get; set; }
        public System.Collections.Generic.List<int>? AdditionalAnnotationObjectIds { get; set; }
        public int? AnnotationStructParentIndex { get; set; }
        public System.Collections.Generic.List<int>? AdditionalAnnotationStructParentIndexes { get; set; }
        public int ObjectId { get; set; }
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
        public bool IsRadial { get; set; }
        public System.Collections.Generic.IReadOnlyList<OfficeGradientStop> Stops { get; set; } = System.Array.Empty<OfficeGradientStop>();
        public double X0 { get; set; }
        public double Y0 { get; set; }
        public double R0 { get; set; }
        public double X1 { get; set; }
        public double Y1 { get; set; }
        public double R1 { get; set; }

        public bool MatchesAxial(double x0, double y0, double x1, double y1, System.Collections.Generic.IReadOnlyList<OfficeGradientStop> stops) =>
            !IsRadial && MatchesCoordinatesAndStops(x0, y0, 0D, x1, y1, 0D, stops);

        public bool MatchesRadial(double x0, double y0, double r0, double x1, double y1, double r1, System.Collections.Generic.IReadOnlyList<OfficeGradientStop> stops) =>
            IsRadial && MatchesCoordinatesAndStops(x0, y0, r0, x1, y1, r1, stops);

        private bool MatchesCoordinatesAndStops(double x0, double y0, double r0, double x1, double y1, double r1, System.Collections.Generic.IReadOnlyList<OfficeGradientStop> stops) {
            if (!X0.Equals(x0) || !Y0.Equals(y0) || !R0.Equals(r0) || !X1.Equals(x1) || !Y1.Equals(y1) || !R1.Equals(r1) || Stops.Count != stops.Count) {
                return false;
            }

            for (int index = 0; index < Stops.Count; index++) {
                if (!Stops[index].Equals(stops[index])) return false;
            }

            return true;
        }
    }

    private sealed class PageEffectGroup {
        public PdfPageContentHandle Content { get; set; }
        public string Token { get; set; } = string.Empty;
        public OfficeTransform Transform { get; set; } = OfficeTransform.Identity;
        public string? GraphicsStateName { get; set; }
        public string Name { get; set; } = string.Empty;
        public int ObjectId { get; set; }
    }

    private sealed class OutlineNode {
        public int Id { get; set; }
        public int Level { get; set; }
        public int PageIndex { get; set; }
        public double Y { get; set; }
        public string Title { get; set; } = string.Empty;
        public OutlineNode? Parent { get; set; }
        public System.Collections.Generic.List<OutlineNode> Children { get; } = new();
    }

    private sealed class PageImage {
        public byte[] Data { get; set; } = System.Array.Empty<byte>();
        public OfficeImageInfo Info { get; set; } = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        public double X { get; set; }
        public double Y { get; set; }
        public double W { get; set; }
        public double H { get; set; }
        public OfficeClipPath? ClipPath { get; set; }
        public double ClipX { get; set; }
        public double ClipY { get; set; }
        public double ClipHeight { get; set; }
        public PdfImageSourceCrop? SourceCrop { get; set; }
        public bool IsBackgroundDecoration { get; set; }
        public bool IsInlineDecoration { get; set; }
        public bool IsDecorativeArtifact => IsBackgroundDecoration || IsInlineDecoration;
        public double Opacity { get; set; } = 1D;
        public double RotationAngle { get; set; }
        public bool HorizontalFlip { get; set; }
        public bool VerticalFlip { get; set; }
        public string? GraphicsStateName { get; set; }
        public string? AlternativeText { get; set; }
        public bool SuppressAccessibilityWrapper { get; set; }
        public int? MarkedContentId { get; set; }
        public int? StructElementIndex { get; set; }
        public string? InlineDrawToken { get; set; }
        public bool DebugBox { get; set; }
        public string Name { get; set; } = string.Empty;
        public int ObjectId { get; set; }
    }
}
