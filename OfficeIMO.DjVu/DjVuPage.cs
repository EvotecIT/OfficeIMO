namespace OfficeIMO.DjVu;

/// <summary>A stable page in an owned DjVu document.</summary>
public sealed partial class DjVuPage {
    internal readonly DjVuDocument Document;
    internal readonly DjVuComponent Component;
    private readonly DjVuTextResult _text;
    internal DjVuPage(DjVuDocument document, DjVuComponent component, int number, DjVuReadBudget budget) {
        Document = document; Component = component; Number = number;
        if (component.Form.FormType == "PM44" || component.Form.FormType == "BM44") {
            var chunks = component.Form.Children;
            if (chunks.Count == 0 || chunks.Any(c => c.Id != component.Form.FormType) || chunks[0].Length < 9)
                throw new InvalidDataException("Invalid standalone IW44 image container.");
            var first = chunks[0];
            Width = DjVuBinary.U16(first.Source, first.Offset + 4);
            Height = DjVuBinary.U16(first.Source, first.Offset + 6);
            if (Width == 0 || Height == 0) throw new InvalidDataException("Empty standalone IW44 image.");
            if ((long)Width * Height > budget.Options.MaxPagePixels) throw new DjVuResourceLimitException(nameof(DjVuReadOptions.MaxPagePixels));
            Dpi = 100; Gamma = 2.2; Rotation = 0;
            _text = new DjVuTextResult(DjVuTextStatus.Absent, string.Empty);
            return;
        }
        var info = component.Form.Children.Where(c => c.Id == "INFO").ToList();
        if (info.Count != 1 || info[0].Length < 5 || component.Form.Children[0].Id != "INFO")
            throw new InvalidDataException("DjVu page must have one leading valid INFO chunk.");
        byte[] data = info[0].Source;
        int offset = info[0].Offset, length = info[0].Length;
        Width = DjVuBinary.U16(data, offset); Height = DjVuBinary.U16(data, offset + 2);
        if (Width == 0 || Height == 0) throw new InvalidDataException("DjVu page has empty dimensions.");
        if ((long)Width * Height > budget.Options.MaxPagePixels) throw new DjVuResourceLimitException(nameof(DjVuReadOptions.MaxPagePixels));
        int version = data[offset + 4] | (length > 5 ? data[offset + 5] << 8 : 0);
        if (version > 26) throw new NotSupportedException("Unsupported DjVu page version.");
        int dpi = length >= 8 ? data[offset + 6] | data[offset + 7] << 8 : 300;
        Dpi = dpi >= 25 && dpi <= 6000 ? dpi : 300;
        Gamma = length >= 9 ? Math.Max(3, Math.Min(50, (int)data[offset + 8])) / 10.0 : 2.2;
        int orientation = length >= 10 ? data[offset + 9] & 7 : 1;
        Rotation = orientation == 6 ? 270 : orientation == 2 ? 180 : orientation == 5 ? 90 : 0;
        var texts = document.TextChunks(component, budget);
        _text = texts.Count > 1 ? new DjVuTextResult(DjVuTextStatus.Corrupt, string.Empty, diagnostic: "Page contains multiple text layers.")
            : DjVuTextReader.Read(texts.FirstOrDefault(), budget);
    }

    /// <summary>One-based position in the document.</summary>
    public int Number { get; }
    /// <summary>Exact source component identity.</summary>
    public string Id => Component.Id;
    /// <summary>Display title stored in the directory.</summary>
    public string Title => Component.Title;
    /// <summary>Unrotated native pixel width.</summary>
    public int Width { get; }
    /// <summary>Unrotated native pixel height.</summary>
    public int Height { get; }
    /// <summary>Declared page resolution in dots per inch, normalized to the format's valid range.</summary>
    public int Dpi { get; }
    /// <summary>Declared image gamma.</summary>
    public double Gamma { get; }
    /// <summary>Clockwise display rotation in degrees: 0, 90, 180 or 270.</summary>
    public int Rotation { get; }
    /// <summary>Returns existing stored text and geometry. Does not run OCR.</summary>
    public DjVuTextResult GetText(CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return _text;
    }
}
