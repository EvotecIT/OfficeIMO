using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

/// <summary>Page orientation used by a print-preview plan.</summary>
public enum PdfPrintOrientation {
    /// <summary>Choose orientation from the first source page on each sheet.</summary>
    Automatic,
    /// <summary>Portrait paper.</summary>
    Portrait,
    /// <summary>Landscape paper.</summary>
    Landscape
}

/// <summary>Scaling used while placing source pages on preview sheets.</summary>
public enum PdfPrintScaleMode {
    /// <summary>Show the whole source page without cropping.</summary>
    Fit,
    /// <summary>Use the source page's physical point size, reducing only when necessary.</summary>
    ActualSize,
    /// <summary>Fill the target slot and mark overflow as clipped.</summary>
    Fill
}

/// <summary>Request for a deterministic print-preview sheet plan.</summary>
public sealed class PdfPrintPlanRequest {
    /// <summary>Source PDF.</summary>
    public required string InputPath { get; set; }

    /// <summary>Document-relative selection such as <c>1-3,last</c>; all pages when omitted.</summary>
    public string? Pages { get; set; }

    /// <summary>Paper size.</summary>
    public PageSize PaperSize { get; set; } = PageSizes.A4;

    /// <summary>Paper orientation.</summary>
    public PdfPrintOrientation Orientation { get; set; } = PdfPrintOrientation.Automatic;

    /// <summary>Source pages placed on each paper sheet: 1, 2, or 4.</summary>
    public int PagesPerSheet { get; set; } = 1;

    /// <summary>Uniform printable margin in points.</summary>
    public double Margin { get; set; } = 18D;

    /// <summary>Source-page scaling behavior.</summary>
    public PdfPrintScaleMode ScaleMode { get; set; } = PdfPrintScaleMode.Fit;

    /// <summary>Optional source password.</summary>
    public string? PdfPassword { get; set; }
}

/// <summary>One source-page placement on a print-preview sheet.</summary>
public sealed class PdfPrintPlacement {
    internal PdfPrintPlacement(
        int pageNumber,
        double x,
        double y,
        double width,
        double height,
        double scale,
        bool clipped,
        double slotX,
        double slotY,
        double slotWidth,
        double slotHeight) {
        PageNumber = pageNumber;
        X = x;
        Y = y;
        Width = width;
        Height = height;
        Scale = scale;
        IsClipped = clipped;
        SlotX = slotX;
        SlotY = slotY;
        SlotWidth = slotWidth;
        SlotHeight = slotHeight;
    }

    /// <summary>One-based source page number.</summary>
    public int PageNumber { get; }
    /// <summary>Left position on paper in points.</summary>
    public double X { get; }
    /// <summary>Top position on paper in points.</summary>
    public double Y { get; }
    /// <summary>Placed source width in points.</summary>
    public double Width { get; }
    /// <summary>Placed source height in points.</summary>
    public double Height { get; }
    /// <summary>Applied source-to-paper scale.</summary>
    public double Scale { get; }
    /// <summary>Whether the fill mode crops source content at the slot boundary.</summary>
    public bool IsClipped { get; }
    /// <summary>Left edge of the target sheet slot in points.</summary>
    public double SlotX { get; }
    /// <summary>Top edge of the target sheet slot in points.</summary>
    public double SlotY { get; }
    /// <summary>Width of the target sheet slot in points.</summary>
    public double SlotWidth { get; }
    /// <summary>Height of the target sheet slot in points.</summary>
    public double SlotHeight { get; }
}

/// <summary>One paper sheet in a print-preview plan.</summary>
public sealed class PdfPrintSheet {
    internal PdfPrintSheet(int sheetNumber, PageSize paperSize, IReadOnlyList<PdfPrintPlacement> placements) {
        SheetNumber = sheetNumber;
        PaperSize = paperSize;
        Placements = placements.ToArray();
    }

    /// <summary>One-based sheet number.</summary>
    public int SheetNumber { get; }
    /// <summary>Resolved paper size.</summary>
    public PageSize PaperSize { get; }
    /// <summary>Source pages on this sheet.</summary>
    public IReadOnlyList<PdfPrintPlacement> Placements { get; }
}

/// <summary>Deterministic print-preview plan over a source PDF.</summary>
public sealed class PdfPrintPlan {
    internal PdfPrintPlan(int sourcePageCount, IReadOnlyList<int> selectedPages, IReadOnlyList<PdfPrintSheet> sheets) {
        SourcePageCount = sourcePageCount;
        SelectedPages = selectedPages.ToArray();
        Sheets = sheets.ToArray();
    }

    /// <summary>Total source document pages.</summary>
    public int SourcePageCount { get; }
    /// <summary>Selected one-based source pages in print order.</summary>
    public IReadOnlyList<int> SelectedPages { get; }
    /// <summary>Preview sheets.</summary>
    public IReadOnlyList<PdfPrintSheet> Sheets { get; }
}

/// <summary>Creates print-preview sheet geometry without depending on a platform print driver.</summary>
public static class PdfPrintPlanner {
    /// <summary>Creates a validated print-preview plan.</summary>
    public static PdfPrintPlan Create(PdfPrintPlanRequest request) {
        ArgumentNullException.ThrowIfNull(request);
        if (string.IsNullOrWhiteSpace(request.InputPath)) throw new ArgumentException("Input path cannot be empty.", nameof(request));
        if (request.PagesPerSheet is not 1 and not 2 and not 4) throw new ArgumentOutOfRangeException(nameof(request.PagesPerSheet));
        if (request.Margin < 0D || double.IsNaN(request.Margin) || double.IsInfinity(request.Margin)) {
            throw new ArgumentOutOfRangeException(nameof(request.Margin));
        }
        if (!Enum.IsDefined(request.Orientation)) throw new ArgumentOutOfRangeException(nameof(request.Orientation));
        if (!Enum.IsDefined(request.ScaleMode)) throw new ArgumentOutOfRangeException(nameof(request.ScaleMode));

        PdfDocument document = PdfDocument.Load(
            Path.GetFullPath(request.InputPath),
            new PdfLoadOptions { Password = request.PdfPassword });
        PdfDocumentInfo info = document.Inspect();
        int[] pages = string.IsNullOrWhiteSpace(request.Pages)
            ? Enumerable.Range(1, info.PageCount).ToArray()
            : PdfPageSelector.Parse(request.Pages)
                .ResolveSelection(info.PageCount)
                .Ranges
                .SelectMany(static range => Enumerable.Range(range.FirstPage, range.PageCount))
                .ToArray();
        if (pages.Length == 0) throw new ArgumentException("The page selection is empty.", nameof(request));

        var sheets = new List<PdfPrintSheet>((pages.Length + request.PagesPerSheet - 1) / request.PagesPerSheet);
        for (int offset = 0; offset < pages.Length; offset += request.PagesPerSheet) {
            int count = Math.Min(request.PagesPerSheet, pages.Length - offset);
            PdfPageInfo firstPage = info.Pages[pages[offset] - 1];
            PageSize paper = ResolvePaper(request, firstPage.Width > firstPage.Height);
            if (paper.Width <= request.Margin * 2D || paper.Height <= request.Margin * 2D) {
                throw new ArgumentException("Print margins leave no printable paper area.", nameof(request));
            }
            IReadOnlyList<PdfPrintPlacement> placements = CreatePlacements(request, info, pages, offset, count, paper);
            sheets.Add(new PdfPrintSheet(sheets.Count + 1, paper, placements));
        }
        return new PdfPrintPlan(info.PageCount, pages, sheets);
    }

    private static PageSize ResolvePaper(PdfPrintPlanRequest request, bool sourceLandscape) => request.Orientation switch {
        PdfPrintOrientation.Portrait => request.PaperSize.Portrait(),
        PdfPrintOrientation.Landscape => request.PaperSize.Landscape(),
        _ => sourceLandscape ? request.PaperSize.Landscape() : request.PaperSize.Portrait()
    };

    private static IReadOnlyList<PdfPrintPlacement> CreatePlacements(
        PdfPrintPlanRequest request,
        PdfDocumentInfo info,
        IReadOnlyList<int> pages,
        int offset,
        int count,
        PageSize paper) {
        int columns = request.PagesPerSheet == 1 ? 1 : 2;
        int rows = request.PagesPerSheet == 4 ? 2 : 1;
        double printableWidth = paper.Width - request.Margin * 2D;
        double printableHeight = paper.Height - request.Margin * 2D;
        double slotWidth = printableWidth / columns;
        double slotHeight = printableHeight / rows;
        var placements = new List<PdfPrintPlacement>(count);

        for (int index = 0; index < count; index++) {
            int pageNumber = pages[offset + index];
            PdfPageInfo source = info.Pages[pageNumber - 1];
            int column = index % columns;
            int row = index / columns;
            double fitScale = Math.Min(slotWidth / source.Width, slotHeight / source.Height);
            double fillScale = Math.Max(slotWidth / source.Width, slotHeight / source.Height);
            double scale = request.ScaleMode switch {
                PdfPrintScaleMode.ActualSize => Math.Min(1D, fitScale),
                PdfPrintScaleMode.Fill => fillScale,
                _ => fitScale
            };
            double width = source.Width * scale;
            double height = source.Height * scale;
            double slotX = request.Margin + column * slotWidth;
            double slotY = request.Margin + row * slotHeight;
            placements.Add(new PdfPrintPlacement(
                pageNumber,
                slotX + (slotWidth - width) / 2D,
                slotY + (slotHeight - height) / 2D,
                width,
                height,
                scale,
                request.ScaleMode == PdfPrintScaleMode.Fill && (width > slotWidth + 0.01D || height > slotHeight + 0.01D),
                slotX,
                slotY,
                slotWidth,
                slotHeight));
        }
        return placements;
    }
}
