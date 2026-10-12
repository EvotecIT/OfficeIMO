using OfficeIMO.Pdf;
using OfficeIMO.Internal;

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
    Fill,
    /// <summary>Apply the explicitly requested percentage of the source page's physical size.</summary>
    Custom
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

    /// <summary>Source pages placed on each paper sheet: 1, 2, 4, 6, or 9, in row-major order.</summary>
    public int PagesPerSheet { get; set; } = 1;

    /// <summary>Uniform printable margin in points.</summary>
    public double Margin { get; set; } = 18D;

    /// <summary>Left margin in points, or null to use <see cref="Margin"/>.</summary>
    public double? MarginLeft { get; set; }
    /// <summary>Top margin in points, or null to use <see cref="Margin"/>.</summary>
    public double? MarginTop { get; set; }
    /// <summary>Right margin in points, or null to use <see cref="Margin"/>.</summary>
    public double? MarginRight { get; set; }
    /// <summary>Bottom margin in points, or null to use <see cref="Margin"/>.</summary>
    public double? MarginBottom { get; set; }

    /// <summary>Source-page scaling behavior.</summary>
    public PdfPrintScaleMode ScaleMode { get; set; } = PdfPrintScaleMode.Fit;

    /// <summary>Source-size percentage, from 1 to 1000, used only by <see cref="PdfPrintScaleMode.Custom"/>. Overflow is clipped to the slot.</summary>
    public double CustomScalePercent { get; set; } = 100D;
    /// <summary>Source-page alignment inside each sheet slot.</summary>
    public PdfPrintAlignment Alignment { get; set; } = PdfPrintAlignment.Center;
    /// <summary>Filter by original document page number after resolving <see cref="Pages"/>, preserving selection order.</summary>
    public PdfPrintPageSubset PageSubset { get; set; }
    /// <summary>Color treatment applied to the prepared sheets used by preview and delivery.</summary>
    public PdfPrintColorMode ColorMode { get; set; }

    /// <summary>Optional source password.</summary>
    public string? PdfPassword { get; set; }

    /// <summary>Controls enforcement of authenticated user-password printing restrictions.</summary>
    public PdfPermissionPolicy PermissionPolicy { get; set; } = PdfPermissionPolicy.Enforce;
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
    /// <summary>Whether source content exceeds the slot boundary and is cropped.</summary>
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
    internal PdfPrintPlan(int sourcePageCount, IReadOnlyList<int> selectedPages, IReadOnlyList<PdfPrintSheet> sheets,
        PdfPrintColorMode colorMode = PdfPrintColorMode.Color) {
        SourcePageCount = sourcePageCount;
        SelectedPages = selectedPages.ToArray();
        Sheets = sheets.ToArray();
        ColorMode = colorMode;
    }

    /// <summary>Total source document pages.</summary>
    public int SourcePageCount { get; }
    /// <summary>Selected one-based source pages in print order.</summary>
    public IReadOnlyList<int> SelectedPages { get; }
    /// <summary>Preview sheets.</summary>
    public IReadOnlyList<PdfPrintSheet> Sheets { get; }
    /// <summary>Color treatment for the exact prepared preview and delivery pixels.</summary>
    public PdfPrintColorMode ColorMode { get; }
}

/// <summary>Creates print-preview sheet geometry without depending on a platform print driver.</summary>
public static class PdfPrintPlanner {
    /// <summary>Plans sheets from an already opened document without reading the input location again.</summary>
    /// <remarks>The document's authenticated permission policy governs printing. The request input path identifies
    /// the original source and may be a provider URI; its password and permission policy do not reopen the document.</remarks>
    public static PdfPrintPlan Create(PdfDocument document, PdfPrintPlanRequest request, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(document);
        return CreateAsync(request, (_, _, _) => Task.FromResult(document), cancellationToken).GetAwaiter().GetResult();
    }

    /// <summary>Creates a validated print-preview plan.</summary>
    public static PdfPrintPlan Create(
        PdfPrintPlanRequest request,
        CancellationToken cancellationToken = default) =>
        CreateAsync(request, cancellationToken).GetAwaiter().GetResult();

    /// <summary>Asynchronously creates a validated print-preview plan.</summary>
    public static async Task<PdfPrintPlan> CreateAsync(
        PdfPrintPlanRequest request,
        CancellationToken cancellationToken = default) {
        return await CreateAsync(
            request,
            static (path, options, token) => PdfDocument.LoadAsync(OfficeStorageIdentity.GetLocalPath(path)
                ?? throw new ArgumentException("Provider input requires an already opened document."), options, token),
            cancellationToken).ConfigureAwait(false);
    }

    internal static async Task<PdfPrintPlan> CreateAsync(
        PdfPrintPlanRequest request,
        Func<string, PdfLoadOptions, CancellationToken, Task<PdfDocument>> loadDocument,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(request);
        ArgumentNullException.ThrowIfNull(loadDocument);
        cancellationToken.ThrowIfCancellationRequested();
        if (string.IsNullOrWhiteSpace(request.InputPath)) throw new ArgumentException("Input path cannot be empty.", nameof(request));
        if (request.PagesPerSheet is not 1 and not 2 and not 4 and not 6 and not 9) throw new ArgumentOutOfRangeException(nameof(request.PagesPerSheet));
        ValidateMargin(request.Margin, nameof(request.Margin));
        double left = request.MarginLeft ?? request.Margin, top = request.MarginTop ?? request.Margin;
        double right = request.MarginRight ?? request.Margin, bottom = request.MarginBottom ?? request.Margin;
        ValidateMargin(left, nameof(request.MarginLeft)); ValidateMargin(top, nameof(request.MarginTop));
        ValidateMargin(right, nameof(request.MarginRight)); ValidateMargin(bottom, nameof(request.MarginBottom));
        if (!double.IsFinite(request.CustomScalePercent) || request.CustomScalePercent is < 1 or > 1000)
            throw new ArgumentOutOfRangeException(nameof(request.CustomScalePercent));
        if (!Enum.IsDefined(request.Alignment)) throw new ArgumentOutOfRangeException(nameof(request.Alignment));
        if (!Enum.IsDefined(request.PageSubset)) throw new ArgumentOutOfRangeException(nameof(request.PageSubset));
        if (!Enum.IsDefined(request.ColorMode)) throw new ArgumentOutOfRangeException(nameof(request.ColorMode));
        if (!Enum.IsDefined(request.Orientation)) throw new ArgumentOutOfRangeException(nameof(request.Orientation));
        if (!Enum.IsDefined(request.ScaleMode)) throw new ArgumentOutOfRangeException(nameof(request.ScaleMode));
        if (!Enum.IsDefined(request.PermissionPolicy)) throw new ArgumentOutOfRangeException(nameof(request.PermissionPolicy));

        var validated = new ValidatedPrintPlanRequest(
            OfficeStorageIdentity.Normalize(request.InputPath),
            request.Pages,
            request.PaperSize,
            request.Orientation,
            request.PagesPerSheet,
            left, top, right, bottom,
            request.ScaleMode,
            request.CustomScalePercent, request.Alignment, request.PageSubset, request.ColorMode,
            request.PdfPassword,
            request.PermissionPolicy);

        PdfDocument document = await loadDocument(
            validated.InputPath,
            new PdfLoadOptions {
                Password = validated.PdfPassword,
                PermissionPolicy = validated.PermissionPolicy
            },
            cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        IReadOnlyList<PdfPageLayoutInfo> pageLayouts = document.GetPrintablePageLayouts(options: null, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        int[] pages = string.IsNullOrWhiteSpace(validated.Pages)
            ? Enumerable.Range(1, pageLayouts.Count).ToArray()
            : PdfPageSelector.Parse(validated.Pages)
                .ResolveSelection(pageLayouts.Count)
                .Ranges
                .SelectMany(static range => Enumerable.Range(range.FirstPage, range.PageCount))
                .ToArray();
        cancellationToken.ThrowIfCancellationRequested();
        if (validated.PageSubset != PdfPrintPageSubset.All) {
            int parity = validated.PageSubset == PdfPrintPageSubset.Odd ? 1 : 0;
            pages = pages.Where(page => page % 2 == parity).ToArray();
        }
        cancellationToken.ThrowIfCancellationRequested();
        if (pages.Length == 0) throw new ArgumentException("The page selection and odd/even filter contain no pages.", nameof(request));

        var sheets = new List<PdfPrintSheet>((pages.Length + validated.PagesPerSheet - 1) / validated.PagesPerSheet);
        for (int offset = 0; offset < pages.Length; offset += validated.PagesPerSheet) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = Math.Min(validated.PagesPerSheet, pages.Length - offset);
            PdfPageLayoutInfo firstPage = pageLayouts[pages[offset] - 1];
            (double firstPageWidth, double firstPageHeight) = GetVisualPageSize(firstPage);
            PageSize paper = ResolvePaper(validated, firstPageWidth > firstPageHeight);
            if (paper.Width <= validated.MarginLeft + validated.MarginRight || paper.Height <= validated.MarginTop + validated.MarginBottom) {
                throw new ArgumentException("Print margins leave no printable paper area.", nameof(request));
            }
            IReadOnlyList<PdfPrintPlacement> placements = CreatePlacements(
                validated,
                pageLayouts,
                pages,
                offset,
                count,
                paper,
                cancellationToken);
            sheets.Add(new PdfPrintSheet(sheets.Count + 1, paper, placements));
        }
        return new PdfPrintPlan(pageLayouts.Count, pages, sheets, validated.ColorMode);
    }

    private static PageSize ResolvePaper(ValidatedPrintPlanRequest request, bool sourceLandscape) => request.Orientation switch {
        PdfPrintOrientation.Portrait => request.PaperSize.Portrait(),
        PdfPrintOrientation.Landscape => request.PaperSize.Landscape(),
        _ => sourceLandscape ? request.PaperSize.Landscape() : request.PaperSize.Portrait()
    };

    private static IReadOnlyList<PdfPrintPlacement> CreatePlacements(
        ValidatedPrintPlanRequest request,
        IReadOnlyList<PdfPageLayoutInfo> pageLayouts,
        IReadOnlyList<int> pages,
        int offset,
        int count,
        PageSize paper,
        CancellationToken cancellationToken) {
        (int columns, int rows) = request.PagesPerSheet switch {
            1 => (1, 1), 2 => (2, 1), 4 => (2, 2),
            6 when paper.Width > paper.Height => (3, 2),
            6 => (2, 3), _ => (3, 3)
        };
        double printableWidth = paper.Width - request.MarginLeft - request.MarginRight;
        double printableHeight = paper.Height - request.MarginTop - request.MarginBottom;
        double slotWidth = printableWidth / columns;
        double slotHeight = printableHeight / rows;
        var placements = new List<PdfPrintPlacement>(count);

        for (int index = 0; index < count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            int pageNumber = pages[offset + index];
            PdfPageLayoutInfo source = pageLayouts[pageNumber - 1];
            (double sourceWidth, double sourceHeight) = GetVisualPageSize(source);
            int column = index % columns;
            int row = index / columns;
            double fitScale = Math.Min(slotWidth / sourceWidth, slotHeight / sourceHeight);
            double fillScale = Math.Max(slotWidth / sourceWidth, slotHeight / sourceHeight);
            double scale = request.ScaleMode switch {
                PdfPrintScaleMode.ActualSize => Math.Min(1D, fitScale),
                PdfPrintScaleMode.Fill => fillScale,
                PdfPrintScaleMode.Custom => request.CustomScalePercent / 100D,
                _ => fitScale
            };
            double width = sourceWidth * scale;
            double height = sourceHeight * scale;
            double slotX = request.MarginLeft + column * slotWidth;
            double slotY = request.MarginTop + row * slotHeight;
            (double horizontal, double vertical) = AlignmentFactors(request.Alignment);
            placements.Add(new PdfPrintPlacement(
                pageNumber,
                slotX + (slotWidth - width) * horizontal,
                slotY + (slotHeight - height) * vertical,
                width,
                height,
                scale,
                width > slotWidth + 0.01D || height > slotHeight + 0.01D,
                slotX,
                slotY,
                slotWidth,
                slotHeight));
        }
        return placements;
    }

    private static (double Width, double Height) GetVisualPageSize(PdfPageLayoutInfo page) =>
        (page.VisualWidth, page.VisualHeight);

    private static void ValidateMargin(double value, string name) {
        if (!double.IsFinite(value) || value < 0D) throw new ArgumentOutOfRangeException(name);
    }

    private static (double Horizontal, double Vertical) AlignmentFactors(PdfPrintAlignment alignment) => alignment switch {
        PdfPrintAlignment.TopLeft => (0, 0), PdfPrintAlignment.Top => (0.5, 0), PdfPrintAlignment.TopRight => (1, 0),
        PdfPrintAlignment.Left => (0, 0.5), PdfPrintAlignment.Right => (1, 0.5),
        PdfPrintAlignment.BottomLeft => (0, 1), PdfPrintAlignment.Bottom => (0.5, 1), PdfPrintAlignment.BottomRight => (1, 1),
        _ => (0.5, 0.5)
    };

    private sealed record ValidatedPrintPlanRequest(
        string InputPath,
        string? Pages,
        PageSize PaperSize,
        PdfPrintOrientation Orientation,
        int PagesPerSheet,
        double MarginLeft, double MarginTop, double MarginRight, double MarginBottom,
        PdfPrintScaleMode ScaleMode,
        double CustomScalePercent, PdfPrintAlignment Alignment, PdfPrintPageSubset PageSubset, PdfPrintColorMode ColorMode,
        string? PdfPassword,
        PdfPermissionPolicy PermissionPolicy);
}
