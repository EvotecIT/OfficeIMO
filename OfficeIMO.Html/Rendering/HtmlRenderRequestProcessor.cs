namespace OfficeIMO.Html;

internal static class HtmlRenderRequestProcessor {
    internal static HtmlRenderResult Complete(
        HtmlRenderRequest request,
        HtmlRenderDocument rendered,
        HtmlRenderOptions options,
        CancellationToken cancellationToken) {
        if (request == null) throw new ArgumentNullException(nameof(request));
        if (rendered == null) throw new ArgumentNullException(nameof(rendered));
        if (options == null) throw new ArgumentNullException(nameof(options));
        cancellationToken.ThrowIfCancellationRequested();

        HtmlRenderDocument projected;
        IReadOnlyList<HtmlRenderSurfaceResult> selectedSurfaces;
        if (request.Pagination == HtmlRenderPaginationPolicy.FixedCanvasSlicing) {
            projected = SliceContinuous(rendered, request.PageSet, options, cancellationToken, out selectedSurfaces);
        } else {
            List<HtmlRenderSurfaceResult> surfaces;
            if (request.Surface == HtmlRenderLayoutSurface.Viewport) {
                projected = ClipViewport(rendered, options, out surfaces);
            } else {
                projected = rendered;
                surfaces = rendered.Pages.Select((page, index) => {
                    ValidateSurface(page.Width, page.Height, options);
                    return new HtmlRenderSurfaceResult(index, page.PageNumber, page.Width, page.Height, 0D, 0D, false);
                }).ToList();
            }

            IReadOnlyList<HtmlRenderPage> selectedPages = request.PageSet.Select(projected.Pages);
            selectedSurfaces = SelectSurfaceDescriptors(projected.Pages, surfaces, selectedPages);
            if (selectedPages.Count != projected.Pages.Count || !selectedPages.SequenceEqual(projected.Pages)) {
                projected = projected.Project(selectedPages, projected.Mode);
            }
        }

        if (request.PageSet.Mode == HtmlRenderPageSetMode.Stitched) {
            projected = Stitch(projected, selectedSurfaces, options, cancellationToken, out HtmlRenderSurfaceResult stitched);
            selectedSurfaces = new[] { stitched };
        }
        return new HtmlRenderResult(request, projected, selectedSurfaces);
    }

    private static HtmlRenderDocument ClipViewport(
        HtmlRenderDocument rendered,
        HtmlRenderOptions options,
        out List<HtmlRenderSurfaceResult> surfaces) {
        if (rendered.Pages.Count != 1) {
            throw new InvalidOperationException("Viewport rendering requires one completed continuous surface.");
        }
        HtmlRenderPage source = rendered.Pages[0];
        ValidateSurface(source.Width, source.Height, options);
        var clip = new HtmlRenderClipGroup(
            0D, 0D, source.Width, source.Height, true, true, source.Scene, 0,
            source: "screen viewport");
        var viewport = new HtmlRenderPage(source.PageNumber, source.Width, source.Height,
            new[] { clip }, fonts: source.Fonts);
        surfaces = new List<HtmlRenderSurfaceResult> {
            new(0, source.PageNumber, source.Width, source.Height, 0D, 0D, true)
        };
        return rendered.Project(new[] { viewport }, HtmlRenderMode.Continuous);
    }

    private static HtmlRenderDocument SliceContinuous(
        HtmlRenderDocument rendered,
        HtmlRenderPageSet pageSet,
        HtmlRenderOptions options,
        CancellationToken cancellationToken,
        out IReadOnlyList<HtmlRenderSurfaceResult> surfaces) {
        if (rendered.Pages.Count != 1) {
            throw new InvalidOperationException("Fixed-canvas slicing requires one completed continuous surface.");
        }
        HtmlRenderPage source = rendered.Pages[0];
        double width = options.PageWidth;
        double height = options.PageHeight;
        ValidateSurface(width, height, options);
        int pageCount = Math.Max(1, checked((int)Math.Ceiling(source.Height / height)));
        if (pageCount > options.MaxPageCount) {
            throw new InvalidOperationException(
                $"Fixed-canvas slicing produced {pageCount} pages, exceeding MaxPageCount {options.MaxPageCount}.");
        }

        IReadOnlyList<int> selectedIndices = pageSet.SelectIndices(pageCount);
        var pages = new List<HtmlRenderPage>(selectedIndices.Count);
        var descriptors = new List<HtmlRenderSurfaceResult>(selectedIndices.Count);
        var projector = new HtmlRenderVisualProjector(options.MaxProjectedVisuals, cancellationToken);
        foreach (int sourceIndex in selectedIndices) {
            cancellationToken.ThrowIfCancellationRequested();
            double sourceY = sourceIndex * height;
            IReadOnlyList<HtmlRenderVisual> projected = projector.Project(
                source.Scene, 0D, sourceY, width, height, 0D, 0D);
            var clip = new HtmlRenderClipGroup(0D, 0D, width, height, true, true, projected, 0,
                source: "screen snapshot page " + (sourceIndex + 1).ToString(System.Globalization.CultureInfo.InvariantCulture));
            pages.Add(new HtmlRenderPage(sourceIndex + 1, width, height, new[] { clip }, fonts: source.Fonts));
            var placement = new HtmlRenderSourcePlacement(
                source.PageNumber, 0D, sourceY,
                Math.Min(width, source.Width), Math.Min(height, Math.Max(0.01D, source.Height - sourceY)),
                0D, 0D, true);
            descriptors.Add(new HtmlRenderSurfaceResult(
                descriptors.Count, source.PageNumber, width, height, 0D, sourceY, true, new[] { placement }));
        }
        surfaces = descriptors.AsReadOnly();
        return rendered.Project(pages, HtmlRenderMode.Paged);
    }

    private static IReadOnlyList<HtmlRenderSurfaceResult> SelectSurfaceDescriptors(
        IReadOnlyList<HtmlRenderPage> allPages,
        IReadOnlyList<HtmlRenderSurfaceResult> allSurfaces,
        IReadOnlyList<HtmlRenderPage> selectedPages) {
        var selected = new List<HtmlRenderSurfaceResult>(selectedPages.Count);
        for (int index = 0; index < selectedPages.Count; index++) {
            int sourceIndex = IndexOfReference(allPages, selectedPages[index]);
            HtmlRenderSurfaceResult source = allSurfaces[sourceIndex];
            selected.Add(new HtmlRenderSurfaceResult(index, source.SourcePageNumber, source.Width, source.Height,
                source.SourceOffsetX, source.SourceOffsetY, source.IsClipped, source.SourcePlacements));
        }
        return selected.AsReadOnly();
    }

    private static int IndexOfReference(IReadOnlyList<HtmlRenderPage> pages, HtmlRenderPage selected) {
        for (int index = 0; index < pages.Count; index++) {
            if (ReferenceEquals(pages[index], selected)) return index;
        }
        throw new InvalidOperationException("The selected page does not belong to the render result.");
    }

    private static HtmlRenderDocument Stitch(
        HtmlRenderDocument rendered,
        IReadOnlyList<HtmlRenderSurfaceResult> surfaces,
        HtmlRenderOptions options,
        CancellationToken cancellationToken,
        out HtmlRenderSurfaceResult descriptor) {
        double width = rendered.Pages.Max(page => page.Width);
        double height = rendered.Pages.Sum(page => page.Height);
        ValidateSurface(width, Math.Max(1D, height), options);
        var visuals = new List<HtmlRenderVisual>(rendered.Pages.Count);
        var placements = new List<HtmlRenderSourcePlacement>();
        var projector = new HtmlRenderVisualProjector(options.MaxProjectedVisuals, cancellationToken);
        double y = 0D;
        for (int index = 0; index < rendered.Pages.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            HtmlRenderPage page = rendered.Pages[index];
            IReadOnlyList<HtmlRenderVisual> projected = projector.Project(
                page.Scene, 0D, 0D, page.Width, page.Height, 0D, y);
            visuals.Add(new HtmlRenderClipGroup(0D, y, page.Width, page.Height, true, true,
                projected, index, source: "stitched page " + page.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture)));
            foreach (HtmlRenderSourcePlacement source in surfaces[index].SourcePlacements) {
                placements.Add(new HtmlRenderSourcePlacement(
                    source.SourcePageNumber, source.SourceOffsetX, source.SourceOffsetY,
                    source.Width, source.Height, source.OutputOffsetX, y + source.OutputOffsetY, source.IsClipped));
            }
            y += page.Height;
        }
        var stitched = new HtmlRenderPage(1, width, Math.Max(1D, height), visuals, fonts: rendered.Fonts);
        HtmlRenderSourcePlacement first = placements[0];
        descriptor = new HtmlRenderSurfaceResult(0, first.SourcePageNumber, width, height,
            first.SourceOffsetX, first.SourceOffsetY, true, placements);
        return rendered.Project(new[] { stitched }, HtmlRenderMode.Continuous);
    }

    private static void ValidateSurface(double width, double height, HtmlRenderOptions options) {
        double scale = options.GetEffectiveScale(width, height);
        double pixelWidth = Math.Ceiling(width * scale);
        double pixelHeight = Math.Ceiling(height * scale);
        if (width <= 0D || height <= 0D
            || double.IsNaN(pixelWidth) || double.IsInfinity(pixelWidth)
            || double.IsNaN(pixelHeight) || double.IsInfinity(pixelHeight)
            || pixelWidth > options.MaxSurfaceWidth || pixelHeight > options.MaxSurfaceHeight) {
            throw new InvalidOperationException("HTML render projection exceeded the configured maximum image surface dimensions.");
        }
    }
}
