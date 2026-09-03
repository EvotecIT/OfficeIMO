namespace OfficeIMO.Pdf;

/// <summary>Builds redaction impact previews without modifying the PDF.</summary>
internal static partial class PdfRedactionPlanner {
    private const double DefaultTextHeight = 12D;

    /// <summary>Plans rectangle-based redaction impact for a PDF byte array.</summary>
    public static PdfRedactionPlan Plan(byte[] pdf, IEnumerable<PdfRedactionArea> areas, PdfTextLayoutOptions? layoutOptions = null, PdfLoadOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(areas, nameof(areas));

        PdfRedactionArea[] areaArray = areas.ToArray();
        if (areaArray.Length == 0) {
            throw new ArgumentException("At least one redaction area is required.", nameof(areas));
        }

        PdfDocumentPreflight preflight = PdfInspector.Preflight(pdf, options);
        var findings = new List<PdfDiagnosticFinding>();
        if (!preflight.CanReadLogicalObjects) {
            foreach (string message in preflight.GetCapabilityDiagnostics(PdfPreflightCapability.ReadLogicalObjects)) {
                findings.Add(new PdfDiagnosticFinding(PdfDiagnosticSeverity.Error, "RedactionPlanBlocked", message));
            }

            return new PdfRedactionPlan(
                preflight,
                areaArray,
                Array.Empty<PdfRedactionMatch>(),
                findings.AsReadOnly(),
                searchCriteria: null,
                PdfRedactionPlan.ComputeSourceSha256(pdf));
        }

        PdfReadDocument readDocument = PdfReadDocument.Open(pdf, options);
        IReadOnlyList<string> pageIdentities = PdfRedactionPlan.CapturePageIdentities(readDocument, areaArray);
        PdfDocumentReadResult logical = PdfDocumentReadResult.From(readDocument, layoutOptions);
        PdfDocumentInfo info = preflight.UncheckedDocumentInfo ?? PdfInspector.Inspect(pdf, options);
        var matches = new List<PdfRedactionMatch>();
        var nestedPathPrimitivesByPage = new Dictionary<int, IReadOnlyList<PdfPageVisualPrimitive>>();

        foreach (PdfRedactionArea area in areaArray) {
            AddTextMatches(area, logical, matches);
            AddImageMatches(area, logical.Images, matches, findings);
            AddAnnotationMatches(area, info.Pages, matches);
            if (area.PageNumber <= readDocument.Pages.Count) {
                PdfReadPage page = readDocument.Pages[area.PageNumber - 1];
                if (!nestedPathPrimitivesByPage.TryGetValue(area.PageNumber, out IReadOnlyList<PdfPageVisualPrimitive>? primitives)) {
                    primitives = page.GetIdentityVisualPrimitives();
                    nestedPathPrimitivesByPage.Add(area.PageNumber, primitives);
                }
                AddNestedPathMatches(area, page, primitives, matches, findings);
            }
        }

        findings.Add(new PdfDiagnosticFinding(
            PdfDiagnosticSeverity.Info,
            "RedactionPlanOnly",
            "This plan reports rectangle intersections only. It does not remove or rewrite PDF content."));

        return new PdfRedactionPlan(
            preflight,
            areaArray,
            matches.AsReadOnly(),
            findings.AsReadOnly(),
            searchCriteria: null,
            PdfRedactionPlan.ComputeSourceSha256(pdf),
            pageIdentities);
    }

    /// <summary>Plans rectangle-based redaction impact for a PDF file.</summary>
    public static PdfRedactionPlan Plan(string path, IEnumerable<PdfRedactionArea> areas, PdfTextLayoutOptions? layoutOptions = null, PdfLoadOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return Plan(File.ReadAllBytes(path), areas, layoutOptions, options);
    }

    /// <summary>Plans rectangle-based redaction impact for a readable PDF stream.</summary>
    public static PdfRedactionPlan Plan(Stream stream, IEnumerable<PdfRedactionArea> areas, PdfTextLayoutOptions? layoutOptions = null, PdfLoadOptions? options = null) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) {
            throw new ArgumentException("Stream must be readable.", nameof(stream));
        }

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return Plan(buffer.ToArray(), areas, layoutOptions, options);
    }

    private static void AddTextMatches(PdfRedactionArea area, PdfDocumentReadResult document, List<PdfRedactionMatch> matches) {
        foreach (PdfLogicalTextBlock block in document.TextBlocks) {
            if (block.PageNumber != area.PageNumber) {
                continue;
            }

            PdfTextSpanBounds bounds = GetTextBlockBounds(block, document.Pages[block.PageNumber - 1]);
            if (!Intersects(area.X, area.Y, area.Width, area.Height, bounds.Left, bounds.Bottom, bounds.Width, bounds.Height)) {
                continue;
            }

            matches.Add(new PdfRedactionMatch(
                PdfRedactionMatchKind.TextBlock,
                area,
                block.PageNumber,
                bounds.Left,
                bounds.Bottom,
                bounds.Width,
                bounds.Height,
                block.Text,
                null,
                null));
        }
    }

    private static double GetEffectiveFontSize(PdfLogicalTextBlock block) {
        return block.FontSize > 0D && !double.IsNaN(block.FontSize) && !double.IsInfinity(block.FontSize)
            ? block.FontSize
            : DefaultTextHeight;
    }

    private static PdfTextSpanBounds GetTextBlockBounds(PdfLogicalTextBlock block, PdfLogicalPage page) {
        if (block.Spans.Count > 0) {
            double left = double.MaxValue;
            double bottom = double.MaxValue;
            double right = double.MinValue;
            double top = double.MinValue;
            for (int index = 0; index < block.Spans.Count; index++) {
                PdfTextSpanBounds spanBounds = PdfTextSpanGeometry.GetAxisAlignedBounds(block.Spans[index]);
                left = Math.Min(left, spanBounds.Left);
                bottom = Math.Min(bottom, spanBounds.Bottom);
                right = Math.Max(right, spanBounds.Right);
                top = Math.Max(top, spanBounds.Top);
            }

            return new PdfTextSpanBounds(left, bottom, right, top);
        }

        if (block.VisualBounds is PdfLogicalVisualBounds visualBounds) {
            PdfVisualBounds userBounds = page.TransformVisualBoundsToUser(
                visualBounds.Left,
                visualBounds.Top,
                visualBounds.Right,
                visualBounds.Bottom);
            return new PdfTextSpanBounds(userBounds.Left, userBounds.Top, userBounds.Right, userBounds.Bottom);
        }

        double fontSize = GetEffectiveFontSize(block);
        double x = Math.Min(block.XStart, block.XEnd);
        double width = Math.Max(1D, Math.Abs(block.XEnd - block.XStart));
        return new PdfTextSpanBounds(x, block.BaselineY - fontSize, x + width, block.BaselineY + fontSize * 0.5D);
    }

    private static void AddAnnotationMatches(PdfRedactionArea area, IReadOnlyList<PdfPageInfo> pages, List<PdfRedactionMatch> matches) {
        foreach (PdfPageInfo page in pages) {
            if (page.PageNumber != area.PageNumber) {
                continue;
            }

            foreach (PdfAnnotation annotation in page.Annotations) {
                if (!Intersects(area.X, area.Y, area.Width, area.Height, annotation.X1, annotation.Y1, annotation.Width, annotation.Height)) {
                    continue;
                }

                matches.Add(new PdfRedactionMatch(
                    PdfRedactionMatchKind.Annotation,
                    area,
                    page.PageNumber,
                    annotation.X1,
                    annotation.Y1,
                    annotation.Width,
                    annotation.Height,
                    annotation.Contents,
                    annotation.Subtype,
                    annotation.ObjectNumber));
            }
        }
    }

    private static void AddImageMatches(PdfRedactionArea area, IReadOnlyList<PdfLogicalImage> images, List<PdfRedactionMatch> matches, List<PdfDiagnosticFinding> findings) {
        foreach (PdfLogicalImage image in images) {
            if (image.PageNumber != area.PageNumber) {
                continue;
            }

            foreach (PdfImagePlacement placement in image.Placements) {
                if (!Intersects(area.X, area.Y, area.Width, area.Height, placement.X, placement.Y, placement.Width, placement.Height)) {
                    continue;
                }

                matches.Add(new PdfRedactionMatch(
                    PdfRedactionMatchKind.ImagePlacement,
                    area,
                    placement.PageNumber,
                    placement.X,
                    placement.Y,
                    placement.Width,
                    placement.Height,
                    null,
                    null,
                    placement.ObjectNumber == 0 ? null : placement.ObjectNumber,
                    placement.ResourceName,
                    placement));

                findings.Add(new PdfDiagnosticFinding(
                    PdfDiagnosticSeverity.Warning,
                    "RedactionPlanImageIntersection",
                    "Redaction area intersects an image placement. Applying the plan rewrites supported image pixels and otherwise follows the configured fail-closed, whole-placement removal, or explicit visual-overlay policy.",
                    placement.ObjectNumber == 0 ? null : placement.ObjectNumber,
                    placement.PageNumber));
            }
        }
    }

    private static void AddNestedPathMatches(
        PdfRedactionArea area,
        PdfReadPage page,
        IReadOnlyList<PdfPageVisualPrimitive> primitives,
        List<PdfRedactionMatch> matches,
        List<PdfDiagnosticFinding> findings) {
        PdfVisualBounds visualArea = page.TransformBoundsToVisual(area.X, area.Y, area.Right, area.Top);
        for (int i = 0; i < primitives.Count; i++) {
            PdfPageVisualPrimitive primitive = primitives[i];
            if (primitive.ContentOrderKey?.Depth <= 1) continue;
            double strokePadding = primitive.HasStrokePaint ? Math.Max(0D, primitive.StrokeWidth) / 2D : 0D;
            double left = primitive.X - strokePadding;
            double top = primitive.Y - strokePadding;
            double right = primitive.X + primitive.Width + strokePadding;
            double bottom = primitive.Y + primitive.Height + strokePadding;
            if (!Intersects(
                    visualArea.Left,
                    visualArea.Top,
                    visualArea.Width,
                    visualArea.Height,
                    left,
                    top,
                    right - left,
                    bottom - top)) {
                continue;
            }

            PdfVisualBounds userBounds = page.TransformVisualBoundsToUser(left, top, right, bottom);
            matches.Add(new PdfRedactionMatch(
                PdfRedactionMatchKind.VectorPath,
                area,
                area.PageNumber,
                userBounds.Left,
                userBounds.Top,
                userBounds.Width,
                userBounds.Height,
                null,
                primitive.Kind.ToString(),
                null));
            findings.Add(new PdfDiagnosticFinding(
                PdfDiagnosticSeverity.Warning,
                "RedactionPlanNestedVectorIntersection",
                "The redaction area intersects vector content inside a nested Form XObject. The current writer cannot remove that path; applied-plan verification will remain unverified while the nested vector content remains.",
                pageNumber: area.PageNumber));
        }
    }

    private static bool Intersects(double ax, double ay, double aw, double ah, double bx, double by, double bw, double bh) {
        return ax < bx + bw &&
            ax + aw > bx &&
            ay < by + bh &&
            ay + ah > by;
    }
}
