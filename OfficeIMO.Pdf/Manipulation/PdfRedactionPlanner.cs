namespace OfficeIMO.Pdf;

/// <summary>Builds redaction impact previews without modifying the PDF.</summary>
public static class PdfRedactionPlanner {
    private const double DefaultTextHeight = 12D;

    /// <summary>Plans rectangle-based redaction impact for a PDF byte array.</summary>
    public static PdfRedactionPlan Plan(byte[] pdf, IEnumerable<PdfRedactionArea> areas, PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? options = null) {
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

            return new PdfRedactionPlan(preflight, areaArray, Array.Empty<PdfRedactionMatch>(), findings.AsReadOnly());
        }

        PdfLogicalDocument logical = PdfLogicalDocument.From(PdfReadDocument.Load(pdf, options), layoutOptions);
        PdfDocumentInfo info = preflight.DocumentInfo ?? PdfInspector.Inspect(pdf, options);
        var matches = new List<PdfRedactionMatch>();

        foreach (PdfRedactionArea area in areaArray) {
            AddTextMatches(area, logical.TextBlocks, matches);
            AddAnnotationMatches(area, info.Pages, matches);
        }

        findings.Add(new PdfDiagnosticFinding(
            PdfDiagnosticSeverity.Info,
            "RedactionPlanOnly",
            "This plan reports rectangle intersections only. It does not remove or rewrite PDF content."));

        return new PdfRedactionPlan(preflight, areaArray, matches.AsReadOnly(), findings.AsReadOnly());
    }

    /// <summary>Plans rectangle-based redaction impact for a PDF file.</summary>
    public static PdfRedactionPlan Plan(string path, IEnumerable<PdfRedactionArea> areas, PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return Plan(File.ReadAllBytes(path), areas, layoutOptions, options);
    }

    /// <summary>Plans rectangle-based redaction impact for a readable PDF stream.</summary>
    public static PdfRedactionPlan Plan(Stream stream, IEnumerable<PdfRedactionArea> areas, PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) {
            throw new ArgumentException("Stream must be readable.", nameof(stream));
        }

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return Plan(buffer.ToArray(), areas, layoutOptions, options);
    }

    private static void AddTextMatches(PdfRedactionArea area, IReadOnlyList<PdfLogicalTextBlock> textBlocks, List<PdfRedactionMatch> matches) {
        foreach (PdfLogicalTextBlock block in textBlocks) {
            if (block.PageNumber != area.PageNumber) {
                continue;
            }

            double x = Math.Min(block.XStart, block.XEnd);
            double width = Math.Abs(block.XEnd - block.XStart);
            double y = block.BaselineY - DefaultTextHeight;
            if (!Intersects(area.X, area.Y, area.Width, area.Height, x, y, width, DefaultTextHeight * 1.5D)) {
                continue;
            }

            matches.Add(new PdfRedactionMatch(
                PdfRedactionMatchKind.TextBlock,
                area,
                block.PageNumber,
                x,
                y,
                width,
                DefaultTextHeight * 1.5D,
                block.Text,
                null,
                null));
        }
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

    private static bool Intersects(double ax, double ay, double aw, double ah, double bx, double by, double bw, double bh) {
        return ax < bx + bw &&
            ax + aw > bx &&
            ay < by + bh &&
            ay + ah > by;
    }
}
