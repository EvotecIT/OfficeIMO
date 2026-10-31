using OfficeIMO.Drawing;
using System.Threading;
using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.Pdf.PowerPoint;

public static partial class PdfPowerPointConverterExtensions {
    private static PdfPowerPointConversionResult ImportEditableContent(
        PdfCore.PdfDocument document,
        PowerPointPdfImportOptions options,
        CancellationToken cancellationToken) {
        PdfCore.PdfDocumentReadResult logical = ReadBoundedLogicalDocument(document, options, cancellationToken);
        return ImportEditableContent(logical, options, document, cancellationToken);
    }

    private static PdfPowerPointConversionResult ImportEditableContent(
        PdfCore.PdfDocumentReadResult logical,
        PowerPointPdfImportOptions options,
        CancellationToken cancellationToken) =>
        ImportEditableContent(logical, options, sourceDocument: null, cancellationToken);

    private static PdfPowerPointConversionResult ImportEditableContent(
        PdfCore.PdfDocumentReadResult logical,
        PowerPointPdfImportOptions options,
        PdfCore.PdfDocument? sourceDocument,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (options.MaxEditableObjectsPerPage <= 0) {
            throw new ArgumentOutOfRangeException(
                nameof(options.MaxEditableObjectsPerPage),
                "The editable object limit must be positive.");
        }

        using var presentationOwner = new PresentationConstructionScope();
        PptCore.PowerPointDocument presentation = presentationOwner.Value;
        cancellationToken.ThrowIfCancellationRequested();
        OfficeDrawing? referenceDrawing = logical.Pages.Count == 0 || sourceDocument == null
            ? null
            : sourceDocument.Render.Drawing(logical.Pages[0].PageNumber);
        cancellationToken.ThrowIfCancellationRequested();
        if (referenceDrawing != null) {
            ConfigureEditableSlideSize(presentation, referenceDrawing);
        } else if (logical.Pages.Count > 0) {
            (double width, double height) = logical.Pages[0].GetVisualPageSize();
            ConfigureEditableSlideSize(presentation, width, height);
        }

        Dictionary<int, IReadOnlyList<PdfCore.PdfLogicalTableExtraction>> tablesByPage =
            PdfCore.PdfLogicalTableAnalysis.ExtractTables(logical, options.MaxRows, cancellationToken)
                .GroupBy(static extraction => extraction.PageIndex)
                .ToDictionary(
                    static group => group.Key,
                    static group => (IReadOnlyList<PdfCore.PdfLogicalTableExtraction>)group.ToArray());
        var editablePages = new List<PdfPowerPointEditablePageEntry>(logical.Pages.Count);
        var tableEntries = new List<PdfPowerPointTableImportEntry>();
        var warnings = new List<PdfCore.PdfConversionWarning>();

        for (int pageIndex = 0; pageIndex < logical.Pages.Count; pageIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfCore.PdfLogicalPage page = logical.Pages[pageIndex];
            OfficeDrawing? drawing = sourceDocument == null
                ? null
                : pageIndex == 0 && referenceDrawing != null
                    ? referenceDrawing
                    : sourceDocument.Render.Drawing(page.PageNumber);
            cancellationToken.ThrowIfCancellationRequested();
            int slideIndex = presentation.Slides.Count;
            PptCore.PowerPointSlide slide = presentation.AddSlide();
            EditablePagePlacement placement = GetEditablePagePlacement(
                drawing?.Width ?? page.GetVisualPageSize().Width,
                drawing?.Height ?? page.GetVisualPageSize().Height,
                presentation.SlideSize.Width.ToPoints(),
                presentation.SlideSize.Height.ToPoints());
            IReadOnlyList<PdfCore.PdfLogicalTableExtraction> tables =
                tablesByPage.TryGetValue(pageIndex, out IReadOnlyList<PdfCore.PdfLogicalTableExtraction>? pageTables)
                    ? pageTables
                    : Array.Empty<PdfCore.PdfLogicalTableExtraction>();
            Dictionary<PdfCore.PdfLogicalTable, EditableBounds> tableBoundsByTable = GetEditableTableBounds(page, tables, placement, drawing, cancellationToken);
            List<EditableBounds> tableBounds = tableBoundsByTable.Values.ToList();
            int remainingObjects = options.MaxEditableObjectsPerPage;

            EditableImportCount shapeImport = drawing == null
                ? new EditableImportCount(
                    0,
                    page.VectorPrimitiveCount,
                    false)
                : ImportEditableShapes(drawing, slide, placement, tableBounds, remainingObjects, cancellationToken);
            remainingObjects -= shapeImport.Imported;
            EditableImportCount imageImport = ImportEditableImages(
                page,
                slide,
                placement,
                remainingObjects,
                warnings,
                cancellationToken);
            remainingObjects -= imageImport.Imported;
            EditableImportCount textImport = ImportEditableTextBlocks(
                page,
                slide,
                placement,
                tables,
                remainingObjects,
                cancellationToken);
            remainingObjects -= textImport.Imported;
            EditableTableImportCount tableImport = ImportEditableTables(
                pageIndex,
                page,
                tables,
                presentation,
                slide,
                slideIndex,
                placement,
                tableBoundsByTable,
                options,
                tableEntries,
                remainingObjects,
                cancellationToken);

            int omittedVectors = Math.Max(
                shapeImport.Omitted,
                Math.Max(0, page.UnrepresentedVectorPrimitiveCount - shapeImport.Imported));
            editablePages.Add(new PdfPowerPointEditablePageEntry(
                page.PageNumber,
                slideIndex,
                textImport.Imported,
                tableImport.PrimarySlideCount,
                shapeImport.Imported,
                imageImport.Imported,
                textImport.Omitted,
                tableImport.Omitted,
                omittedVectors,
                imageImport.Omitted));
            if (page.OmittedArtifactTextSpanCount > 0) {
                warnings.Add(new PdfCore.PdfConversionWarning(
                    "OfficeIMO.Pdf.PowerPoint",
                    "pdf-artifact-text-omitted",
                    "PDF page " + page.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    "Text marked as a PDF artifact (running headers, footers or page numbers) was not imported. Set ReadOptions.IncludeArtifactText to keep it.",
                    PdfCore.PdfConversionWarningSeverity.Warning,
                    OfficeConversionLossKind.Omission,
                    details: new Dictionary<string, string> {
                        ["pageNumber"] = page.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture),
                        ["construct"] = "Artifact text",
                        ["Count"] = page.OmittedArtifactTextSpanCount.ToString(System.Globalization.CultureInfo.InvariantCulture),
                        ["Disposition"] = "Omitted"
                    }));
            }
            AddEditablePageWarnings(
                warnings,
                page.PageNumber,
                textImport.Omitted,
                omittedVectors,
                imageImport.Omitted,
                shapeImport.LimitReached || imageImport.LimitReached || textImport.LimitReached || tableImport.LimitReached);
            if (sourceDocument != null) {
                cancellationToken.ThrowIfCancellationRequested();
                AddEditableRendererWarnings(
                    warnings,
                    page.PageNumber,
                    sourceDocument.Render.CapabilityDiagnostics(page.PageNumber),
                    cancellationToken);
            }
        }

        if (logical.Pages.Count == 0) {
            PptCore.PowerPointSlide slide = presentation.AddSlide();
            slide.AddTitle("PDF");
            slide.AddTextBox("No PDF pages were selected.");
        }

        cancellationToken.ThrowIfCancellationRequested();
        PdfCore.PdfTableExtractionScopeReport scope = PdfCore.PdfLogicalTableAnalysis.AnalyzeExtractionScope(logical);
        AddEditableDocumentWarnings(warnings, scope);
        return presentationOwner.Release(new PdfPowerPointConversionResult(
            presentation,
            new PdfPowerPointConversionReport(editablePages, tableEntries, scope, warnings)));
    }

    private static void ConfigureEditableSlideSize(
        PptCore.PowerPointDocument presentation,
        OfficeDrawing? drawing) {
        if (drawing == null || drawing.Width <= 0D || drawing.Height <= 0D) return;
        ConfigureEditableSlideSize(presentation, drawing.Width, drawing.Height);
    }

    private static void ConfigureEditableSlideSize(
        PptCore.PowerPointDocument presentation,
        double sourceWidth,
        double sourceHeight) {
        if (sourceWidth <= 0D || sourceHeight <= 0D) return;
        const double maximumSlideDimensionPoints = 720D;
        double width = sourceWidth >= sourceHeight
            ? maximumSlideDimensionPoints
            : maximumSlideDimensionPoints * sourceWidth / sourceHeight;
        double height = sourceHeight >= sourceWidth
            ? maximumSlideDimensionPoints
            : maximumSlideDimensionPoints * sourceHeight / sourceWidth;
        presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(width),OfficeIMO.OfficeLength.Points(height));
    }

    private static EditablePagePlacement GetEditablePagePlacement(
        double pageWidth,
        double pageHeight,
        double slideWidth,
        double slideHeight) {
        double scale = Math.Min(
            slideWidth / Math.Max(1D, pageWidth),
            slideHeight / Math.Max(1D, pageHeight));
        double width = pageWidth * scale;
        double height = pageHeight * scale;
        return new EditablePagePlacement(
            (slideWidth - width) / 2D,
            (slideHeight - height) / 2D,
            scale);
    }

    private static EditableImportCount ImportEditableShapes(
        OfficeDrawing drawing,
        PptCore.PowerPointSlide slide,
        EditablePagePlacement placement,
        IReadOnlyList<EditableBounds> tableBounds,
        int limit,
        CancellationToken cancellationToken) {
        int imported = 0;
        int omitted = 0;
        bool limitReached = false;
        foreach (OfficeDrawingShape drawingShape in drawing.Elements.OfType<OfficeDrawingShape>()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (imported >= limit) {
                omitted++;
                limitReached = true;
                continue;
            }
            // Table bounds are in slide coordinates; compare the shape after the same page-to-slide mapping.
            EditableBounds slideBounds = placement.Map(
                drawingShape.X,
                drawingShape.Y,
                drawingShape.Shape.Width,
                drawingShape.Shape.Height);
            if (tableBounds.Any(bounds => bounds.Inflate(2D * placement.Scale).ContainsCenterOf(slideBounds) ||
                    bounds.OverlapFractionOf(slideBounds) >= 0.6D)) {
                // Cell fills and rules are carried by the native table.
                continue;
            }
            if (!TryAddEditableShape(slide, drawingShape, placement)) {
                omitted++;
                continue;
            }
            imported++;
        }
        return new EditableImportCount(imported, omitted, limitReached);
    }

    private static bool TryAddEditableShape(
        PptCore.PowerPointSlide slide,
        OfficeDrawingShape drawingShape,
        EditablePagePlacement placement) {
        OfficeShape source = drawingShape.Shape;
        if (source.Transform.HasValue && source.Transform.Value != OfficeTransform.Identity ||
            source.ClipPath != null ||
            source.FillGradient != null ||
            source.FillRadialGradient != null ||
            source.StrokeGradient != null ||
            source.StrokeRadialGradient != null) {
            return false;
        }

        double left = placement.MapX(drawingShape.X);
        double top = placement.MapY(drawingShape.Y);
        double width = Math.Max(0.01D, source.Width * placement.Scale);
        double height = Math.Max(0.01D, source.Height * placement.Scale);
        PptCore.PowerPointAutoShape target;
        switch (source.Kind) {
            case OfficeShapeKind.Rectangle:
                target = slide.AddShape(OfficePresetShapeType.Rectangle, OfficeIMO.OfficeLength.Points(left), OfficeIMO.OfficeLength.Points(top), OfficeIMO.OfficeLength.Points(width), OfficeIMO.OfficeLength.Points(height));
                break;
            case OfficeShapeKind.RoundedRectangle:
                target = slide.AddShape(OfficePresetShapeType.RoundRectangle, OfficeIMO.OfficeLength.Points(left), OfficeIMO.OfficeLength.Points(top), OfficeIMO.OfficeLength.Points(width), OfficeIMO.OfficeLength.Points(height));
                break;
            case OfficeShapeKind.Ellipse:
                target = slide.AddShape(OfficePresetShapeType.Ellipse, OfficeIMO.OfficeLength.Points(left), OfficeIMO.OfficeLength.Points(top), OfficeIMO.OfficeLength.Points(width), OfficeIMO.OfficeLength.Points(height));
                break;
            case OfficeShapeKind.Line when source.Points.Count >= 2:
                OfficePoint start = source.Points[0];
                OfficePoint end = source.Points[source.Points.Count - 1];
                target = slide.AddLine(
                    OfficeIMO.OfficeLength.Points(placement.MapX(drawingShape.X + start.X)),
                    OfficeIMO.OfficeLength.Points(placement.MapY(drawingShape.Y + start.Y)),
                    OfficeIMO.OfficeLength.Points(placement.MapX(drawingShape.X + end.X)),
                    OfficeIMO.OfficeLength.Points(placement.MapY(drawingShape.Y + end.Y)));
                break;
            default:
                return false;
        }

        ApplyEditableShapeStyle(target, source);
        return true;
    }

    private static void ApplyEditableShapeStyle(PptCore.PowerPointShape target, OfficeShape source) {
        if (source.Kind != OfficeShapeKind.Line) {
            target.FillColor = ToHex(source.FillColor ?? OfficeColor.White);
            target.FillTransparency = source.FillColor.HasValue
                ? ToTransparency(source.FillOpacity)
                : 100;
        }
        target.OutlineColor = ToHex(source.StrokeColor ?? OfficeColor.White);
        target.OutlineTransparency = source.StrokeColor.HasValue
            ? ToTransparency(source.StrokeOpacity)
            : 100;
        target.OutlineWidth = global::OfficeIMO.OfficeLength.Points(Math.Max(0D, source.StrokeWidth));
        target.OutlineDash = source.StrokeDashStyle switch {
            OfficeStrokeDashStyle.Dash => PptCore.PowerPointLineDashStyle.Dash,
            OfficeStrokeDashStyle.Dot => PptCore.PowerPointLineDashStyle.Dot,
            OfficeStrokeDashStyle.DashDot => PptCore.PowerPointLineDashStyle.DashDot,
            OfficeStrokeDashStyle.DashDotDot => PptCore.PowerPointLineDashStyle.LargeDashDotDot,
            _ => PptCore.PowerPointLineDashStyle.Solid
        };
    }

    private static EditableImportCount ImportEditableImages(
        PdfCore.PdfLogicalPage page,
        PptCore.PowerPointSlide slide,
        EditablePagePlacement placement,
        int limit,
        ICollection<PdfCore.PdfConversionWarning> warnings,
        CancellationToken cancellationToken) {
        int imported = 0;
        int omitted = 0;
        bool limitReached = false;
        for (int imageIndex = 0; imageIndex < page.Images.Count; imageIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfCore.PdfLogicalImage image = page.Images[imageIndex];
            OfficeImageFormat format = ResolveImageFormat(image.SourceImage);
            bool payloadSupported = image.SourceImage.IsImageFile &&
                !image.SourceImage.IsImageMask &&
                format != OfficeImageFormat.Unknown;
            bool unsupportedPayloadCounted = false;
            if (image.Placements.Count == 0) {
                AddEditableImagePlacementWarning(
                    warnings,
                    image,
                    PdfCore.PdfImagePlacementImportPolicy.Analyze(page, image, placement: null));
                continue;
            }
            for (int placementIndex = 0; placementIndex < image.Placements.Count; placementIndex++) {
                cancellationToken.ThrowIfCancellationRequested();
                PdfCore.PdfImagePlacement sourcePlacement = image.Placements[placementIndex];
                PdfCore.PdfImagePlacementImportAssessment assessment =
                    PdfCore.PdfImagePlacementImportPolicy.Analyze(page, image, sourcePlacement);
                if (!assessment.CanImport) {
                    AddEditableImagePlacementWarning(warnings, image, assessment);
                }
                if (assessment.IsSuppressed) {
                    continue;
                }
                if (!payloadSupported) {
                    if (!unsupportedPayloadCounted) {
                        omitted++;
                        unsupportedPayloadCounted = true;
                    }
                    continue;
                }
                if (!assessment.CanImport || !sourcePlacement.IsAxisAligned || imported >= limit) {
                    omitted++;
                    limitReached |= imported >= limit;
                    continue;
                }
                PdfCore.PdfVisualBounds visual = page.TransformBoundsToVisual(
                    sourcePlacement.X,
                    sourcePlacement.Y,
                    sourcePlacement.X + sourcePlacement.Width,
                    sourcePlacement.Y + sourcePlacement.Height);
                EditableBounds bounds = placement.Map(visual.Left, visual.Top, visual.Width, visual.Height);
                double pageRotation = NormalizeRotation(page.RotationDegrees);
                if (pageRotation is 90D or 270D) {
                    bounds = bounds.SwapDimensionsAroundCenter();
                }
                using var stream = new MemoryStream(image.SourceImage.Bytes, writable: false);
                PptCore.PowerPointPicture picture = slide.AddImage(
                    stream,
                    format,
                    OfficeIMO.OfficeLength.Points(bounds.Left),
                    OfficeIMO.OfficeLength.Points(bounds.Top),
                    OfficeIMO.OfficeLength.Points(bounds.Width),
                    OfficeIMO.OfficeLength.Points(bounds.Height));
                // Preserve the PDF frame independently of the image's intrinsic pixel ratio.
                picture.SetSize(OfficeLength.Points(bounds.Width), OfficeLength.Points(bounds.Height));
                if (pageRotation > 0.01D) {
                    picture.Rotation = pageRotation;
                }
                picture.HorizontalFlip = sourcePlacement.A < 0D;
                picture.VerticalFlip = sourcePlacement.D < 0D;
                if (assessment.HasNonDefaultOpacity) {
                    picture.FillTransparency = assessment.MappedTransparencyPercent;
                }
                AddEditableImagePlacementWarning(warnings, image, assessment);
                imported++;
            }
        }
        return new EditableImportCount(imported, omitted, limitReached);
    }

    private static void AddEditableImagePlacementWarning(
        ICollection<PdfCore.PdfConversionWarning> warnings,
        PdfCore.PdfLogicalImage image,
        PdfCore.PdfImagePlacementImportAssessment assessment) {
        string source = "Page " + image.PageNumber.ToString(CultureInfo.InvariantCulture) + "/Image";
        var details = new Dictionary<string, string> {
            ["ResourceName"] = image.ResourceName,
            ["Opacity"] = assessment.Opacity.ToString("R", CultureInfo.InvariantCulture),
            ["BlendMode"] = assessment.BlendMode.ToString()
        };
        switch (assessment.Disposition) {
            case PdfCore.PdfImagePlacementImportDisposition.SuppressInvisible:
                warnings.Add(new PdfCore.PdfConversionWarning(
                    "OfficeIMO.Pdf.PowerPoint",
                    "pdf-invisible-image-placement-suppressed",
                    source,
                    "A fully transparent PDF image placement was suppressed instead of exposing its raw image pixels.",
                    PdfCore.PdfConversionWarningSeverity.Information,
                    OfficeConversionLossKind.None,
                    details: details));
                return;
            case PdfCore.PdfImagePlacementImportDisposition.SuppressOutsideVisibleArea:
                warnings.Add(new PdfCore.PdfConversionWarning(
                    "OfficeIMO.Pdf.PowerPoint",
                    "pdf-non-visible-image-placement-suppressed",
                    source,
                    "A PDF image placement with no visible page intersection was suppressed instead of exposing its raw image pixels.",
                    PdfCore.PdfConversionWarningSeverity.Information,
                    OfficeConversionLossKind.None,
                    details: details));
                return;
            case PdfCore.PdfImagePlacementImportDisposition.SuppressUnplaced:
                warnings.Add(new PdfCore.PdfConversionWarning(
                    "OfficeIMO.Pdf.PowerPoint",
                    "pdf-unplaced-image-resource-not-embedded",
                    source,
                    "An extracted image resource without a visible page placement was not embedded as raw pixels.",
                    PdfCore.PdfConversionWarningSeverity.Information,
                    OfficeConversionLossKind.None,
                    details: details));
                return;
            case PdfCore.PdfImagePlacementImportDisposition.OmitClippedPixels:
                AddEditableImageOmissionWarning(warnings, image, "pdf-image-clip-not-safely-editable",
                    "The raw PDF image was not embedded because its clip hides source pixels that an editable PowerPoint picture could reveal.", details);
                return;
            case PdfCore.PdfImagePlacementImportDisposition.OmitSoftMask:
                AddEditableImageOmissionWarning(warnings, image, "pdf-image-placement-soft-mask-not-safely-editable",
                    "The raw PDF image was not embedded because its placement soft mask cannot be reproduced safely as an editable PowerPoint picture.", details);
                return;
            case PdfCore.PdfImagePlacementImportDisposition.OmitUnsupportedBlendMode:
                AddEditableImageOmissionWarning(warnings, image, "pdf-image-unsupported-blend-mode-not-safely-editable",
                    "The raw PDF image was not embedded because its unsupported blend mode cannot be reproduced safely as an editable PowerPoint picture.", details);
                return;
            case PdfCore.PdfImagePlacementImportDisposition.OmitUnsupportedPaintEffect:
                AddEditableImageOmissionWarning(warnings, image, "pdf-image-paint-effect-not-safely-editable",
                    "The raw PDF image was not embedded because its PDF paint effect cannot be reproduced safely as an editable PowerPoint picture.", details);
                return;
            case PdfCore.PdfImagePlacementImportDisposition.OmitUnsupportedTransform:
                AddEditableImageOmissionWarning(warnings, image, "pdf-image-transform-not-safely-editable",
                    "The raw PDF image was not embedded because its rotation or shear cannot be reproduced safely as an editable PowerPoint picture.", details);
                return;
            case PdfCore.PdfImagePlacementImportDisposition.OmitUnappliedDecode:
                AddEditableImageOmissionWarning(warnings, image, "pdf-image-decode-not-safely-editable",
                    "The raw PDF image was not embedded because its PDF decode mapping is not represented by the extracted JPEG 2000 payload.", details);
                return;
            case PdfCore.PdfImagePlacementImportDisposition.OmitUnresolvedTransparencyMask:
                AddEditableImageOmissionWarning(warnings, image, "pdf-image-transparency-mask-not-resolved",
                    "The raw PDF image was not embedded because its unresolved transparency mask could hide pixels that an editable PowerPoint picture would reveal.", details);
                return;
        }

        if (assessment.HasNonDefaultOpacity) {
            bool opacityIsOmitted = assessment.MappedOpacityIsOmitted;
            details["MappedTransparencyPercent"] = assessment.MappedTransparencyPercent.ToString(CultureInfo.InvariantCulture);
            warnings.Add(new PdfCore.PdfConversionWarning(
                "OfficeIMO.Pdf.PowerPoint",
                "pdf-image-opacity-mapped",
                source,
                opacityIsOmitted
                    ? "PDF image opacity was below PowerPoint picture-transparency precision, so the image was made fully transparent."
                    : "PDF image opacity was mapped to native PowerPoint picture transparency.",
                opacityIsOmitted
                    ? PdfCore.PdfConversionWarningSeverity.Warning
                    : PdfCore.PdfConversionWarningSeverity.Information,
                opacityIsOmitted ? OfficeConversionLossKind.Omission : OfficeConversionLossKind.None,
                details: details));
        }
        if (assessment.HasNonNormalBlendMode) {
            warnings.Add(new PdfCore.PdfConversionWarning(
                "OfficeIMO.Pdf.PowerPoint",
                "pdf-image-blend-mode-approximated",
                source,
                "PowerPoint does not reproduce the PDF image blend mode; the safely visible image pixels were embedded with normal compositing.",
                PdfCore.PdfConversionWarningSeverity.Warning,
                OfficeConversionLossKind.Approximation,
                details: details));
        }
    }

    private static void AddEditableImageOmissionWarning(
        ICollection<PdfCore.PdfConversionWarning> warnings,
        PdfCore.PdfLogicalImage image,
        string code,
        string message,
        IReadOnlyDictionary<string, string> details) {
        warnings.Add(new PdfCore.PdfConversionWarning(
            "OfficeIMO.Pdf.PowerPoint",
            code,
            "Page " + image.PageNumber.ToString(CultureInfo.InvariantCulture) + "/Image",
            message,
            PdfCore.PdfConversionWarningSeverity.Warning,
            OfficeConversionLossKind.Omission,
            details: details));
    }

    private static EditableImportCount ImportEditableTextBlocks(
        PdfCore.PdfLogicalPage page,
        PptCore.PowerPointSlide slide,
        EditablePagePlacement placement,
        IReadOnlyList<PdfCore.PdfLogicalTableExtraction> tables,
        int limit,
        CancellationToken cancellationToken) {
        int imported = 0;
        int omitted = 0;
        bool limitReached = false;
        for (int blockIndex = 0; blockIndex < page.TextBlocks.Count; blockIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfCore.PdfLogicalTextBlock block = page.TextBlocks[blockIndex];
            IReadOnlyList<PdfCore.PdfLogicalTextRun>? partialRuns = null;
            if (!CanReconstructEditableTextBlock(block, page)) {
                // Keep the runs that can be projected (for example the letters next to a colour emoji glyph)
                // instead of dropping the whole line; the omission is still reported.
                omitted++;
                partialRuns = block.Runs
                    .Where(run => run.SourceSpan != null && run.SourceSpan.CanProjectCompleteText(page) && !string.IsNullOrWhiteSpace(run.Text))
                    .ToArray();
                if (partialRuns.Count == 0) continue;
            }
            if (IsTextInsideTable(block, tables, cancellationToken)) continue;
            if (imported >= limit) {
                omitted++;
                limitReached = true;
                continue;
            }
            double fontSize = Math.Max(1D, block.FontSize);
            PdfCore.PdfVisualBounds visual = block.VisualBounds is PdfCore.PdfLogicalVisualBounds direct
                ? new PdfCore.PdfVisualBounds(direct.Left, direct.Top, direct.Right, direct.Bottom)
                : page.TransformBoundsToVisual(
                    block.XStart,
                    block.BaselineY - fontSize * 0.3D,
                    Math.Max(block.XStart + 1D, block.XEnd),
                    block.BaselineY + fontSize * 0.9D);
            EditableBounds bounds = placement.Map(
                visual.Left,
                visual.Top,
                Math.Max(visual.Width, fontSize),
                Math.Max(visual.Height, fontSize * 1.2D));
            double sourceRotation = block.Spans.Count > 0 ? block.Spans[0].RotationDegrees : 0D;
            // PDF text angles are counterclockwise in y-up space; page and PowerPoint rotation are clockwise.
            double visualRotation = NormalizeRotation(page.RotationDegrees - sourceRotation);
            if (Math.Abs(visualRotation - 90D) < 0.01D || Math.Abs(visualRotation - 270D) < 0.01D) {
                bounds = bounds.SwapDimensionsAroundCenter();
            }
            // A source line remains one line when substitution changes its measured width.
            double widenedWidth = bounds.Width * 1.08D + fontSize * placement.Scale;
            PptCore.PowerPointTextBox textBox = slide.AddTextBox(string.Empty,
                OfficeLength.Points(bounds.Left), OfficeLength.Points(bounds.Top),
                OfficeLength.Points(widenedWidth), OfficeLength.Points(bounds.Height));
            textBox.SetTextMargins(OfficeLength.Zero, OfficeLength.Zero, OfficeLength.Zero, OfficeLength.Zero);
            if (textBox.TextBody?.BodyProperties is DocumentFormat.OpenXml.Drawing.BodyProperties bodyProperties) {
                bodyProperties.Wrap = DocumentFormat.OpenXml.Drawing.TextWrappingValues.None;
            }
            textBox.FillColor = "FFFFFF";
            textBox.FillTransparency = 100;
            textBox.OutlineColor = "FFFFFF";
            textBox.OutlineTransparency = 100;
            ApplyEditableTextRuns(
                textBox,
                block,
                GetEditableTypographyScale(page, placement.Scale),
                cancellationToken,
                partialRuns);
            if (Math.Abs(visualRotation) > 0.01D) {
                textBox.Rotation = visualRotation;
            }
            imported++;
        }
        return new EditableImportCount(imported, omitted, limitReached);
    }

    private static bool CanReconstructEditableTextBlock(
        PdfCore.PdfLogicalTextBlock block,
        PdfCore.PdfLogicalPage page) =>
        block.Spans.Count == 0
            ? block.VisualBounds != null && block.VisualBounds.Left >= 0D && block.VisualBounds.Top >= 0D &&
                block.VisualBounds.Right <= page.GetVisualPageSize().Width &&
                block.VisualBounds.Bottom <= page.GetVisualPageSize().Height
            : block.Spans.All(span => span.CanProjectCompleteText(page));

    private static void ApplyEditableTextRuns(
        PptCore.PowerPointTextBox textBox,
        PdfCore.PdfLogicalTextBlock block,
        double scale,
        CancellationToken cancellationToken,
        IReadOnlyList<PdfCore.PdfLogicalTextRun>? runs = null) {
        PptCore.PowerPointParagraph paragraph = textBox.Paragraphs[0];
        IReadOnlyList<PdfCore.PdfLogicalTextRun> sourceRuns = runs ?? block.Runs;
        if (sourceRuns.Count == 0) {
            paragraph.Text = block.Text;
            PptCore.PowerPointTextRun targetRun = paragraph.Runs[0];
            targetRun.FontSize = global::OfficeIMO.OfficeLength.Points(ScaleEditableFontSize(block.FontSize, scale));
            return;
        }

        paragraph.Text = sourceRuns[0].Text;
        ApplyEditableTextRunStyle(paragraph.Runs[0], sourceRuns[0], block.FontSize, scale);
        for (int runIndex = 1; runIndex < sourceRuns.Count; runIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfCore.PdfLogicalTextRun sourceRun = sourceRuns[runIndex];
            PptCore.PowerPointTextRun targetRun = paragraph.AddRun(sourceRun.Text);
            ApplyEditableTextRunStyle(targetRun, sourceRun, block.FontSize, scale);
        }
    }

    private static void ApplyEditableTextRunStyle(
        PptCore.PowerPointTextRun target,
        PdfCore.PdfLogicalTextRun source,
        double fallbackFontSize,
        double scale) {
        target.FontSize = global::OfficeIMO.OfficeLength.Points(ScaleEditableFontSize(
            source.FontSize > 0D ? source.FontSize : fallbackFontSize,
            scale));
        target.FontName = ResolvePowerPointFontFamily(source.BaseFont);
        target.Bold = source.IsBold;
        target.Italic = source.IsItalic;
        if (source.Color.HasValue) target.Color = ToHex(source.Color.Value);
    }

    private static EditableTableImportCount ImportEditableTables(
        int pageIndex,
        PdfCore.PdfLogicalPage page,
        IReadOnlyList<PdfCore.PdfLogicalTableExtraction> tables,
        PptCore.PowerPointDocument presentation,
        PptCore.PowerPointSlide primarySlide,
        int primarySlideIndex,
        EditablePagePlacement placement,
        IReadOnlyDictionary<PdfCore.PdfLogicalTable, EditableBounds> tableBounds,
        PowerPointPdfImportOptions options,
        ICollection<PdfPowerPointTableImportEntry> entries,
        int remainingObjects,
        CancellationToken cancellationToken) {
        int primaryCount = 0;
        int omitted = 0;
        for (int tableIndex = 0; tableIndex < tables.Count; tableIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfCore.PdfLogicalTableExtraction extraction = tables[tableIndex];
            PdfCore.PdfLogicalTableData data = extraction.Data;
            if (data.Columns.Count <= 0) continue;
            bool headerRowIncluded = options.IncludeColumnHeaderRows && HasSourceHeaderRow(data);
            List<TableSegment> segments = BuildTableSegments(data, options, cancellationToken);
            for (int segmentIndex = 0; segmentIndex < segments.Count; segmentIndex++) {
                cancellationToken.ThrowIfCancellationRequested();
                TableSegment segment = segments[segmentIndex];
                int rowCount = segment.RowCount + (headerRowIncluded ? 1 : 0);
                if (rowCount <= 0 || segment.ColumnCount <= 0) continue;
                if (remainingObjects <= 0) {
                    omitted++;
                    continue;
                }
                bool primary = segmentIndex == 0;
                PptCore.PowerPointSlide slide = primary ? primarySlide : presentation.AddSlide();
                int slideIndex = primary ? primarySlideIndex : presentation.Slides.Count - 1;
                if (!primary && options.IncludeSourceTitles) {
                    slide.AddTitle(BuildTitle(extraction, segmentIndex, segments.Count));
                }
                EditableBounds bounds = primary
                    ? (tableBounds.TryGetValue(extraction.Table, out EditableBounds sourceBounds) ? sourceBounds : MapTableBounds(page, extraction.Table, placement))
                    : GetContinuationTableBounds(presentation, options);
                PptCore.PowerPointTable table = slide.AddTable(
                    rowCount,                    segment.ColumnCount,                    ResolveReconstructedTableStyle(options.TableStyle),                    global::OfficeIMO.OfficeLength.Emus(PowerPointUnits.FromPoints(bounds.Left)),                    global::OfficeIMO.OfficeLength.Emus(PowerPointUnits.FromPoints(bounds.Top)),                    global::OfficeIMO.OfficeLength.Emus(PowerPointUnits.FromPoints(Math.Max(1D, bounds.Width))),                    global::OfficeIMO.OfficeLength.Emus(PowerPointUnits.FromPoints(Math.Max(1D, bounds.Height))));
                PopulateTable(table, extraction.Table, data, segment, headerRowIncluded, options, cancellationToken);
                ApplyEditableTableTypography(
                    table,
                    page,
                    extraction.Table,
                    GetEditableTypographyScale(page, placement.Scale),
                    isContinuation: !primary,
                    cancellationToken);
                entries.Add(new PdfPowerPointTableImportEntry(
                    pageIndex,
                    extraction.PageNumber,
                    extraction.TableIndex,
                    extraction.DetectionKind,
                    slideIndex,
                    segmentIndex,
                    segments.Count,
                    segment.RowStartIndex,
                    segment.ColumnStartIndex,
                    data.Columns.Count,
                    segment.ColumnCount,
                    segment.RowCount,
                    data.TotalRowCount,
                    data.Truncated,
                    headerRowIncluded));
                if (primary) primaryCount++;
                remainingObjects--;
            }
        }
        return new EditableTableImportCount(primaryCount, omitted, omitted > 0);
    }

    private static EditableBounds GetContinuationTableBounds(
        PptCore.PowerPointDocument presentation,
        PowerPointPdfImportOptions options) {
        double slideWidth = Math.Max(1D, presentation.SlideSize.Width.ToPoints());
        double slideHeight = Math.Max(1D, presentation.SlideSize.Height.ToPoints());
        double left = Math.Min(Math.Max(0D, options.TableLeft / 12700D), slideWidth - 1D);
        double top = Math.Min(Math.Max(0D, options.TableTop / 12700D), slideHeight - 1D);
        double width = Math.Min(Math.Max(1D, options.TableWidth / 12700D), slideWidth - left);
        double height = Math.Min(Math.Max(1D, options.TableHeight / 12700D), slideHeight - top);
        return new EditableBounds(left, top, width, height);
    }

    private static Dictionary<PdfCore.PdfLogicalTable, EditableBounds> GetEditableTableBounds(
        PdfCore.PdfLogicalPage page,
        IReadOnlyList<PdfCore.PdfLogicalTableExtraction> tables,
        EditablePagePlacement placement,
        OfficeDrawing? drawing,
        CancellationToken cancellationToken) {
        var bounds = new Dictionary<PdfCore.PdfLogicalTable, EditableBounds>();
        EditableBounds[] cellShapes = drawing == null
            ? Array.Empty<EditableBounds>()
            : drawing.Elements.OfType<OfficeDrawingShape>()
                .Where(static shape => shape.Shape.Kind is OfficeShapeKind.Rectangle or OfficeShapeKind.Line)
                .Select(shape => placement.Map(shape.X, shape.Y, shape.Shape.Width, shape.Shape.Height))
                .ToArray();
        for (int i = 0; i < tables.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (tables[i].Table.Columns.Count > 0 && !bounds.ContainsKey(tables[i].Table)) {
                bounds.Add(tables[i].Table, GrowToCellShapes(MapTableBounds(page, tables[i].Table, placement), cellShapes, placement.Scale));
            }
        }
        return bounds;
    }

    /// <summary>
    /// Text gives only the inner extent of a table. Grow it to the cell fills and rules drawn around and next to that
    /// text (a last column whose cells are wider than their text), so the native table covers the drawn grid.
    /// </summary>
    private static EditableBounds GrowToCellShapes(EditableBounds bounds, EditableBounds[] shapes, double scale) {
        EditableBounds grown = GrowToCellShapesCore(bounds, shapes, scale);
        // Never chain into unrelated rules or panels far beyond the text grid.
        double maxHorizontal = bounds.Width * 0.5D, maxVertical = Math.Max(bounds.Height * 0.25D, 12D * scale);
        double left = Math.Max(grown.Left, bounds.Left - maxHorizontal);
        double top = Math.Max(grown.Top, bounds.Top - maxVertical);
        double right = Math.Min(grown.Left + grown.Width, bounds.Left + bounds.Width + maxHorizontal);
        double bottom = Math.Min(grown.Top + grown.Height, bounds.Top + bounds.Height + maxVertical);
        return new EditableBounds(left, top, right - left, bottom - top);
    }

    private static EditableBounds GrowToCellShapesCore(EditableBounds bounds, EditableBounds[] shapes, double scale) {
        double tolerance = 2D * scale;
        for (int pass = 0; pass < 8; pass++) {
            double left = bounds.Left, top = bounds.Top, right = bounds.Left + bounds.Width, bottom = bounds.Top + bounds.Height;
            foreach (EditableBounds shape in shapes) {
                double shapeRight = shape.Left + shape.Width, shapeBottom = shape.Top + shape.Height;
                double centerY = shape.Top + shape.Height / 2D, centerX = shape.Left + shape.Width / 2D;
                bool rowAligned = centerY >= top && centerY <= bottom && shape.Height <= (bottom - top) + tolerance;
                bool columnAligned = centerX >= left && centerX <= right && shape.Width <= (right - left) + tolerance;
                bool touchesHorizontally = shape.Left <= right + tolerance && shapeRight >= left - tolerance;
                bool touchesVertically = shape.Top <= bottom + tolerance && shapeBottom >= top - tolerance;
                if (rowAligned && touchesHorizontally) {
                    left = Math.Min(left, shape.Left);
                    right = Math.Max(right, shapeRight);
                }
                if (columnAligned && touchesVertically) {
                    top = Math.Min(top, shape.Top);
                    bottom = Math.Max(bottom, shapeBottom);
                }
            }
            var grown = new EditableBounds(left, top, right - left, bottom - top);
            if (Math.Abs(grown.Left - bounds.Left) < 0.01D && Math.Abs(grown.Top - bounds.Top) < 0.01D &&
                Math.Abs(grown.Width - bounds.Width) < 0.01D && Math.Abs(grown.Height - bounds.Height) < 0.01D) return grown;
            bounds = grown;
        }
        return bounds;
    }

    private static EditableBounds MapTableBounds(
        PdfCore.PdfLogicalPage page,
        PdfCore.PdfLogicalTable table,
        EditablePagePlacement placement) {
        EditableBounds textBounds = MapTableTextBounds(page, table, placement);
        // Table geometry comes from its text; the drawn cells extend about half a row beyond the first and last
        // baselines and a few points beyond the outer text, so the native table covers the source cell area.
        double rowPitch = textBounds.Height / Math.Max(1, table.Rows.Count);
        double vertical = Math.Max(2D * placement.Scale, rowPitch * 0.35D);
        double horizontal = 4D * placement.Scale;
        return new EditableBounds(
            textBounds.Left - horizontal,
            textBounds.Top - vertical,
            textBounds.Width + horizontal * 2D,
            textBounds.Height + vertical * 2D);
    }

    private static EditableBounds MapTableTextBounds(
        PdfCore.PdfLogicalPage page,
        PdfCore.PdfLogicalTable table,
        EditablePagePlacement placement) {
        if (table.VisualBounds is PdfCore.PdfLogicalVisualBounds direct) {
            return placement.Map(direct.Left, direct.Top, direct.Width, direct.Height);
        }
        double left = table.Columns.Min(static column => column.From);
        double right = table.Columns.Max(static column => column.To);
        PdfCore.PdfVisualBounds visual = page.TransformBoundsToVisual(
            left,
            table.YBottom,
            right,
            table.YTop);
        return placement.Map(visual.Left, visual.Top, visual.Width, visual.Height);
    }

    private static bool IsTextInsideTable(
        PdfCore.PdfLogicalTextBlock block,
        IReadOnlyList<PdfCore.PdfLogicalTableExtraction> tables,
        CancellationToken cancellationToken) {
        for (int i = 0; i < tables.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfCore.PdfLogicalTable table = tables[i].Table;
            // Text inside the table area that the table did not recover (a row the detector missed) is still
            // imported as a text box rather than silently dropped.
            if (!IsTextRepresentedByTable(block, table)) continue;
            if (block.VisualBounds is PdfCore.PdfLogicalVisualBounds blockBounds &&
                table.VisualBounds is PdfCore.PdfLogicalVisualBounds tableBounds) {
                double centerX = (blockBounds.Left + blockBounds.Right) / 2D;
                double centerY = (blockBounds.Top + blockBounds.Bottom) / 2D;
                if (centerX >= tableBounds.Left && centerX <= tableBounds.Right &&
                    centerY >= tableBounds.Top && centerY <= tableBounds.Bottom) return true;
                continue;
            }
            if (table.Columns.Count == 0 ||
                block.BaselineY < table.YBottom ||
                block.BaselineY > table.YTop) continue;
            double left = table.Columns.Min(static column => column.From);
            double right = table.Columns.Max(static column => column.To);
            if (block.XEnd >= left && block.XStart <= right) return true;
        }
        return false;
    }

    private static bool IsTextRepresentedByTable(PdfCore.PdfLogicalTextBlock block, PdfCore.PdfLogicalTable table) {
        string text = NormalizeCellText(block.Text);
        if (text.Length == 0) return true;
        foreach (IReadOnlyList<string> row in table.Rows) {
            foreach (string cell in row) {
                string normalized = NormalizeCellText(cell);
                if (normalized.Length > 0 && (ContainsOrdinal(normalized, text) || ContainsOrdinal(text, normalized))) return true;
            }
        }
        return false;
    }

    private static bool ContainsOrdinal(string value, string candidate) {
#if NETSTANDARD2_0 || NETFRAMEWORK
        return value.IndexOf(candidate, StringComparison.Ordinal) >= 0;
#else
        return value.Contains(candidate, StringComparison.Ordinal);
#endif
    }

    private static string NormalizeCellText(string? value) =>
        string.Concat((value ?? string.Empty).Where(static ch => !char.IsWhiteSpace(ch)));

    private static OfficeImageFormat ResolveImageFormat(PdfCore.PdfExtractedImage image) =>
        image.MimeType?.ToLowerInvariant() switch {
            "image/png" => OfficeImageFormat.Png,
            "image/jpeg" => OfficeImageFormat.Jpeg,
            "image/gif" => OfficeImageFormat.Gif,
            "image/bmp" => OfficeImageFormat.Bmp,
            "image/tiff" => OfficeImageFormat.Tiff,
            "image/jp2" => OfficeImageFormat.Jpeg2000,
            _ => image.FileExtension?.TrimStart('.').ToLowerInvariant() switch {
                "png" => OfficeImageFormat.Png,
                "jpg" or "jpeg" => OfficeImageFormat.Jpeg,
                "gif" => OfficeImageFormat.Gif,
                "bmp" => OfficeImageFormat.Bmp,
                "tif" or "tiff" => OfficeImageFormat.Tiff,
                "jp2" => OfficeImageFormat.Jpeg2000,
                _ => OfficeImageFormat.Unknown
            }
        };

    private static string ResolvePowerPointFontFamily(string? baseFont) {
        string? family = PdfCore.PdfFontFamilyNames.ToFamilyName(baseFont);
        if (string.IsNullOrWhiteSpace(family)) return "Arial";
        // Helvetica is not installed with Office; Arial has the same metrics.
        return family!.StartsWith("Helvetica", StringComparison.OrdinalIgnoreCase) ? "Arial" : family!;
    }

    private static void AddEditablePageWarnings(
        ICollection<PdfCore.PdfConversionWarning> warnings,
        int pageNumber,
        int omittedText,
        int omittedVectors,
        int omittedImages,
        bool limitReached) {
        string source = "PDF page " + pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (omittedText > 0) {
            warnings.Add(CreateEditableOmissionWarning(
                "pdf-text-not-reconstructed",
                source,
                "Text blocks",
                omittedText,
                pageNumber));
        }
        if (omittedVectors > 0) {
            warnings.Add(CreateEditableOmissionWarning(
                "pdf-vectors-not-reconstructed",
                source,
                "Vector primitives",
                omittedVectors,
                pageNumber));
        }
        if (omittedImages > 0) {
            warnings.Add(CreateEditableOmissionWarning(
                "pdf-images-not-reconstructed",
                source,
                "Images",
                omittedImages,
                pageNumber));
        }
        if (limitReached) {
            warnings.Add(new PdfCore.PdfConversionWarning(
                "OfficeIMO.Pdf.PowerPoint",
                "pdf-editable-object-limit-reached",
                source,
                "The per-page editable object limit was reached; remaining objects were omitted.",
                PdfCore.PdfConversionWarningSeverity.Warning,
                OfficeConversionLossKind.Omission,
                details: new Dictionary<string, string> {
                    ["pageNumber"] = pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    ["construct"] = "Editable objects",
                    ["Disposition"] = "Omitted"
                }));
        }
    }

    private static PdfCore.PdfConversionWarning CreateEditableOmissionWarning(
        string code,
        string source,
        string construct,
        int count,
        int pageNumber) => new(
            "OfficeIMO.Pdf.PowerPoint",
            code,
            source,
            count.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                " " + construct.ToLowerInvariant() +
                " could not be reconstructed safely as editable PowerPoint objects.",
            PdfCore.PdfConversionWarningSeverity.Warning,
            OfficeConversionLossKind.Omission,
            details: new Dictionary<string, string> {
                ["pageNumber"] = pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture),
                ["construct"] = construct,
                ["Count"] = count.ToString(System.Globalization.CultureInfo.InvariantCulture),
                ["Disposition"] = "Omitted"
            });

    private static void AddEditableRendererWarnings(
        ICollection<PdfCore.PdfConversionWarning> warnings,
        int pageNumber,
        IReadOnlyList<PdfCore.PdfRenderCapabilityDiagnostic> diagnostics,
        CancellationToken cancellationToken) {
        for (int i = 0; i < diagnostics.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfCore.PdfRenderCapabilityDiagnostic diagnostic = diagnostics[i];
            if (diagnostic.Code.Contains("font", StringComparison.OrdinalIgnoreCase)) continue;
            warnings.Add(new PdfCore.PdfConversionWarning(
                "OfficeIMO.Pdf.PowerPoint",
                diagnostic.Code,
                "PDF page " + pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture),
                diagnostic.Message,
                diagnostic.SupportLevel == PdfCore.PdfRenderSupportLevel.Unsupported
                    ? PdfCore.PdfConversionWarningSeverity.Warning
                    : PdfCore.PdfConversionWarningSeverity.Information,
                diagnostic.LossKind,
                details: new Dictionary<string, string> {
                    ["pageNumber"] = pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    ["construct"] = diagnostic.Capability.Feature,
                    ["supportLevel"] = diagnostic.SupportLevel.ToString(),
                    ["Disposition"] = diagnostic.SupportLevel == PdfCore.PdfRenderSupportLevel.Unsupported
                        ? "Omitted"
                        : "Simplified"
                }));
        }
    }

    private static void AddEditableDocumentWarnings(
        ICollection<PdfCore.PdfConversionWarning> warnings,
        PdfCore.PdfTableExtractionScopeReport scope) {
        warnings.Add(new PdfCore.PdfConversionWarning(
            "OfficeIMO.Pdf.PowerPoint",
            "pdf-editable-content-reconstructed",
            "Document",
            "Text blocks, detected tables, safe vector primitives, and supported images were reconstructed as separate PowerPoint objects. Original grouping, charts, and authoring intent cannot be recovered reliably from arbitrary PDFs.",
            PdfCore.PdfConversionWarningSeverity.Information,
            OfficeConversionLossKind.Approximation,
            details: new Dictionary<string, string> {
                ["construct"] = "Editable content",
                ["Disposition"] = "Reconstructed"
            }));
        AddEditableDocumentOmission(warnings, "pdf-navigation-not-reconstructed", "Navigation", scope.LinkCount + scope.PageActionCount, "links and page actions");
        AddEditableDocumentOmission(warnings, "pdf-document-actions-not-reconstructed", "Document actions", scope.DocumentActionCount, "catalog and document-open actions");
        AddEditableDocumentOmission(warnings, "pdf-forms-not-reconstructed", "Forms", scope.FormContentCount, "forms and interactive controls");
        AddEditableDocumentOmission(warnings, "pdf-outlines-not-reconstructed", "Outlines", scope.OutlineCount, "outline navigation entries");
        AddEditableDocumentOmission(warnings, "pdf-attachments-not-reconstructed", "Attachments", scope.AttachmentCount, "embedded attachments");
        AddEditableDocumentOmission(warnings, "pdf-tagged-structure-not-reconstructed", "Tagged structure", scope.HasTaggedContent ? 1 : 0, "tagged accessibility structure");
        AddEditableDocumentOmission(warnings, "pdf-source-security-not-reconstructed", "Document security", scope.HasSourceSecurityState ? 1 : 0, "encryption, signature, permission, or revision state");
        AddEditableDocumentOmission(warnings, "pdf-page-labels-not-reconstructed", "Page labels", scope.PageLabelCount, "page-label rules");
        AddEditableDocumentOmission(warnings, "pdf-metadata-not-reconstructed", "Document metadata", scope.HasDocumentMetadata ? 1 : 0, "document metadata");
        AddEditableDocumentOmission(warnings, "pdf-annotations-not-reconstructed", "Annotations", scope.AnnotationCount, "annotations");
        AddEditableDocumentOmission(warnings, "pdf-groups-not-reconstructed", "Groups", scope.PagesWithOptionalContent, "pages using optional content");
        AddEditableDocumentOmission(warnings, "pdf-animations-not-reconstructed", "Animations", scope.InteractiveMediaAnnotationCount, "interactive media and animations");
        if (scope.AnalysisTruncated) {
            AddEditableDocumentOmission(warnings, "pdf-projection-analysis-truncated", "Document", 1, "bounded source-content analysis");
        }
    }

    private static void AddEditableDocumentOmission(
        ICollection<PdfCore.PdfConversionWarning> warnings,
        string code,
        string construct,
        int count,
        string description) {
        if (count <= 0) return;
        warnings.Add(new PdfCore.PdfConversionWarning(
            "OfficeIMO.Pdf.PowerPoint",
            code,
            "Document",
            "PDF " + description + " are not reconstructed in editable-content mode.",
            PdfCore.PdfConversionWarningSeverity.Warning,
            OfficeConversionLossKind.Omission,
            details: new Dictionary<string, string> {
                ["construct"] = construct,
                ["Count"] = count.ToString(System.Globalization.CultureInfo.InvariantCulture),
                ["Disposition"] = "Omitted"
            }));
    }

    private static string ToHex(OfficeColor color) =>
        color.R.ToString("X2", System.Globalization.CultureInfo.InvariantCulture) +
        color.G.ToString("X2", System.Globalization.CultureInfo.InvariantCulture) +
        color.B.ToString("X2", System.Globalization.CultureInfo.InvariantCulture);

    private static int ToTransparency(double? opacity) =>
        opacity.HasValue
            ? Math.Min(100, Math.Max(0, (int)Math.Round((1D - opacity.Value) * 100D)))
            : 0;

    private static double NormalizeRotation(double value) {
        double normalized = value % 360D;
        return normalized < 0D ? normalized + 360D : normalized;
    }

    private readonly record struct EditableImportCount(int Imported, int Omitted, bool LimitReached);

    private readonly record struct EditableTableImportCount(int PrimarySlideCount, int Omitted, bool LimitReached);

    private readonly record struct EditablePagePlacement(double Left, double Top, double Scale) {
        internal double MapX(double value) => Left + value * Scale;
        internal double MapY(double value) => Top + value * Scale;
        internal EditableBounds Map(double left, double top, double width, double height) =>
            new(MapX(left), MapY(top), Math.Max(0.01D, width * Scale), Math.Max(0.01D, height * Scale));
    }

    private readonly record struct EditableBounds(double Left, double Top, double Width, double Height) {
        internal EditableBounds SwapDimensionsAroundCenter() => new(
            Left + (Width - Height) / 2D,
            Top + (Height - Width) / 2D,
            Height,
            Width);

        internal EditableBounds Inflate(double amount) => new(Left - amount, Top - amount, Width + amount * 2D, Height + amount * 2D);

        internal double OverlapFractionOf(EditableBounds candidate) {
            double width = Math.Min(Left + Width, candidate.Left + candidate.Width) - Math.Max(Left, candidate.Left);
            double height = Math.Min(Top + Height, candidate.Top + candidate.Height) - Math.Max(Top, candidate.Top);
            double area = Math.Max(0.0001D, candidate.Width * candidate.Height);
            return width <= 0D || height <= 0D ? 0D : width * height / area;
        }

        internal bool ContainsCenterOf(EditableBounds candidate) {
            double x = candidate.Left + candidate.Width / 2D;
            double y = candidate.Top + candidate.Height / 2D;
            return x >= Left && x <= Left + Width && y >= Top && y <= Top + Height;
        }
    }
}
