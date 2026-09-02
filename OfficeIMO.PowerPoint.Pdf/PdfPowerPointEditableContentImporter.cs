using OfficeIMO.Drawing;
using System.Threading;
using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    private static PdfPowerPointConversionResult ImportEditableContent(
        PdfCore.PdfDocument document,
        PdfPowerPointImportOptions options,
        CancellationToken cancellationToken) {
        PdfCore.PdfDocumentReadResult logical = ReadBoundedLogicalDocument(document, options, cancellationToken);
        return ImportEditableContent(logical, options, document, cancellationToken);
    }

    private static PdfPowerPointConversionResult ImportEditableContent(
        PdfCore.PdfDocumentReadResult logical,
        PdfPowerPointImportOptions options,
        CancellationToken cancellationToken) =>
        ImportEditableContent(logical, options, sourceDocument: null, cancellationToken);

    private static PdfPowerPointConversionResult ImportEditableContent(
        PdfCore.PdfDocumentReadResult logical,
        PdfPowerPointImportOptions options,
        PdfCore.PdfDocument? sourceDocument,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (options.MaxEditableObjectsPerPage <= 0) {
            throw new ArgumentOutOfRangeException(
                nameof(options.MaxEditableObjectsPerPage),
                "The editable object limit must be positive.");
        }

        using var presentationOwner = new PresentationConstructionScope();
        PptCore.PowerPointPresentation presentation = presentationOwner.Value;
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
                presentation.SlideSize.WidthPoints,
                presentation.SlideSize.HeightPoints);
            IReadOnlyList<PdfCore.PdfLogicalTableExtraction> tables =
                tablesByPage.TryGetValue(pageIndex, out IReadOnlyList<PdfCore.PdfLogicalTableExtraction>? pageTables)
                    ? pageTables
                    : Array.Empty<PdfCore.PdfLogicalTableExtraction>();
            List<EditableBounds> tableBounds = GetEditableTableBounds(page, tables, placement, cancellationToken);
            int remainingObjects = options.MaxEditableObjectsPerPage;

            EditableImportCount shapeImport = drawing == null
                ? new EditableImportCount(
                    0,
                    Math.Max(0, page.VectorPrimitiveCount - page.UnrepresentedVectorPrimitiveCount),
                    false)
                : ImportEditableShapes(drawing, slide, placement, tableBounds, remainingObjects, cancellationToken);
            remainingObjects -= shapeImport.Imported;
            EditableImportCount imageImport = ImportEditableImages(
                page,
                slide,
                placement,
                remainingObjects,
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
                options,
                tableEntries,
                remainingObjects,
                cancellationToken);

            int omittedVectors = checked(shapeImport.Omitted + page.UnrepresentedVectorPrimitiveCount);
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
        PptCore.PowerPointPresentation presentation,
        OfficeDrawing? drawing) {
        if (drawing == null || drawing.Width <= 0D || drawing.Height <= 0D) return;
        ConfigureEditableSlideSize(presentation, drawing.Width, drawing.Height);
    }

    private static void ConfigureEditableSlideSize(
        PptCore.PowerPointPresentation presentation,
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
        presentation.SlideSize.SetSizePoints(width, height);
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
            EditableBounds visualBounds = new(
                drawingShape.X,
                drawingShape.Y,
                drawingShape.Shape.Width,
                drawingShape.Shape.Height);
            if (tableBounds.Any(bounds => bounds.ContainsCenterOf(visualBounds))) {
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
                target = slide.AddShapePoints(OfficePresetShapeType.Rectangle, left, top, width, height);
                break;
            case OfficeShapeKind.RoundedRectangle:
                target = slide.AddShapePoints(OfficePresetShapeType.RoundRectangle, left, top, width, height);
                break;
            case OfficeShapeKind.Ellipse:
                target = slide.AddShapePoints(OfficePresetShapeType.Ellipse, left, top, width, height);
                break;
            case OfficeShapeKind.Line when source.Points.Count >= 2:
                OfficePoint start = source.Points[0];
                OfficePoint end = source.Points[source.Points.Count - 1];
                target = slide.AddLinePoints(
                    placement.MapX(drawingShape.X + start.X),
                    placement.MapY(drawingShape.Y + start.Y),
                    placement.MapX(drawingShape.X + end.X),
                    placement.MapY(drawingShape.Y + end.Y));
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
        target.OutlineWidthPoints = Math.Max(0D, source.StrokeWidth);
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
        CancellationToken cancellationToken) {
        int imported = 0;
        int omitted = 0;
        bool limitReached = false;
        for (int imageIndex = 0; imageIndex < page.Images.Count; imageIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfCore.PdfLogicalImage image = page.Images[imageIndex];
            OfficeImageFormat format = ResolveImageFormat(image.SourceImage);
            if (!image.SourceImage.IsImageFile ||
                image.SourceImage.IsImageMask ||
                image.SourceImage.HasUnresolvedTransparencyMask ||
                format == OfficeImageFormat.Unknown ||
                image.Placements.Count == 0) {
                omitted++;
                continue;
            }
            for (int placementIndex = 0; placementIndex < image.Placements.Count; placementIndex++) {
                cancellationToken.ThrowIfCancellationRequested();
                PdfCore.PdfImagePlacement sourcePlacement = image.Placements[placementIndex];
                if (!sourcePlacement.IsAxisAligned || imported >= limit) {
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
                using var stream = new MemoryStream(image.SourceImage.Bytes, writable: false);
                slide.AddPicturePoints(stream, format, bounds.Left, bounds.Top, bounds.Width, bounds.Height);
                imported++;
            }
        }
        return new EditableImportCount(imported, omitted, limitReached);
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
            if (!CanReconstructEditableTextBlock(block, page.Height)) {
                omitted++;
                continue;
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
            PptCore.PowerPointTextBox textBox = slide.AddTextBoxPoints(
                string.Empty,
                bounds.Left,
                bounds.Top,
                bounds.Width,
                bounds.Height);
            textBox.SetTextMarginsPoints(0D, 0D, 0D, 0D);
            textBox.FillColor = "FFFFFF";
            textBox.FillTransparency = 100;
            textBox.OutlineColor = "FFFFFF";
            textBox.OutlineTransparency = 100;
            ApplyEditableTextRuns(textBox, block, placement.Scale, cancellationToken);
            double sourceRotation = block.Spans.Count > 0 ? block.Spans[0].RotationDegrees : 0D;
            double visualRotation = -(page.RotationDegrees + sourceRotation);
            if (Math.Abs(visualRotation) > 0.01D) {
                textBox.Rotation = NormalizeRotation(visualRotation);
            }
            imported++;
        }
        return new EditableImportCount(imported, omitted, limitReached);
    }

    private static bool CanReconstructEditableTextBlock(
        PdfCore.PdfLogicalTextBlock block,
        double pageHeight) =>
        block.Spans.Count == 0
            ? block.VisualBounds != null
            : block.Spans.All(span => span.CanProjectCompleteText(pageHeight));

    private static void ApplyEditableTextRuns(
        PptCore.PowerPointTextBox textBox,
        PdfCore.PdfLogicalTextBlock block,
        double scale,
        CancellationToken cancellationToken) {
        PptCore.PowerPointParagraph paragraph = textBox.Paragraphs[0];
        IReadOnlyList<PdfCore.PdfLogicalTextRun> sourceRuns = block.Runs;
        if (sourceRuns.Count == 0) {
            paragraph.Text = block.Text;
            PptCore.PowerPointTextRun targetRun = paragraph.Runs[0];
            targetRun.FontSizePoints = ScaleEditableFontSize(block.FontSize, scale);
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
        target.FontSizePoints = ScaleEditableFontSize(
            source.FontSize > 0D ? source.FontSize : fallbackFontSize,
            scale);
        target.FontName = ResolvePowerPointFontFamily(source.BaseFont);
        target.Bold = source.IsBold;
        target.Italic = source.IsItalic;
        if (source.Color.HasValue) target.Color = ToHex(source.Color.Value);
    }

    private static EditableTableImportCount ImportEditableTables(
        int pageIndex,
        PdfCore.PdfLogicalPage page,
        IReadOnlyList<PdfCore.PdfLogicalTableExtraction> tables,
        PptCore.PowerPointPresentation presentation,
        PptCore.PowerPointSlide primarySlide,
        int primarySlideIndex,
        EditablePagePlacement placement,
        PdfPowerPointImportOptions options,
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
                    ? MapTableBounds(page, extraction.Table, placement)
                    : GetContinuationTableBounds(presentation, options);
                PptCore.PowerPointTable table = slide.AddTable(
                    rowCount,
                    segment.ColumnCount,
                    options.TableStyle,
                    PowerPointUnits.FromPoints(bounds.Left),
                    PowerPointUnits.FromPoints(bounds.Top),
                    PowerPointUnits.FromPoints(Math.Max(1D, bounds.Width)),
                    PowerPointUnits.FromPoints(Math.Max(1D, bounds.Height)));
                PopulateTable(table, extraction.Table, data, segment, headerRowIncluded, options, cancellationToken);
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
        PptCore.PowerPointPresentation presentation,
        PdfPowerPointImportOptions options) {
        double slideWidth = Math.Max(1D, presentation.SlideSize.WidthPoints);
        double slideHeight = Math.Max(1D, presentation.SlideSize.HeightPoints);
        double left = Math.Min(Math.Max(0D, options.TableLeft / 12700D), slideWidth - 1D);
        double top = Math.Min(Math.Max(0D, options.TableTop / 12700D), slideHeight - 1D);
        double width = Math.Min(Math.Max(1D, options.TableWidth / 12700D), slideWidth - left);
        double height = Math.Min(Math.Max(1D, options.TableHeight / 12700D), slideHeight - top);
        return new EditableBounds(left, top, width, height);
    }

    private static List<EditableBounds> GetEditableTableBounds(
        PdfCore.PdfLogicalPage page,
        IReadOnlyList<PdfCore.PdfLogicalTableExtraction> tables,
        EditablePagePlacement placement,
        CancellationToken cancellationToken) {
        var bounds = new List<EditableBounds>(tables.Count);
        for (int i = 0; i < tables.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (tables[i].Table.Columns.Count > 0) {
                bounds.Add(MapTableBounds(page, tables[i].Table, placement));
            }
        }
        return bounds;
    }

    private static EditableBounds MapTableBounds(
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

    private static OfficeImageFormat ResolveImageFormat(PdfCore.PdfExtractedImage image) =>
        image.MimeType?.ToLowerInvariant() switch {
            "image/png" => OfficeImageFormat.Png,
            "image/jpeg" => OfficeImageFormat.Jpeg,
            "image/gif" => OfficeImageFormat.Gif,
            "image/bmp" => OfficeImageFormat.Bmp,
            "image/tiff" => OfficeImageFormat.Tiff,
            _ => image.FileExtension?.TrimStart('.').ToLowerInvariant() switch {
                "png" => OfficeImageFormat.Png,
                "jpg" or "jpeg" => OfficeImageFormat.Jpeg,
                "gif" => OfficeImageFormat.Gif,
                "bmp" => OfficeImageFormat.Bmp,
                "tif" or "tiff" => OfficeImageFormat.Tiff,
                _ => OfficeImageFormat.Unknown
            }
        };

    private static string ResolvePowerPointFontFamily(string? baseFont) {
        if (string.IsNullOrWhiteSpace(baseFont)) return "Arial";
        string value = baseFont!.Trim();
        if (value.Length > 7 && value[6] == '+') value = value.Substring(7);
        if (value.StartsWith("Helvetica", StringComparison.OrdinalIgnoreCase)) return "Arial";
        if (value.StartsWith("Times", StringComparison.OrdinalIgnoreCase)) return "Times New Roman";
        if (value.StartsWith("Courier", StringComparison.OrdinalIgnoreCase)) return "Courier New";
        int delimiter = value.IndexOfAny(['-', ',']);
        return delimiter > 0 ? value.Substring(0, delimiter) : value;
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
                "PdfTextNotReconstructed",
                source,
                "Text blocks",
                omittedText,
                pageNumber));
        }
        if (omittedVectors > 0) {
            warnings.Add(CreateEditableOmissionWarning(
                "PdfVectorsNotReconstructed",
                source,
                "Vector primitives",
                omittedVectors,
                pageNumber));
        }
        if (omittedImages > 0) {
            warnings.Add(CreateEditableOmissionWarning(
                "PdfImagesNotReconstructed",
                source,
                "Images",
                omittedImages,
                pageNumber));
        }
        if (limitReached) {
            warnings.Add(new PdfCore.PdfConversionWarning(
                "OfficeIMO.PowerPoint.Pdf",
                "PdfEditableObjectLimitReached",
                source,
                "The per-page editable object limit was reached; remaining objects were omitted.",
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
            "OfficeIMO.PowerPoint.Pdf",
            code,
            source,
            count.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                " " + construct.ToLowerInvariant() +
                " could not be reconstructed safely as editable PowerPoint objects.",
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
                "OfficeIMO.PowerPoint.Pdf",
                diagnostic.Code,
                "PDF page " + pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture),
                diagnostic.Message,
                diagnostic.SupportLevel == PdfCore.PdfRenderSupportLevel.Unsupported
                    ? PdfCore.PdfConversionWarningSeverity.Warning
                    : PdfCore.PdfConversionWarningSeverity.Information,
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
            "OfficeIMO.PowerPoint.Pdf",
            "PdfEditableContentReconstructed",
            "Document",
            "Text blocks, detected tables, safe vector primitives, and supported images were reconstructed as separate PowerPoint objects. Original grouping, charts, and authoring intent cannot be recovered reliably from arbitrary PDFs.",
            PdfCore.PdfConversionWarningSeverity.Information,
            details: new Dictionary<string, string> {
                ["construct"] = "Editable content",
                ["Disposition"] = "Reconstructed"
            }));
        AddEditableDocumentOmission(warnings, "PdfNavigationNotReconstructed", "Navigation", scope.LinkCount + scope.PageActionCount, "links and page actions");
        AddEditableDocumentOmission(warnings, "PdfFormsNotReconstructed", "Forms", scope.FormWidgetCount, "forms and interactive controls");
        AddEditableDocumentOmission(warnings, "PdfAnnotationsNotReconstructed", "Annotations", scope.AnnotationCount, "annotations");
        AddEditableDocumentOmission(warnings, "PdfGroupsNotReconstructed", "Groups", scope.OptionalContentGroupCount, "optional-content groups");
        AddEditableDocumentOmission(warnings, "PdfAnimationsNotReconstructed", "Animations", scope.InteractiveMediaAnnotationCount, "interactive media and animations");
        if (scope.AnalysisTruncated) {
            AddEditableDocumentOmission(warnings, "PdfProjectionAnalysisTruncated", "Document", 1, "bounded source-content analysis");
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
            "OfficeIMO.PowerPoint.Pdf",
            code,
            "Document",
            "PDF " + description + " are not reconstructed in editable-content mode.",
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

    private static double ScaleEditableFontSize(double fontSize, double scale) =>
        Math.Min(4000D, Math.Max(1D, fontSize * scale));

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
        internal bool ContainsCenterOf(EditableBounds candidate) {
            double x = candidate.Left + candidate.Width / 2D;
            double y = candidate.Top + candidate.Height / 2D;
            return x >= Left && x <= Left + Width && y >= Top && y <= Top + Height;
        }
    }
}
