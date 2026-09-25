using OfficeIMO.Drawing;
using OfficeIMO.Html;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Html;

public static partial class HtmlPowerPointConverterExtensions {
    private static void ImportGenericDocument(
        HtmlSemanticDocument document,
        PptCore.PowerPointPresentation presentation,
        HtmlToPowerPointOptions options,
        HtmlToPowerPointResult result,
        HtmlImportBudget budget,
        HtmlEditableLayoutProjection? editableLayout) {
        var firstSlideIndexes = new List<int>();
        double slideBottom = presentation.SlideSize.HeightPoints - 30D;
        bool slideLimitReached = false;
        foreach (HtmlSemanticSection section in document.Sections) {
            if (!TryAddGenericSlide(presentation, result, budget, out PptCore.PowerPointSlide slide)) break;
            firstSlideIndexes.Add(presentation.Slides.Count - 1);
            double contentTop = 30D;
            if (!string.IsNullOrWhiteSpace(section.Title)) {
                HtmlSemanticBlock? titleBlock = section.Blocks.FirstOrDefault();
                contentTop = ImportTextBox(titleBlock?.SourceElement, section.Title, slide, 30D, result, budget, 44D);
            }

            double pictureTop = contentTop;
            foreach (HtmlSemanticBlock block in section.Blocks) {
                bool isSectionTitle = block.Kind == HtmlSemanticBlockKind.Heading
                    && string.Equals(block.Text, section.Title, StringComparison.Ordinal);
                bool importText = IsGenericTextBlock(block.Kind);
                bool importTable = options.ImportTables && block.Kind == HtmlSemanticBlockKind.Table;
                bool importPicture = options.ImportPictures && block.Kind == HtmlSemanticBlockKind.Image;
                if (importText && !isSectionTitle) {
                    double textHeight = block.Kind == HtmlSemanticBlockKind.List
                        ? Math.Max(52D, CountSemanticListItems(block) * 30D)
                        : 52D;
                    if (block.Text.Length > 0 && NeedsGenericContinuation(contentTop, textHeight, slideBottom)) {
                        if (!TryAddGenericSlide(presentation, result, budget, out slide)) {
                            slideLimitReached = true;
                            break;
                        }
                        contentTop = pictureTop = 30D;
                    }
                    int previousTextBoxes = result.TextBoxes;
                    contentTop = ImportTextBox(block.SourceElement, block.Text, slide, contentTop, result, budget,
                        textHeight, block);
                    if (block.Kind == HtmlSemanticBlockKind.Form && result.TextBoxes > previousTextBoxes) {
                        AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ContentApproximated,
                            "An HTML form was imported as editable visible text without its interactive controls.",
                            lossKind: OfficeConversionLossKind.Approximation,
                            detail: "block=Form; preserved=visibleText; interaction=omitted");
                    }
                } else if (importTable) {
                    if (TryGetOversizedGenericTableText(block, budget, out string tableText)) {
                        AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ContentApproximated,
                            "A long single-cell HTML table was split into editable slide text; native cell structure and rich cell runs were not retained.",
                            lossKind: OfficeConversionLossKind.Approximation,
                            detail: "cellTextLength=" + tableText.Length + "; projection=paginatedText");
                        int omittedLinks = block.Table!.Rows[0].Cells[0].Runs.Count(run =>
                            !string.IsNullOrWhiteSpace(run.Hyperlink));
                        if (omittedLinks > 0) {
                            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ContentOmitted,
                                "Hyperlinks inside a paginated single-cell HTML table were not retained.",
                                lossKind: OfficeConversionLossKind.Omission,
                                detail: "hyperlinkRuns=" + omittedLinks);
                        }
                        foreach (string chunk in SplitGenericTableText(tableText)) {
                            if (NeedsGenericContinuation(contentTop, 130D, slideBottom)) {
                                if (!TryAddGenericSlide(presentation, result, budget, out slide)) {
                                    slideLimitReached = true;
                                    break;
                                }
                                contentTop = pictureTop = 30D;
                            }
                            contentTop = ImportTextBox(null, chunk, slide, contentTop, result, budget, 130D);
                        }
                    } else {
                        double tableHeight = Math.Max(90D, (block.Table?.Rows.Count ?? 1) * 40D);
                        if (NeedsGenericContinuation(contentTop, tableHeight, slideBottom)) {
                            if (!TryAddGenericSlide(presentation, result, budget, out slide)) {
                                slideLimitReached = true;
                                break;
                            }
                            contentTop = pictureTop = 30D;
                        }
                        contentTop = ImportTable(block.SourceElement, slide, contentTop, result, budget, block);
                    }
                } else if (importPicture) {
                    if (NeedsGenericContinuation(contentTop, 90D, slideBottom)) {
                        if (!TryAddGenericSlide(presentation, result, budget, out slide)) {
                            slideLimitReached = true;
                            break;
                        }
                        contentTop = pictureTop = 30D;
                    }
                    pictureTop = Math.Max(pictureTop, contentTop);
                    ImportPicture(block.SourceElement, slide, result, budget, ref pictureTop, fallbackLeft: 64D);
                    contentTop = Math.Max(contentTop, pictureTop);
                }
                if (slideLimitReached) break;
                if (options.ImportPictures) {
                    foreach (HtmlSemanticResource resource in EnumerateInlineResources(block)) {
                        GetGenericResourcePictureSize(resource, presentation, budget,
                            out _, out double imageHeight, out _);
                        if (HtmlImageDataUri.TryParse(resource.Source, out _)
                            && NeedsGenericContinuation(contentTop, imageHeight + 18D, slideBottom)) {
                            if (!TryAddGenericSlide(presentation, result, budget, out slide)) {
                                slideLimitReached = true;
                                break;
                            }
                            contentTop = pictureTop = 30D;
                        }
                        pictureTop = Math.Max(pictureTop, contentTop);
                        ImportSemanticResourcePicture(resource, slide, presentation, result, budget, ref pictureTop);
                        contentTop = Math.Max(contentTop, pictureTop);
                    }
                }
                if (slideLimitReached) break;
            }
            if (slideLimitReached) break;
        }

        if (editableLayout?.Regions.Count > 0) {
            ImportEditableLayoutRegions(editableLayout.Regions, firstSlideIndexes, presentation, options, result, budget);
        }
        ReportGenericOffSlideShapes(presentation, result, budget);
    }

    private static bool TryAddGenericSlide(
        PptCore.PowerPointPresentation presentation,
        HtmlToPowerPointResult result,
        HtmlImportBudget budget,
        out PptCore.PowerPointSlide slide) {
        if (!budget.TryReserveSemanticContainer(out string containerLimit)) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                "Additional generic HTML content was omitted because the shared slide limit was reached.",
                HtmlDiagnosticSeverity.Error, OfficeConversionLossKind.Omission, detail: containerLimit);
            slide = null!;
            return false;
        }
        slide = presentation.AddSlide();
        result.Slides++;
        return true;
    }

    private static bool NeedsGenericContinuation(double top, double height, double slideBottom) =>
        top > 30D && top + height > slideBottom;

    private static bool TryGetOversizedGenericTableText(
        HtmlSemanticBlock block,
        HtmlImportBudget budget,
        out string text) {
        text = string.Empty;
        if (block.Table?.Rows.Count != 1 || block.Table.Rows[0].Cells.Count != 1
            || block.SourceElement.HasAttribute("data-officeimo-left")
            || block.SourceElement.HasAttribute("data-officeimo-top")
            || block.SourceElement.HasAttribute("data-officeimo-width")
            || block.SourceElement.HasAttribute("data-officeimo-height")) return false;
        HtmlSemanticTableCell cell = block.Table.Rows[0].Cells[0];
        if (cell.RowSpan != 1 || cell.ColumnSpan != 1) return false;
        text = cell.Text;
        return text.Length > 300 && budget.IsMetadataWithinLimit(text, out _);
    }

    private static IEnumerable<string> SplitGenericTableText(string text) {
        const int maximumChunkLength = 180;
        int start = 0;
        while (start < text.Length) {
            while (start < text.Length && char.IsWhiteSpace(text[start])) start++;
            if (start >= text.Length) yield break;
            int end = Math.Min(text.Length, start + maximumChunkLength);
            if (end < text.Length) {
                int wordEnd = text.LastIndexOf(' ', end - 1, end - start);
                if (wordEnd > start + maximumChunkLength / 2) end = wordEnd;
            }
            yield return text.Substring(start, end - start).Trim();
            start = end;
        }
    }

    private static void GetGenericResourcePictureSize(
        HtmlSemanticResource resource,
        PptCore.PowerPointPresentation presentation,
        HtmlImportBudget budget,
        out double width,
        out double height,
        out bool fitted) {
        double maximum = budget.Limits.MaxAbsoluteGeometry;
        width = ReadGenericResourceDimension(resource.WidthPixels, 160D, maximum);
        height = ReadGenericResourceDimension(resource.HeightPixels, 90D, maximum);
        double widthLimit = Math.Max(1D, presentation.SlideSize.WidthPoints - 128D);
        double heightLimit = Math.Max(1D, presentation.SlideSize.HeightPoints - 60D);
        double scale = Math.Min(1D, Math.Min(widthLimit / width, heightLimit / height));
        fitted = scale < 1D;
        width *= scale;
        height *= scale;
    }

    private static double ReadGenericResourceDimension(double? pixels, double fallback, double maximum) {
        double value = pixels.GetValueOrDefault(fallback);
        if (!double.IsFinite(value) || value <= 0D) value = fallback;
        return Math.Min(maximum, Math.Max(1D, value * 0.75D));
    }

    private static void ReportGenericOffSlideShapes(
        PptCore.PowerPointPresentation presentation,
        HtmlToPowerPointResult result,
        HtmlImportBudget budget) {
        var options = new PptCore.PowerPointDeckPreflightOptions {
            MaximumShapeCount = budget.Limits.MaxShapes,
            DetectTextOverflow = false,
            DetectUnreadableFontReduction = false,
            DetectShapeCollisions = false,
            DetectMissingVisualAssets = false,
            IncludeVisualSnapshotDiagnostics = false
        };
        foreach (PptCore.PowerPointDeckPreflightFinding finding in presentation.InspectPreflight(options).Findings
                     .Where(item => item.Code == "Layout.ShapeOffSlide")) {
            PptCore.PowerPointLayoutBox? bounds = finding.Bounds;
            bool whollyOutside = bounds.HasValue && (bounds.Value.Right <= 0L || bounds.Value.Bottom <= 0L
                || bounds.Value.Left >= presentation.SlideSize.WidthEmus
                || bounds.Value.Top >= presentation.SlideSize.HeightEmus);
            AddImportDiagnostic(result,
                whollyOutside ? HtmlConversionDiagnosticCodes.ContentOmitted : HtmlConversionDiagnosticCodes.ContentApproximated,
                whollyOutside
                    ? "A generic HTML slide shape lies outside the visible slide."
                    : "A generic HTML slide shape extends beyond the visible slide and may be clipped.",
                lossKind: whollyOutside ? OfficeConversionLossKind.Omission : OfficeConversionLossKind.Approximation,
                detail: "slide=" + (finding.SlideIndex + 1) + "; shape=" + (finding.ShapeIndex.GetValueOrDefault() + 1));
        }
    }

    private static void ImportEditableLayoutRegions(
        IReadOnlyList<HtmlRenderLayoutRegion> regions,
        IReadOnlyList<int> firstSlideIndexes,
        PptCore.PowerPointPresentation presentation,
        HtmlToPowerPointOptions options,
        HtmlToPowerPointResult result,
        HtmlImportBudget budget) {
        if (presentation.Slides.Count == 0) {
            presentation.AddSlide();
            result.Slides++;
        }

        foreach (IGrouping<int, HtmlRenderLayoutRegion> sectionGroup in regions
                     .GroupBy(region => region.SemanticSectionNumber)
                     .OrderBy(group => group.Key)) {
            int slideIndex = sectionGroup.Key > 0 && sectionGroup.Key <= firstSlideIndexes.Count
                ? firstSlideIndexes[sectionGroup.Key - 1]
                : firstSlideIndexes.Count == 0 && sectionGroup.Key == 1 ? 0 : -1;
            if (slideIndex < 0 || slideIndex >= presentation.Slides.Count) {
                foreach (HtmlRenderLayoutRegion region in sectionGroup) {
                    AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                        "An editable HTML layout region was omitted because its owning semantic slide was not created.",
                        HtmlDiagnosticSeverity.Error, OfficeConversionLossKind.Omission, region.Source,
                        "semanticSection=" + sectionGroup.Key + "; slides=" + presentation.Slides.Count);
                }
                continue;
            }

            PptCore.PowerPointSlide slide = presentation.Slides[slideIndex];
            int semanticShapeCount = slide.Shapes.Count;
            var negativeRegionShapes = new List<PptCore.PowerPointShape>();
            double maximumGeometry = budget.Limits.MaxAbsoluteGeometry;
            IReadOnlyList<HtmlRenderLayoutRegion> orderedRegions = sectionGroup
                .OrderBy(item => item.PaintOrder)
                .ToList();
            var shapeReservations = new Dictionary<HtmlRenderLayoutRegion, HtmlImportBudgetReservation>();
            var shapeReservationFailures = new Dictionary<HtmlRenderLayoutRegion, string>();
            foreach (HtmlRenderLayoutRegion region in orderedRegions) {
                if (!budget.IsMetadataWithinLimit(region.SourceText, out _)) continue;
                if (budget.TryReserveShape(out HtmlImportBudgetReservation reservation, out string detail)) {
                    shapeReservations[region] = reservation;
                } else {
                    shapeReservationFailures[region] = detail;
                }
            }
            var occupied = slide.TextBoxes
                .Select(box => new EditableLayoutSlideBounds(box.LeftPoints, box.TopPoints, box.WidthPoints, box.HeightPoints))
                .Concat(slide.Pictures.Select(picture => new EditableLayoutSlideBounds(
                    picture.LeftPoints, picture.TopPoints, picture.WidthPoints, picture.HeightPoints)))
                .Concat(slide.Tables.Select(table => new EditableLayoutSlideBounds(
                    table.LeftPoints, table.TopPoints, table.WidthPoints, table.HeightPoints)))
                .Concat(sectionGroup.Where(region =>
                        region.RegionKind == HtmlRenderLayoutRegionKind.Positioned
                        && shapeReservations.ContainsKey(region))
                    .Select(region => CreateBoundedCollisionBounds(region, maximumGeometry)))
                .ToList();

            for (int regionIndex = 0; regionIndex < orderedRegions.Count; regionIndex++) {
                HtmlRenderLayoutRegion region = orderedRegions[regionIndex];
                if (!budget.IsMetadataWithinLimit(region.SourceText, out string metadataLimit)) {
                    AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                        "An editable HTML layout region was omitted because its text exceeded the shared metadata limit.",
                        HtmlDiagnosticSeverity.Error, OfficeConversionLossKind.Omission, region.Source, metadataLimit);
                    continue;
                }
                if (!shapeReservations.TryGetValue(region, out HtmlImportBudgetReservation? shapeReservation)) {
                    AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                        "An editable HTML layout region was omitted because the native shape limit was reached.",
                        HtmlDiagnosticSeverity.Error, OfficeConversionLossKind.Omission, region.Source,
                        shapeReservationFailures.TryGetValue(region, out string? shapeLimit)
                            ? shapeLimit
                            : nameof(HtmlImportLimits.MaxShapes));
                    continue;
                }
                using HtmlImportBudgetReservation reservedShape = shapeReservation!;
                double localRegionX = region.X - region.SemanticSectionOriginX;
                double localRegionY = region.Y - region.SemanticSectionOriginY;
                double left = NormalizeGeometry(localRegionX * 0.75D, 0D, -maximumGeometry,
                    budget, result, "editable layout region left");
                double top = NormalizeGeometry(localRegionY * 0.75D, 0D, -maximumGeometry,
                    budget, result, "editable layout region top");
                double width = NormalizeGeometry(region.Width * 0.75D, 1D, 1D,
                    budget, result, "editable layout region width");
                double height = NormalizeGeometry(region.Height * 0.75D, 1D, 1D,
                    budget, result, "editable layout region height");
                double requestedTop = top;
                var bounds = new EditableLayoutSlideBounds(left, top, width, height);
                bool placementAvailable = true;
                if (region.RegionKind != HtmlRenderLayoutRegionKind.Positioned) {
                    while (occupied.Any(existing => existing.Intersects(bounds))) {
                        double nextTop = occupied.Where(existing => existing.Intersects(bounds))
                            .Max(existing => existing.Bottom) + 8D;
                        if (nextTop > maximumGeometry) {
                            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                                "An editable HTML layout region was omitted because no bounded non-overlapping slide position remained.",
                                HtmlDiagnosticSeverity.Error, OfficeConversionLossKind.Omission, region.Source,
                                "MaxAbsoluteGeometry=" + maximumGeometry.ToString(
                                    System.Globalization.CultureInfo.InvariantCulture));
                            placementAvailable = false;
                            break;
                        }
                        top = nextTop;
                        bounds = new EditableLayoutSlideBounds(left, top, width, height);
                    }
                    if (placementAvailable && Math.Abs(top - requestedTop) > 0.01D) {
                        AddImportDiagnostic(result, HtmlEditableLayoutDiagnosticCodes.PlacementSimplified,
                            "PowerPoint moved an in-flow editable layout region below existing native slide content.",
                            lossKind: OfficeConversionLossKind.Approximation, source: region.Source,
                            detail: "requestedTop=" + requestedTop.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)
                                + "; actualTop=" + top.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture));
                    }
                }
                if (!placementAvailable) {
                    shapeReservation.Dispose();
                    shapeReservations.Remove(region);
                    RetryShapeReservationsAfterRollback(
                        orderedRegions,
                        regionIndex + 1,
                        shapeReservations,
                        shapeReservationFailures,
                        budget,
                        occupied,
                        maximumGeometry);
                    continue;
                }
                double topOffset = top - requestedTop;
                var nativeRegionShapes = new List<PptCore.PowerPointShape>();
                bool hasBackgroundPicture = options.ImportPictures
                    && region.BackgroundColor.HasValue
                    && HtmlEditableLayoutProjector.EnumeratePictures(region.Visuals, includeBackgroundImages: true)
                        .Any(item => item.IsBackground);
                bool importBackgroundPictures = options.ImportPictures;
                HtmlImportBudgetReservation? backgroundFillReservation = null;
                if (hasBackgroundPicture) {
                    if (budget.TryReserveShape(out HtmlImportBudgetReservation reservation, out string fillLimit)) {
                        backgroundFillReservation = reservation;
                    } else {
                        importBackgroundPictures = false;
                        AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                            "PowerPoint omitted editable background pictures because the separate native backing fill exceeded the shape limit.",
                            lossKind: OfficeConversionLossKind.Omission, source: region.Source, detail: fillLimit);
                    }
                }

                using HtmlImportBudgetReservation? backgroundFillReservationScope = backgroundFillReservation;
                if (backgroundFillReservation != null) {
                    PptCore.PowerPointAutoShape backingFill = slide.AddRectanglePoints(
                        left, top, width, height, "HTML background " + region.SourceKey);
                    backingFill.FillColor = region.BackgroundColor!.Value.ToRgbHex();
                    backgroundFillReservation.Commit();
                    nativeRegionShapes.Add(backingFill);
                }
                int retainedBackgroundPictures = 0;
                if (importBackgroundPictures) {
                    retainedBackgroundPictures = AddEditableLayoutPictures(slide, region, backgroundImages: true,
                        left, top, topOffset, maximumGeometry, budget, result, nativeRegionShapes);
                }
                PptCore.PowerPointTextBox textBox = slide.AddTextBoxPoints(region.SourceText, left, top, width, height);
                reservedShape.Commit();
                nativeRegionShapes.Add(textBox);
                textBox.Name = "HTML " + region.RegionKind + " " + region.SourceKey;
                if (backgroundFillReservation != null) textBox.FillTransparency = 100;
                else if (region.BackgroundColor.HasValue) textBox.FillColor = region.BackgroundColor.Value.ToRgbHex();
                else textBox.FillTransparency = 100;

                if (options.ImportPictures) {
                    AddEditableLayoutPictures(slide, region, backgroundImages: false,
                        left, top, topOffset, maximumGeometry, budget, result, nativeRegionShapes);
                }

                if (region.BoxShadowLayerCount > 0) {
                    textBox.SetShadow("000000", blurPoints: 4D, distancePoints: 2D, angleDegrees: 45D, transparencyPercent: 45);
                    AddImportDiagnostic(result, HtmlEditableLayoutDiagnosticCodes.EffectUnsupported,
                        region.BoxShadowLayerCount > 1
                            ? "PowerPoint approximated the first editable CSS shadow and omitted additional shadow layers."
                            : "PowerPoint approximated the editable CSS shadow with one native outer shadow.",
                        lossKind: OfficeConversionLossKind.Approximation, source: region.Source,
                        detail: "shadowLayers=" + region.BoxShadowLayerCount
                            + "; nativeShadowParameters=approximated");
                }
                if (region.BackgroundLayerCount > 0) {
                    int omittedBackgroundLayers = Math.Max(0, region.BackgroundLayerCount - retainedBackgroundPictures);
                    string backgroundLayerMessage = omittedBackgroundLayers == 0
                        ? "PowerPoint retained supported background images as native pictures and used the editable text-box fill for the region background."
                        : retainedBackgroundPictures > 0
                            ? "PowerPoint retained supported background images as native pictures but omitted other background layers without a native editable representation."
                            : "PowerPoint omitted background layers without a native editable representation.";
                    AddImportDiagnostic(result, HtmlEditableLayoutDiagnosticCodes.BackgroundLayersFlattened,
                        backgroundLayerMessage,
                        omittedBackgroundLayers > 0 ? HtmlDiagnosticSeverity.Warning : HtmlDiagnosticSeverity.Info,
                        omittedBackgroundLayers > 0 ? OfficeConversionLossKind.Omission : OfficeConversionLossKind.None,
                        region.Source,
                        "backgroundLayers=" + region.BackgroundLayerCount
                            + "; retainedNativePictures=" + retainedBackgroundPictures
                            + "; omittedLayers=" + omittedBackgroundLayers);
                }
                if (region.ZIndex < 0) {
                    negativeRegionShapes.AddRange(nativeRegionShapes);
                } else if (semanticShapeCount > 0) {
                    AddImportDiagnostic(result, HtmlEditableLayoutDiagnosticCodes.PlacementSimplified,
                        "PowerPoint appended an editable layout region above semantic slide content because exact mixed-flow stacking has no native mapping.",
                        lossKind: OfficeConversionLossKind.Approximation, source: region.Source,
                        detail: "stacking=appended-after-semantic-content; zIndex=" + region.ZIndex
                            + "; paintOrder=" + region.PaintOrder);
                }
                result.TextBoxes++;
                occupied.Add(bounds);
            }
            for (int shapeIndex = negativeRegionShapes.Count - 1; shapeIndex >= 0; shapeIndex--) {
                slide.SendToBack(negativeRegionShapes[shapeIndex]);
            }
        }
    }

    private static void RetryShapeReservationsAfterRollback(
        IReadOnlyList<HtmlRenderLayoutRegion> orderedRegions,
        int startIndex,
        IDictionary<HtmlRenderLayoutRegion, HtmlImportBudgetReservation> reservations,
        IDictionary<HtmlRenderLayoutRegion, string> failures,
        HtmlImportBudget budget,
        ICollection<EditableLayoutSlideBounds> occupied,
        double maximumGeometry) {
        for (int index = startIndex; index < orderedRegions.Count; index++) {
            HtmlRenderLayoutRegion candidate = orderedRegions[index];
            if (reservations.ContainsKey(candidate)
                || !budget.IsMetadataWithinLimit(candidate.SourceText, out _)) {
                continue;
            }
            if (!budget.TryReserveShape(out HtmlImportBudgetReservation reservation, out string detail)) {
                failures[candidate] = detail;
                continue;
            }
            reservations[candidate] = reservation;
            failures.Remove(candidate);
            if (candidate.RegionKind == HtmlRenderLayoutRegionKind.Positioned) {
                occupied.Add(CreateBoundedCollisionBounds(candidate, maximumGeometry));
            }
            return;
        }
    }

    private static int AddEditableLayoutPictures(
        PptCore.PowerPointSlide slide,
        HtmlRenderLayoutRegion region,
        bool backgroundImages,
        double left,
        double top,
        double topOffset,
        double maximumGeometry,
        HtmlImportBudget budget,
        HtmlToPowerPointResult result,
        ICollection<PptCore.PowerPointShape> nativeRegionShapes) {
        int retainedPictures = 0;
        foreach (HtmlEditableLayoutPicture image in
                 HtmlEditableLayoutProjector.EnumeratePictures(region.Visuals, includeBackgroundImages: true)
                     .Where(item => item.IsBackground == backgroundImages)) {
            if (!TryGetImagePartType(image.ContentType, out OfficeImageFormat imageType)) {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ResourceTypeUnsupported,
                    "A layout-region picture used an unsupported native PowerPoint image type.",
                    lossKind: OfficeConversionLossKind.Omission, source: image.Source);
                continue;
            }
            if (!budget.TryReserveImageWithShape(image.Bytes.LongLength,
                    out HtmlImportBudgetReservation imageReservation, out string imageLimit)) {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                    "A layout-region picture was omitted because the shared image or shape limit was reached.",
                    lossKind: OfficeConversionLossKind.Omission, source: image.Source, detail: imageLimit);
                continue;
            }
            using HtmlImportBudgetReservation imageReservationScope = imageReservation;
            double pictureLeft = NormalizeGeometry(
                (image.X - region.SemanticSectionOriginX) * 0.75D, left, -maximumGeometry,
                budget, result, "editable layout picture left");
            double pictureTop = NormalizeGeometry(
                (image.Y - region.SemanticSectionOriginY) * 0.75D + topOffset, top, -maximumGeometry,
                budget, result, "editable layout picture top");
            double pictureWidth = NormalizeGeometry(image.Width * 0.75D, 1D, 1D,
                budget, result, "editable layout picture width");
            double pictureHeight = NormalizeGeometry(image.Height * 0.75D, 1D, 1D,
                budget, result, "editable layout picture height");
            using var stream = new MemoryStream(image.Bytes);
            PptCore.PowerPointPicture picture = slide.AddPicturePoints(stream, imageType,
                pictureLeft, pictureTop, pictureWidth, pictureHeight);
            nativeRegionShapes.Add(picture);
            if (!string.IsNullOrWhiteSpace(image.AlternativeText)) picture.AltText = image.AlternativeText;
            if (image.Opacity < 0.999D) picture.FillTransparency = (int)Math.Round((1D - image.Opacity) * 100D);
            if (image.SourceCrop.HasCrop) {
                picture.Crop(
                    image.SourceCrop.Left * 100D,
                    image.SourceCrop.Top * 100D,
                    image.SourceCrop.Right * 100D,
                    image.SourceCrop.Bottom * 100D);
            }
            result.Pictures++;
            imageReservation.Commit();
            retainedPictures++;
        }
        return retainedPictures;
    }

    private static EditableLayoutSlideBounds CreateBoundedCollisionBounds(
        HtmlRenderLayoutRegion region,
        double maximumGeometry) {
        double left = Math.Max(-maximumGeometry, Math.Min(maximumGeometry,
            (region.X - region.SemanticSectionOriginX) * 0.75D));
        double top = Math.Max(-maximumGeometry, Math.Min(maximumGeometry,
            (region.Y - region.SemanticSectionOriginY) * 0.75D));
        double width = Math.Max(1D, Math.Min(maximumGeometry, region.Width * 0.75D));
        double height = Math.Max(1D, Math.Min(maximumGeometry, region.Height * 0.75D));
        return new EditableLayoutSlideBounds(left, top, width, height);
    }

    private readonly struct EditableLayoutSlideBounds {
        internal EditableLayoutSlideBounds(double left, double top, double width, double height) {
            Left = left;
            Top = top;
            Width = width;
            Height = height;
        }

        internal double Left { get; }
        internal double Top { get; }
        internal double Width { get; }
        internal double Height { get; }
        internal double Right => Left + Width;
        internal double Bottom => Top + Height;

        internal bool Intersects(EditableLayoutSlideBounds other) =>
            Left < other.Right && Right > other.Left && Top < other.Bottom && Bottom > other.Top;
    }

    private static IEnumerable<HtmlSemanticResource> EnumerateInlineResources(HtmlSemanticBlock block) {
        foreach (HtmlSemanticResource resource in block.InlineResources.Where(item => item.Kind == HtmlResourceKind.Image)) yield return resource;
        if (block.Table != null) {
            foreach (HtmlSemanticResource resource in block.Table.Rows.SelectMany(row => row.Cells)
                .SelectMany(cell => cell.Resources).Where(item => item.Kind == HtmlResourceKind.Image)) yield return resource;
        }
        foreach (HtmlSemanticBlock child in block.Children) {
            foreach (HtmlSemanticResource resource in EnumerateInlineResources(child)) yield return resource;
        }
    }

    private static void ImportSemanticResourcePicture(
        HtmlSemanticResource resource,
        PptCore.PowerPointSlide slide,
        PptCore.PowerPointPresentation presentation,
        HtmlToPowerPointResult result,
        HtmlImportBudget budget,
        ref double top) {
        if (!HtmlImageDataUri.TryParse(resource.Source, out HtmlImageDataUri dataUri)
            || !TryGetImagePartType(dataUri.MediaType, out OfficeImageFormat imagePartType)) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ResourceTypeUnsupported,
                "An inline generic slide image was omitted because native import requires a supported bounded image data URI.",
                lossKind: OfficeConversionLossKind.Omission, source: resource.Source);
            return;
        }
        if (!budget.TryReserveImageWithShape(dataUri, out HtmlImportBudgetReservation imageReservation, out string limit)) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                "An inline generic slide image was omitted because the shared image or shape limit was reached.",
                lossKind: OfficeConversionLossKind.Omission, source: resource.Source, detail: limit);
            return;
        }
        using HtmlImportBudgetReservation imageReservationScope = imageReservation;
        if (!dataUri.TryDecodeBytes(out byte[] bytes)) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ResourceDecodeFailed,
                "An inline generic slide image could not be decoded.",
                lossKind: OfficeConversionLossKind.Omission, source: resource.Source);
            return;
        }
        GetGenericResourcePictureSize(resource, presentation, budget,
            out double width, out double height, out bool fitted);
        using var stream = new MemoryStream(bytes);
        PptCore.PowerPointPicture picture = slide.AddPicturePoints(stream, imagePartType, 64D, top, width, height);
        if (!string.IsNullOrWhiteSpace(resource.AlternateText)) picture.AltText = resource.AlternateText;
        result.Pictures++;
        imageReservation.Commit();
        top += height + 18D;
        if (fitted) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ContentApproximated,
                "An inline generic slide image was proportionally fitted to the slide canvas.",
                lossKind: OfficeConversionLossKind.Approximation, source: resource.Source,
                detail: "fittedWidthPoints=" + width.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)
                    + "; fittedHeightPoints=" + height.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture));
        }
    }

    private static bool IsGenericTextBlock(HtmlSemanticBlockKind kind) =>
        kind == HtmlSemanticBlockKind.Heading || kind == HtmlSemanticBlockKind.Paragraph
        || kind == HtmlSemanticBlockKind.Code || kind == HtmlSemanticBlockKind.Quote
        || kind == HtmlSemanticBlockKind.List || kind == HtmlSemanticBlockKind.Note
        || kind == HtmlSemanticBlockKind.Form;

    private static int CountSemanticListItems(HtmlSemanticBlock list) =>
        list.Children.Sum(item => 1 + item.Children
            .Where(child => child.Kind == HtmlSemanticBlockKind.List)
            .Sum(CountSemanticListItems));
}
