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
        foreach (HtmlSemanticSection section in document.Sections) {
            if (!budget.TryReserveSemanticContainer(out string containerLimit)) {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                    "Additional HTML sections were omitted because the shared slide limit was reached.",
                    HtmlDiagnosticSeverity.Error, OfficeConversionLossKind.Omission, detail: containerLimit);
                break;
            }

            PptCore.PowerPointSlide slide = presentation.AddSlide();
            result.Slides++;
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
                    contentTop = ImportTextBox(block.SourceElement, block.Text, slide, contentTop, result, budget,
                        block.Kind == HtmlSemanticBlockKind.List ? Math.Max(52D, CountSemanticListItems(block) * 30D) : 52D,
                        block);
                } else if (importTable) {
                    contentTop = ImportTable(block.SourceElement, slide, contentTop, result, budget, block);
                } else if (importPicture) {
                    pictureTop = Math.Max(pictureTop, contentTop);
                    ImportPicture(block.SourceElement, slide, result, budget, ref pictureTop);
                    contentTop = Math.Max(contentTop, pictureTop);
                }
                if (options.ImportPictures) {
                    foreach (HtmlSemanticResource resource in EnumerateInlineResources(block)) {
                        pictureTop = Math.Max(pictureTop, contentTop);
                        ImportSemanticResourcePicture(resource, slide, result, budget, ref pictureTop);
                        contentTop = Math.Max(contentTop, pictureTop);
                    }
                }
            }
        }

        if (editableLayout?.Regions.Count > 0) {
            ImportEditableLayoutRegions(editableLayout.Regions, presentation, options, result, budget);
        }
    }

    private static void ImportEditableLayoutRegions(
        IReadOnlyList<HtmlRenderLayoutRegion> regions,
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
            int slideIndex = sectionGroup.Key - 1;
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
            var occupied = slide.TextBoxes
                .Select(box => new EditableLayoutSlideBounds(box.LeftPoints, box.TopPoints, box.WidthPoints, box.HeightPoints))
                .Concat(slide.Pictures.Select(picture => new EditableLayoutSlideBounds(
                    picture.LeftPoints, picture.TopPoints, picture.WidthPoints, picture.HeightPoints)))
                .Concat(slide.Tables.Select(table => new EditableLayoutSlideBounds(
                    table.LeftPoints, table.TopPoints, table.WidthPoints, table.HeightPoints)))
                .Concat(sectionGroup.Where(region => region.RegionKind == HtmlRenderLayoutRegionKind.Positioned)
                    .Select(region => CreateBoundedCollisionBounds(region, maximumGeometry)))
                .ToList();

            foreach (HtmlRenderLayoutRegion region in sectionGroup.OrderBy(item => item.PaintOrder)) {
                if (!budget.IsMetadataWithinLimit(region.SourceText, out string metadataLimit)) {
                    AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                        "An editable HTML layout region was omitted because its text exceeded the shared metadata limit.",
                        HtmlDiagnosticSeverity.Error, OfficeConversionLossKind.Omission, region.Source, metadataLimit);
                    continue;
                }
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
                if (!placementAvailable) continue;
                if (!budget.TryReserveShape(out string shapeLimit)) {
                    AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                        "An editable HTML layout region was omitted because the native shape limit was reached.",
                        HtmlDiagnosticSeverity.Error, OfficeConversionLossKind.Omission, region.Source, shapeLimit);
                    continue;
                }
                double topOffset = top - requestedTop;
                var nativeRegionShapes = new List<PptCore.PowerPointShape>();
                if (options.ImportPictures) {
                    AddEditableLayoutPictures(slide, region, backgroundImages: true,
                        left, top, topOffset, maximumGeometry, budget, result, nativeRegionShapes);
                }
                PptCore.PowerPointTextBox textBox = slide.AddTextBoxPoints(region.SourceText, left, top, width, height);
                nativeRegionShapes.Add(textBox);
                textBox.Name = "HTML " + region.RegionKind + " " + region.SourceKey;
                if (region.BackgroundColor.HasValue) textBox.FillColor = region.BackgroundColor.Value.ToRgbHex();
                else textBox.FillTransparency = 100;

                if (options.ImportPictures) {
                    AddEditableLayoutPictures(slide, region, backgroundImages: false,
                        left, top, topOffset, maximumGeometry, budget, result, nativeRegionShapes);
                }

                if (region.BoxShadowLayerCount > 0) {
                    textBox.SetShadow("000000", blurPoints: 4D, distancePoints: 2D, angleDegrees: 45D, transparencyPercent: 45);
                    if (region.BoxShadowLayerCount > 1) {
                        AddImportDiagnostic(result, HtmlEditableLayoutDiagnosticCodes.EffectUnsupported,
                            "PowerPoint retained the first editable shadow and omitted additional CSS shadow layers.",
                            lossKind: OfficeConversionLossKind.Approximation, source: region.Source,
                            detail: "shadowLayers=" + region.BoxShadowLayerCount);
                    }
                }
                if (region.BackgroundLayerCount > 0) {
                    AddImportDiagnostic(result, HtmlEditableLayoutDiagnosticCodes.BackgroundLayersFlattened,
                        "PowerPoint retained supported image layers as native pictures and used the editable text-box fill for the region background.",
                        HtmlDiagnosticSeverity.Info, source: region.Source,
                        detail: "backgroundLayers=" + region.BackgroundLayerCount);
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

    private static void AddEditableLayoutPictures(
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
        foreach ((HtmlRenderImage Image, double Opacity) image in
                 HtmlEditableLayoutProjector.EnumerateImages(region.Visuals, includeBackgroundImages: true)
                     .Where(item => HtmlEditableLayoutProjector.IsBackgroundImage(item.Image) == backgroundImages)) {
            if (!TryGetImagePartType(image.Image.ContentType, out OfficeImageFormat imageType)) {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ResourceTypeUnsupported,
                    "A layout-region picture used an unsupported native PowerPoint image type.",
                    lossKind: OfficeConversionLossKind.Omission, source: image.Image.Source);
                continue;
            }
            if (!budget.TryReserveImageWithShape(image.Image.Bytes.LongLength,
                    out HtmlImportBudgetReservation imageReservation, out string imageLimit)) {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                    "A layout-region picture was omitted because the shared image or shape limit was reached.",
                    lossKind: OfficeConversionLossKind.Omission, source: image.Image.Source, detail: imageLimit);
                continue;
            }
            using HtmlImportBudgetReservation imageReservationScope = imageReservation;
            double pictureLeft = NormalizeGeometry(
                (image.Image.X - region.SemanticSectionOriginX) * 0.75D, left, -maximumGeometry,
                budget, result, "editable layout picture left");
            double pictureTop = NormalizeGeometry(
                (image.Image.Y - region.SemanticSectionOriginY) * 0.75D + topOffset, top, -maximumGeometry,
                budget, result, "editable layout picture top");
            double pictureWidth = NormalizeGeometry(image.Image.Width * 0.75D, 1D, 1D,
                budget, result, "editable layout picture width");
            double pictureHeight = NormalizeGeometry(image.Image.Height * 0.75D, 1D, 1D,
                budget, result, "editable layout picture height");
            using var stream = new MemoryStream(image.Image.Bytes);
            PptCore.PowerPointPicture picture = slide.AddPicturePoints(stream, imageType,
                pictureLeft, pictureTop, pictureWidth, pictureHeight);
            nativeRegionShapes.Add(picture);
            if (!string.IsNullOrWhiteSpace(image.Image.AlternativeText)) picture.AltText = image.Image.AlternativeText;
            if (image.Opacity < 0.999D) picture.FillTransparency = (int)Math.Round((1D - image.Opacity) * 100D);
            if (image.Image.SourceCrop.HasCrop) {
                picture.Crop(
                    image.Image.SourceCrop.Left * 100D,
                    image.Image.SourceCrop.Top * 100D,
                    image.Image.SourceCrop.Right * 100D,
                    image.Image.SourceCrop.Bottom * 100D);
            }
            result.Pictures++;
            imageReservation.Commit();
        }
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
        double maximum = budget.Limits.MaxAbsoluteGeometry;
        double width = Math.Min(maximum, Math.Max(1D, (resource.WidthPixels ?? 160D) * 0.75D));
        double height = Math.Min(maximum, Math.Max(1D, (resource.HeightPixels ?? 90D) * 0.75D));
        using var stream = new MemoryStream(bytes);
        PptCore.PowerPointPicture picture = slide.AddPicturePoints(stream, imagePartType, 64D, top, width, height);
        if (!string.IsNullOrWhiteSpace(resource.AlternateText)) picture.AltText = resource.AlternateText;
        result.Pictures++;
        imageReservation.Commit();
        top += height + 18D;
    }

    private static bool IsGenericTextBlock(HtmlSemanticBlockKind kind) =>
        kind == HtmlSemanticBlockKind.Heading || kind == HtmlSemanticBlockKind.Paragraph
        || kind == HtmlSemanticBlockKind.Code || kind == HtmlSemanticBlockKind.Quote
        || kind == HtmlSemanticBlockKind.List || kind == HtmlSemanticBlockKind.Note;

    private static int CountSemanticListItems(HtmlSemanticBlock list) =>
        list.Children.Sum(item => 1 + item.Children
            .Where(child => child.Kind == HtmlSemanticBlockKind.List)
            .Sum(CountSemanticListItems));
}
