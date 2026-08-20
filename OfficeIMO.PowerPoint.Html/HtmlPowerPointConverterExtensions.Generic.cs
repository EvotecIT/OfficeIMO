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
            ImportEditableLayoutRegions(editableLayout.Regions, presentation, result, budget);
        }
    }

    private static void ImportEditableLayoutRegions(
        IReadOnlyList<HtmlRenderLayoutRegion> regions,
        PptCore.PowerPointPresentation presentation,
        HtmlToPowerPointResult result,
        HtmlImportBudget budget) {
        PptCore.PowerPointSlide slide;
        if (presentation.Slides.Count == 0) {
            slide = presentation.AddSlide();
            result.Slides++;
        } else {
            slide = presentation.Slides[0];
        }

        var occupied = slide.TextBoxes
            .Select(box => new EditableLayoutSlideBounds(box.LeftPoints, box.TopPoints, box.WidthPoints, box.HeightPoints))
            .Concat(slide.Pictures.Select(picture => new EditableLayoutSlideBounds(
                picture.LeftPoints, picture.TopPoints, picture.WidthPoints, picture.HeightPoints)))
            .Concat(slide.Tables.Select(table => new EditableLayoutSlideBounds(
                table.LeftPoints, table.TopPoints, table.WidthPoints, table.HeightPoints)))
            .Concat(regions.Where(region => region.RegionKind == HtmlRenderLayoutRegionKind.Positioned)
                .Select(region => new EditableLayoutSlideBounds(
                    region.X * 0.75D, region.Y * 0.75D,
                    Math.Max(1D, region.Width * 0.75D), Math.Max(1D, region.Height * 0.75D))))
            .ToList();

        foreach (HtmlRenderLayoutRegion region in regions.OrderBy(item => item.PaintOrder)) {
            if (!budget.TryReserveShape(out string shapeLimit)) {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                    "An editable HTML layout region was omitted because the native shape limit was reached.",
                    HtmlDiagnosticSeverity.Error, OfficeConversionLossKind.Omission, region.Source, shapeLimit);
                continue;
            }
            double left = region.X * 0.75D;
            double top = region.Y * 0.75D;
            double width = Math.Max(1D, region.Width * 0.75D);
            double height = Math.Max(1D, region.Height * 0.75D);
            double requestedTop = top;
            var bounds = new EditableLayoutSlideBounds(left, top, width, height);
            if (region.RegionKind != HtmlRenderLayoutRegionKind.Positioned) {
                while (occupied.Any(existing => existing.Intersects(bounds))) {
                    top = occupied.Where(existing => existing.Intersects(bounds))
                        .Max(existing => existing.Bottom) + 8D;
                    bounds = new EditableLayoutSlideBounds(left, top, width, height);
                }
                if (Math.Abs(top - requestedTop) > 0.01D) {
                    AddImportDiagnostic(result, HtmlEditableLayoutDiagnosticCodes.PlacementSimplified,
                        "PowerPoint moved an in-flow editable layout region below existing native slide content.",
                        lossKind: OfficeConversionLossKind.Approximation, source: region.Source,
                        detail: "requestedTop=" + requestedTop.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)
                            + "; actualTop=" + top.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture));
                }
            }
            double topOffset = top - requestedTop;

            foreach ((HtmlRenderImage Image, double Opacity) image in EnumerateLayoutImages(region.Visuals, 1D)) {
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
                using var stream = new MemoryStream(image.Image.Bytes);
                PptCore.PowerPointPicture picture = slide.AddPicturePoints(stream, imageType,
                    image.Image.X * 0.75D, image.Image.Y * 0.75D + topOffset,
                    image.Image.Width * 0.75D, image.Image.Height * 0.75D);
                if (!string.IsNullOrWhiteSpace(image.Image.AlternativeText)) picture.AltText = image.Image.AlternativeText;
                if (image.Opacity < 0.999D) picture.FillTransparency = (int)Math.Round((1D - image.Opacity) * 100D);
                result.Pictures++;
                imageReservation.Commit();
            }

            PptCore.PowerPointTextBox textBox = slide.AddTextBoxPoints(region.SourceText, left, top, width, height);
            textBox.Name = "HTML " + region.RegionKind + " " + region.SourceKey;
            if (region.BackgroundColor.HasValue) textBox.FillColor = region.BackgroundColor.Value.ToRgbHex();
            else textBox.FillTransparency = 100;
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
            result.TextBoxes++;
            occupied.Add(bounds);
        }
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

    private static IEnumerable<(HtmlRenderImage Image, double Opacity)> EnumerateLayoutImages(
        IEnumerable<HtmlRenderVisual> visuals,
        double opacity) {
        foreach (HtmlRenderVisual visual in visuals) {
            if (visual is HtmlRenderImage image) yield return (image, opacity);
            IEnumerable<HtmlRenderVisual>? children = visual switch {
                HtmlRenderEffectGroup effect => effect.Visuals,
                HtmlRenderLayoutRegion region => region.Visuals,
                HtmlRenderSemanticGroup semantic => semantic.Visuals,
                HtmlRenderLogicalTextGroup logical => logical.Visuals,
                HtmlRenderClipGroup clip => clip.Visuals,
                HtmlRenderPathClipGroup pathClip => pathClip.Visuals,
                _ => null
            };
            if (children == null) continue;
            double childOpacity = visual is HtmlRenderEffectGroup group ? opacity * group.Opacity : opacity;
            foreach ((HtmlRenderImage Image, double Opacity) child in EnumerateLayoutImages(children, childOpacity)) yield return child;
        }
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
