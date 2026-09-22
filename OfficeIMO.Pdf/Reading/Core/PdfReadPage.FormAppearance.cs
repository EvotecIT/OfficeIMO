using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    internal WidgetAppearanceScanBudget CreateWidgetAppearanceScanBudget() => new WidgetAppearanceScanBudget(this);

    internal bool DoesWidgetNormalAppearancePresentAllText(
        int widgetObjectNumber,
        IReadOnlyList<string> expectedValues,
        WidgetAppearanceScanBudget budget,
        double maximumTinyFontSizePoints) {
        if (!_objects.TryGetValue(widgetObjectNumber, out PdfIndirectObject? widgetObject) ||
            widgetObject.Value is not PdfDictionary widget ||
            !TryGetNormalAppearanceStream(widget, out PdfStream appearance)) {
            return false;
        }

        try {
            var spans = new List<PdfTextSpan>();
            if (!TryReadBox(
                    appearance.Dictionary.Items.TryGetValue("BBox", out PdfObject? bboxObject) ? bboxObject : null,
                    out (double X1, double Y1, double X2, double Y2) bbox) ||
                bbox.X2 <= bbox.X1 ||
                bbox.Y2 <= bbox.Y1) {
                return false;
            }
            PdfDictionary? pageResources = ResolveDictionary(GetInheritedValue("Resources"));
            PdfDictionary? appearanceResources = ResolveDictionary(
                appearance.Dictionary.Items.TryGetValue("Resources", out PdfObject? resourcesObject)
                    ? resourcesObject
                    : null) ?? pageResources;
            PdfFontResourceSet fontResources = _fontResourceCache.GetOrCreate(appearanceResources, _objects);
            string content = WrapFormContentWithBoundingBoxClip(
                PdfEncoding.Latin1GetString(budget._pageContentBudget.Decode(appearance)),
                appearance.Dictionary);
            if (content.Length == 0) return false;

            CollectTextAndForms(
                content,
                appearanceResources,
                fontResources.Decoders,
                fontResources.WidthProviders,
                fontResources.Fonts,
                spans,
                new HashSet<PdfStream>(),
                GetPageSize().Height,
                includeArtifactText: true,
                textOutputBudget: budget._textOutputBudget,
                textClippingBudget: budget._textClippingBudget,
                pageContentBudget: budget._pageContentBudget);

            double pageHeight = GetPageSize().Height;
            string presentedText = string.Concat(spans
                .Where(span => span.IsVisible &&
                    span.Color?.A > 3 &&
                    Math.Abs(span.Advance) > 0.01D &&
                    span.FontSize > maximumTinyFontSizePoints &&
                    span.CanProjectCompleteText(pageHeight))
                .Select(static span => span.Text));
            for (int i = 0; i < expectedValues.Count; i++) {
                string expected = expectedValues[i];
                if (!string.IsNullOrEmpty(expected) &&
                    presentedText.IndexOf(expected, StringComparison.Ordinal) < 0) {
                    return false;
                }
            }
            return true;
        } catch (System.IO.InvalidDataException) {
            return false;
        }
    }

    internal bool DoesWidgetNormalAppearancePresentButtonState(
        int widgetObjectNumber,
        WidgetAppearanceScanBudget budget) {
        if (!_objects.TryGetValue(widgetObjectNumber, out PdfIndirectObject? widgetObject) ||
            widgetObject.Value is not PdfDictionary widget ||
            !TryGetNormalAppearanceStream(widget, out _)) {
            return false;
        }

        try {
            (double width, double height) = GetVisualPageSize();
            if (!TryReadRectangle(widget.Items.TryGetValue("Rect", out PdfObject? rectangleObject)
                    ? rectangleObject : null, out (double X1, double Y1, double X2, double Y2) rectangle)) return false;
            PdfVisualBounds widgetBounds = TransformBoundsToVisual(
                Math.Min(rectangle.X1, rectangle.X2), Math.Min(rectangle.Y1, rectangle.Y2),
                Math.Max(rectangle.X1, rectangle.X2), Math.Max(rectangle.Y1, rectangle.Y2));
            double left = Math.Max(0D, widgetBounds.Left);
            double top = Math.Max(0D, widgetBounds.Top);
            double cropWidth = Math.Min(width, widgetBounds.Right) - left;
            double cropHeight = Math.Min(height, widgetBounds.Bottom) - top;
            if (cropWidth <= 0D || cropHeight <= 0D) return false;
            const long maximumPixels = 1_000_000L;
            OfficeRasterScaleLimit rasterLimit = OfficeRasterScaleLimiter.Resolve(cropWidth, cropHeight, 1D, maximumPixels);
            if (!budget.TryConsumeRasterPixels(rasterLimit.PixelCount)) return false;
            var drawing = new OfficeDrawing(width, height);
            var selected = new PdfArray();
            selected.Items.Add(widget);
            budget.AddSelectedAnnotationAppearances(this, drawing, height, selected);
            if (drawing.Elements.Count == 0) return false;
            var cropped = new OfficeDrawing(cropWidth, cropHeight);
            cropped.AddClippedDrawingForRendering(drawing, 0D, 0D,
                OfficeClipPath.Rectangle(cropWidth, cropHeight), -left, -top);
            byte[] pixels = OfficeDrawingRasterRenderer.Render(cropped, new OfficeDrawingRasterRenderOptions {
                Scale = rasterLimit.Scale,
                MaximumRasterPixels = maximumPixels,
                ThrowOnImageDecodeFailure = true,
                CancellationToken = budget._pageContentBudget.CancellationToken
            }).GetPixels();
            for (int index = 3; index < pixels.Length; index += 4)
                if (pixels[index] > 0) return true;
            return false;
        } catch (System.IO.InvalidDataException) {
            return false;
        } catch (FormatException) {
            return false;
        } catch (InvalidCastException) {
            return false;
        } catch (OverflowException) {
            return false;
        } catch (ArgumentOutOfRangeException) {
            return false;
        } catch (NotSupportedException) {
            return false;
        } catch (OfficeImageExportLimitException) {
            return false;
        }
    }

    internal sealed class WidgetAppearanceScanBudget {
        internal readonly PageContentBudget _pageContentBudget;
        internal readonly TextContentParser.TextOutputBudget _textOutputBudget;
        internal readonly PdfTextClippingBudget _textClippingBudget;
        private readonly PdfTextClippingBudget _patternClippingBudget;
        private readonly Type3GlyphBudget _type3GlyphBudget;
        private long _remainingRasterPixels = 10_000_000L;

        internal bool TryConsumeRasterPixels(long pixels) {
            if (pixels <= 0L || pixels > _remainingRasterPixels) return false;
            _remainingRasterPixels -= pixels;
            return true;
        }

        internal void AddSelectedAnnotationAppearances(PdfReadPage page, OfficeDrawing drawing,
            double height, PdfArray selected) {
            page.AddAnnotationAppearances(drawing, height, page.GetVisualPageTransform(), _textOutputBudget,
                _pageContentBudget, _type3GlyphBudget, _textClippingBudget, _patternClippingBudget,
                _pageContentBudget.CancellationToken, selected);
        }

        internal WidgetAppearanceScanBudget(PdfReadPage page) {
            _pageContentBudget = new PageContentBudget(page);
            _textOutputBudget = new TextContentParser.TextOutputBudget(
                page._limits.MaxActualTextCharacters,
                page._limits.MaxDecodedTextCharacters);
            _textClippingBudget = new PdfTextClippingBudget();
            _patternClippingBudget = new PdfTextClippingBudget();
            _type3GlyphBudget = new Type3GlyphBudget(page._limits.MaxType3GlyphInvocationsPerPage);
        }
    }
}
