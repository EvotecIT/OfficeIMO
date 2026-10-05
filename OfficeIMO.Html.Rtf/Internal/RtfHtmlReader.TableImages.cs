namespace OfficeIMO.Html;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private void FitIntrinsicCellImages(RtfTableCell cell, int availableWidth) {
            if (!_intrinsicCellImages.TryGetValue(cell, out var images)) return;
            foreach (var item in images) {
                RtfImage image = item.Image;
                if (!OfficeImageReader.TryIdentifyByContent(image.Data, null, out OfficeImageInfo info) ||
                    info.Width <= 0 || info.Height <= 0) continue;
                double width = image.DesiredWidthTwips ?? info.Width * 1440D / info.DpiX;
                double height = image.DesiredHeightTwips ?? info.Height * 1440D / info.DpiY;
                if (width <= availableWidth) continue;
                double scale = availableWidth / width;
                image.SourceWidth = info.Width;
                image.SourceHeight = info.Height;
                image.DesiredWidthTwips = availableWidth;
                image.DesiredHeightTwips = Math.Max(1, (int)Math.Round(height * scale));
                _options.AddDiagnostic("HtmlRtfImageFittedToCell",
                    "An unstyled HTML image was proportionally fitted to its RTF table cell.",
                    HtmlRenderStyleResolver.DescribeSource(item.Source), action: RtfConversionAction.Substituted);
            }
        }
    }
}
