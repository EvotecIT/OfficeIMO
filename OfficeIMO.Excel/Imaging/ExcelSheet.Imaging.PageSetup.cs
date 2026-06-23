using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const double ImageExportDpi = 96D;
        private const double DefaultPageWidthInches = 8.5D;
        private const double DefaultPageHeightInches = 11D;
        private const double DefaultMarginLeftInches = 0.7D;
        private const double DefaultMarginRightInches = 0.7D;
        private const double DefaultMarginTopInches = 0.75D;
        private const double DefaultMarginBottomInches = 0.75D;

        private OfficeImageExportResult ApplyPageSetupCanvas(
            OfficeImageExportFormat format,
            OfficeImageExportResult content,
            ExcelWorksheetImageExportOptions options) {
            ExcelSheetPageSetup pageSetup = GetPageSetup();
            if (!ShouldApplyPageSetupCanvas(pageSetup)) {
                return content;
            }

            PageSetupCanvasGeometry geometry = ResolvePageSetupCanvasGeometry(pageSetup, options.Scale);
            OfficeImageLayer? contentLayer = CreatePageSetupContentLayer(format, content, geometry);
            if (contentLayer == null) {
                return content;
            }

            var diagnostics = new List<OfficeImageExportDiagnostic>(content.Diagnostics.Count + 1);
            diagnostics.AddRange(content.Diagnostics);
            diagnostics.Add(new OfficeImageExportDiagnostic(
                OfficeImageExportDiagnosticSeverity.Info,
                ExcelImageExportDiagnosticCodes.PageSetupPaperSizeDefaulted,
                "Worksheet image page output used default Letter paper size because paper-size-specific image page geometry is not implemented yet.",
                Name + "!pageSetup"));

            byte[] bytes = format == OfficeImageExportFormat.Svg
                ? OfficeImageComposer.ComposeSvgBytes(
                    geometry.Width,
                    geometry.Height,
                    options.BackgroundColor,
                    new[] { contentLayer })
                : OfficeImageComposer.ComposePng(
                    geometry.Width,
                    geometry.Height,
                    options.BackgroundColor,
                    new[] { contentLayer });

            return new OfficeImageExportResult(
                format,
                geometry.Width,
                geometry.Height,
                bytes,
                content.Name,
                content.Source,
                diagnostics.AsReadOnly());
        }

        private static bool ShouldApplyPageSetupCanvas(ExcelSheetPageSetup pageSetup) =>
            pageSetup.Orientation.HasValue ||
            pageSetup.Margins != null ||
            (pageSetup.Scale.HasValue && !HasFitToPageScale(pageSetup));

        private static bool HasFitToPageScale(ExcelSheetPageSetup pageSetup) =>
            pageSetup.FitToWidth.HasValue || pageSetup.FitToHeight.HasValue;

        private static PageSetupCanvasGeometry ResolvePageSetupCanvasGeometry(ExcelSheetPageSetup pageSetup, double outputScale) {
            bool landscape = pageSetup.Orientation == ExcelPageOrientation.Landscape;
            double pageWidthInches = landscape ? DefaultPageHeightInches : DefaultPageWidthInches;
            double pageHeightInches = landscape ? DefaultPageWidthInches : DefaultPageHeightInches;
            ExcelSheetPageMargins? margins = pageSetup.Margins;
            double contentScale = HasFitToPageScale(pageSetup)
                ? 1D
                : Math.Max(0.1D, Math.Min(4D, (pageSetup.Scale ?? 100U) / 100D));

            int width = Math.Max(1, (int)Math.Ceiling(pageWidthInches * ImageExportDpi * outputScale));
            int height = Math.Max(1, (int)Math.Ceiling(pageHeightInches * ImageExportDpi * outputScale));
            double x = ClampMargin((margins?.Left ?? DefaultMarginLeftInches) * ImageExportDpi * outputScale, width);
            double y = ClampMargin((margins?.Top ?? DefaultMarginTopInches) * ImageExportDpi * outputScale, height);
            double right = ClampMargin((margins?.Right ?? DefaultMarginRightInches) * ImageExportDpi * outputScale, width);
            double bottom = ClampMargin((margins?.Bottom ?? DefaultMarginBottomInches) * ImageExportDpi * outputScale, height);
            return new PageSetupCanvasGeometry(width, height, x, y, contentScale);
        }

        private static double ClampMargin(double margin, int pageSize) {
            if (double.IsNaN(margin) || double.IsInfinity(margin) || margin <= 0D) {
                return 0D;
            }

            return Math.Min(margin, Math.Max(0D, pageSize - 1D));
        }

        private static OfficeImageLayer? CreatePageSetupContentLayer(
            OfficeImageExportFormat format,
            OfficeImageExportResult content,
            PageSetupCanvasGeometry geometry) {
            double width = Math.Max(1D, content.Width * geometry.ContentScale);
            double height = Math.Max(1D, content.Height * geometry.ContentScale);
            if (format == OfficeImageExportFormat.Svg) {
                return OfficeImageLayer.FromSvgInner(
                    OfficeSvgFormatting.ExtractSvgInner(Encoding.UTF8.GetString(content.Bytes)),
                    geometry.ContentX,
                    geometry.ContentY,
                    width,
                    height);
            }

            if (!OfficePngReader.TryDecode(content.Bytes, out OfficeRasterImage? image) || image == null) {
                return null;
            }

            return OfficeImageLayer.FromRaster(image, geometry.ContentX, geometry.ContentY, width, height);
        }

        private readonly struct PageSetupCanvasGeometry {
            internal PageSetupCanvasGeometry(
                int width,
                int height,
                double contentX,
                double contentY,
                double contentScale) {
                Width = width;
                Height = height;
                ContentX = contentX;
                ContentY = contentY;
                ContentScale = contentScale;
            }

            internal int Width { get; }
            internal int Height { get; }
            internal double ContentX { get; }
            internal double ContentY { get; }
            internal double ContentScale { get; }
        }
    }
}
