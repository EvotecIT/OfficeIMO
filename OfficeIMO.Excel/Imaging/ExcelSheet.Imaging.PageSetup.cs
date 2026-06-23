using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const double ImageExportDpi = 96D;
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

            PageSetupCanvasGeometry geometry = ResolvePageSetupCanvasGeometry(pageSetup, options.Scale, content.Width, content.Height);
            OfficeImageLayer? contentLayer = CreatePageSetupContentLayer(format, content, geometry);
            if (contentLayer == null) {
                return content;
            }

            var diagnostics = new List<OfficeImageExportDiagnostic>(content.Diagnostics.Count + 1);
            diagnostics.AddRange(content.Diagnostics);
            AddPageSetupPaperSizeDiagnostic(pageSetup, diagnostics);

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
            pageSetup.PaperSizeCode.HasValue ||
            HasFitToPageScale(pageSetup) ||
            (pageSetup.Scale.HasValue && !HasFitToPageScale(pageSetup));

        private static bool HasFitToPageScale(ExcelSheetPageSetup pageSetup) =>
            pageSetup.FitToWidth.HasValue || pageSetup.FitToHeight.HasValue;

        private static bool HasUnsupportedFitToPageScale(ExcelSheetPageSetup pageSetup) =>
            IsUnsupportedFitDimension(pageSetup.FitToWidth) ||
            IsUnsupportedFitDimension(pageSetup.FitToHeight);

        private static bool IsUnsupportedFitDimension(uint? value) =>
            value.HasValue && value.Value > 1U;

        private static PageSetupCanvasGeometry ResolvePageSetupCanvasGeometry(
            ExcelSheetPageSetup pageSetup,
            double outputScale,
            int contentWidth,
            int contentHeight) {
            OfficePageSize pageSize = ResolvePageSize(pageSetup);
            pageSize = pageSetup.Orientation == ExcelPageOrientation.Landscape
                ? pageSize.Landscape()
                : pageSize.Portrait();
            ExcelSheetPageMargins? margins = pageSetup.Margins;

            int width = pageSize.ToPixelWidth(ImageExportDpi, outputScale);
            int height = pageSize.ToPixelHeight(ImageExportDpi, outputScale);
            double x = ClampMargin((margins?.Left ?? DefaultMarginLeftInches) * ImageExportDpi * outputScale, width);
            double y = ClampMargin((margins?.Top ?? DefaultMarginTopInches) * ImageExportDpi * outputScale, height);
            double right = ClampMargin((margins?.Right ?? DefaultMarginRightInches) * ImageExportDpi * outputScale, width);
            double bottom = ClampMargin((margins?.Bottom ?? DefaultMarginBottomInches) * ImageExportDpi * outputScale, height);
            double printableWidth = Math.Max(1D, width - x - right);
            double printableHeight = Math.Max(1D, height - y - bottom);
            double contentScale = ResolvePageSetupContentScale(
                pageSetup,
                Math.Max(1, contentWidth),
                Math.Max(1, contentHeight),
                printableWidth,
                printableHeight);
            return new PageSetupCanvasGeometry(width, height, x, y, contentScale);
        }

        private static double ResolvePageSetupContentScale(
            ExcelSheetPageSetup pageSetup,
            int contentWidth,
            int contentHeight,
            double printableWidth,
            double printableHeight) {
            if (!HasFitToPageScale(pageSetup)) {
                return Math.Max(0.1D, Math.Min(4D, (pageSetup.Scale ?? 100U) / 100D));
            }

            double scale = 1D;
            if (IsSupportedFitDimension(pageSetup.FitToWidth)) {
                scale = Math.Min(scale, printableWidth / contentWidth);
            }

            if (IsSupportedFitDimension(pageSetup.FitToHeight)) {
                scale = Math.Min(scale, printableHeight / contentHeight);
            }

            return Math.Max(0.1D, Math.Min(4D, scale));
        }

        private static bool IsSupportedFitDimension(uint? value) =>
            value.HasValue && value.Value == 1U;

        private static OfficePageSize ResolvePageSize(ExcelSheetPageSetup pageSetup) =>
            TryResolvePageSize(pageSetup.PaperSize, out OfficePageSize pageSize)
                ? pageSize
                : OfficePageSizes.Letter;

        private static bool TryResolvePageSize(ExcelPaperSize? paperSize, out OfficePageSize pageSize) {
            switch (paperSize) {
                case ExcelPaperSize.Letter:
                case ExcelPaperSize.LetterSmall:
                    pageSize = OfficePageSizes.Letter;
                    return true;
                case ExcelPaperSize.Tabloid:
                    pageSize = OfficePageSizes.Tabloid;
                    return true;
                case ExcelPaperSize.Ledger:
                    pageSize = OfficePageSizes.Ledger;
                    return true;
                case ExcelPaperSize.Legal:
                    pageSize = OfficePageSizes.Legal;
                    return true;
                case ExcelPaperSize.Statement:
                    pageSize = OfficePageSizes.Statement;
                    return true;
                case ExcelPaperSize.Executive:
                    pageSize = OfficePageSizes.Executive;
                    return true;
                case ExcelPaperSize.A3:
                    pageSize = OfficePageSizes.A3;
                    return true;
                case ExcelPaperSize.A4:
                case ExcelPaperSize.A4Small:
                    pageSize = OfficePageSizes.A4;
                    return true;
                case ExcelPaperSize.A5:
                    pageSize = OfficePageSizes.A5;
                    return true;
                case ExcelPaperSize.B4Jis:
                    pageSize = OfficePageSizes.B4Jis;
                    return true;
                case ExcelPaperSize.B5Jis:
                    pageSize = OfficePageSizes.B5Jis;
                    return true;
                default:
                    pageSize = default;
                    return false;
            }
        }

        private void AddPageSetupPaperSizeDiagnostic(
            ExcelSheetPageSetup pageSetup,
            List<OfficeImageExportDiagnostic> diagnostics) {
            if (!pageSetup.PaperSizeCode.HasValue) {
                diagnostics.Add(new OfficeImageExportDiagnostic(
                    OfficeImageExportDiagnosticSeverity.Info,
                    ExcelImageExportDiagnosticCodes.PageSetupPaperSizeDefaulted,
                    "Worksheet image page output used default Letter paper size because no worksheet paper size is configured.",
                    Name + "!pageSetup"));
                return;
            }

            if (!TryResolvePageSize(pageSetup.PaperSize, out _)) {
                diagnostics.Add(new OfficeImageExportDiagnostic(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.PageSetupPaperSizeUnsupported,
                    "Worksheet image page output used default Letter paper size because paper size code " + pageSetup.PaperSizeCode.Value + " is not supported yet.",
                    Name + "!pageSetup"));
            }
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
