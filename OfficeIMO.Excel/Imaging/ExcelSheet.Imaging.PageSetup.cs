using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const double ImageExportDpi = 96D;

        private OfficeImageExportResult ApplyPageSetupCanvas(
            OfficeImageExportFormat format,
            OfficeImageExportFormat rasterPlanningFormat,
            OfficeImageExportResult content,
            ExcelWorksheetImageExportOptions options,
            ref ExcelRasterRenderState rasterState) {
            ExcelSheetPageSetup pageSetup = GetPageSetup();
            if (!ShouldApplyPageSetupCanvas(pageSetup)) {
                return content;
            }

            double contentScaleRatio = 1D;
            PageSetupCanvasGeometry geometry = ResolvePageSetupCanvasGeometry(
                pageSetup,
                rasterState.Scale,
                content.Width,
                content.Height);
            OfficeRasterExportPlan? plan = null;
            if (format != OfficeImageExportFormat.Svg) {
                plan = OfficeRasterExportPlanner.Resolve(
                    geometry.Width / rasterState.Scale,
                    geometry.Height / rasterState.Scale,
                    rasterPlanningFormat,
                    options,
                    Name + "!pageSetup");
                if (plan.Value.Limit.Scale < rasterState.Scale) {
                    contentScaleRatio = plan.Value.Limit.Scale / rasterState.Scale;
                    rasterState = ExcelRasterRenderState.FromPlan(plan.Value);
                    geometry = ResolvePageSetupCanvasGeometry(
                        pageSetup,
                        rasterState.Scale,
                        Math.Max(1, (int)Math.Ceiling(content.Width * contentScaleRatio)),
                        Math.Max(1, (int)Math.Ceiling(content.Height * contentScaleRatio)));
                    geometry = new PageSetupCanvasGeometry(
                        plan.Value.Limit.PixelWidth,
                        plan.Value.Limit.PixelHeight,
                        geometry.ContentX,
                        geometry.ContentY,
                        geometry.ContentScale);
                }
            }

            OfficeImageLayer? contentLayer = CreatePageSetupContentLayer(
                format,
                content,
                geometry,
                contentScaleRatio);
            if (contentLayer == null) {
                return content;
            }

            var diagnostics = new List<OfficeImageExportDiagnostic>(content.Diagnostics.Count + 2);
            diagnostics.AddRange(content.Diagnostics);
            if (contentScaleRatio < 1D && plan?.Diagnostic != null) {
                diagnostics.Add(plan.Value.Diagnostic);
            }
            AddPageSetupPaperSizeDiagnostic(pageSetup, diagnostics);

            byte[] bytes;
            if (format == OfficeImageExportFormat.Svg) {
                bytes = OfficeImageComposer.ComposeSvgBytes(
                    geometry.Width,
                    geometry.Height,
                    options.BackgroundColor,
                    new[] { contentLayer });
            } else {
                OfficeRasterImage image = OfficeImageComposer.ComposeRaster(
                    geometry.Width,
                    geometry.Height,
                    options.BackgroundColor,
                    new[] { contentLayer });
                bytes = OfficeRasterImageEncoder.Encode(
                    image,
                    format,
                    rasterState.EncodingOptions);
            }

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
            ExcelPageSetupGeometry.HasFitToPageScale(pageSetup) ||
            (pageSetup.Scale.HasValue && !ExcelPageSetupGeometry.HasFitToPageScale(pageSetup));

        private static PageSetupCanvasGeometry ResolvePageSetupCanvasGeometry(
            ExcelSheetPageSetup pageSetup,
            double outputScale,
            int contentWidth,
            int contentHeight) {
            OfficePageSize pageSize = ExcelPageSetupGeometry.ResolvePageSize(pageSetup, OfficePageSizes.Letter);
            ExcelSheetPageMargins? margins = pageSetup.Margins;

            int width = pageSize.ToPixelWidth(ImageExportDpi, outputScale);
            int height = pageSize.ToPixelHeight(ImageExportDpi, outputScale);
            double x = ExcelPageSetupGeometry.ClampMargin((margins?.Left ?? ExcelPageSetupGeometry.DefaultMarginLeftInches) * ImageExportDpi * outputScale, width);
            double y = ExcelPageSetupGeometry.ClampMargin((margins?.Top ?? ExcelPageSetupGeometry.DefaultMarginTopInches) * ImageExportDpi * outputScale, height);
            double right = ExcelPageSetupGeometry.ClampMargin((margins?.Right ?? ExcelPageSetupGeometry.DefaultMarginRightInches) * ImageExportDpi * outputScale, width);
            double bottom = ExcelPageSetupGeometry.ClampMargin((margins?.Bottom ?? ExcelPageSetupGeometry.DefaultMarginBottomInches) * ImageExportDpi * outputScale, height);
            double printableWidth = Math.Max(1D, width - x - right);
            double printableHeight = Math.Max(1D, height - y - bottom);
            double contentScale = ExcelPageSetupGeometry.ResolveContentScale(
                pageSetup,
                Math.Max(1, contentWidth),
                Math.Max(1, contentHeight),
                printableWidth,
                printableHeight,
                0.1D,
                4D);
            return new PageSetupCanvasGeometry(width, height, x, y, contentScale);
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

            if (!ExcelPageSetupGeometry.TryResolvePageSize(pageSetup.PaperSize, out _)) {
                diagnostics.Add(new OfficeImageExportDiagnostic(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.PageSetupPaperSizeUnsupported,
                    "Worksheet image page output used default Letter paper size because paper size code " + pageSetup.PaperSizeCode.Value + " is not supported yet.",
                    Name + "!pageSetup"));
            }
        }

        private static OfficeImageLayer? CreatePageSetupContentLayer(
            OfficeImageExportFormat format,
            OfficeImageExportResult content,
            PageSetupCanvasGeometry geometry,
            double contentScaleRatio) {
            double width = Math.Max(1D, content.Width * contentScaleRatio * geometry.ContentScale);
            double height = Math.Max(1D, content.Height * contentScaleRatio * geometry.ContentScale);
            if (format == OfficeImageExportFormat.Svg) {
                string svgInner = OfficeSvgFormatting.ExtractSvgInner(Encoding.UTF8.GetString(content.Bytes));
                double svgScale = contentScaleRatio * geometry.ContentScale;
                if (Math.Abs(svgScale - 1D) > 0.0000001D) {
                    svgInner = "<g transform=\"scale(" + OfficeSvgFormatting.FormatNumber(svgScale) + ")\">" + svgInner + "</g>";
                }

                return OfficeImageLayer.FromSvgInner(
                    svgInner,
                    geometry.ContentX,
                    geometry.ContentY,
                    width,
                    height);
            }

            if (!OfficeRasterImageDecoder.TryDecode(content.Bytes, out OfficeRasterImage? image) || image == null) {
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
