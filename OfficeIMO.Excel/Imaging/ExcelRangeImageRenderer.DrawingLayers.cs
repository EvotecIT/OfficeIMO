using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    internal static partial class ExcelRangeImageRenderer {
        private static void RenderRasterDrawingLayers(
            OfficeRasterCanvas canvas,
            ExcelRangeVisualSnapshot snapshot,
            ExcelImageExportOptions options,
            List<OfficeImageExportDiagnostic>? diagnostics,
            System.Threading.CancellationToken cancellationToken) {
            foreach (ExcelVisualDrawingLayer layer in snapshot.DrawingLayers) {
                cancellationToken.ThrowIfCancellationRequested();
                switch (layer.Kind) {
                    case ExcelVisualDrawingLayerKind.DrawingObject:
                        if (layer.DrawingObject != null) {
                            RenderRasterDrawingObject(canvas, layer.DrawingObject, options, diagnostics, cancellationToken);
                        }

                        break;
                    case ExcelVisualDrawingLayerKind.Image:
                        if (layer.Image != null) {
                            RenderRasterImage(canvas, layer.Image, options, diagnostics, cancellationToken);
                        }

                        break;
                    case ExcelVisualDrawingLayerKind.Chart:
                        if (layer.Chart != null) {
                            RenderRasterChart(canvas, snapshot, layer.Chart, options, diagnostics, cancellationToken);
                        }

                        break;
                    case ExcelVisualDrawingLayerKind.CommentBody:
                        if (layer.CommentBody != null) {
                            RenderRasterCommentBody(canvas, layer.CommentBody, options);
                        }

                        break;
                }
            }
        }

        private static void AppendSvgDrawingLayers(
            StringBuilder builder,
            ExcelRangeVisualSnapshot snapshot,
            ExcelImageExportOptions options,
            List<OfficeImageExportDiagnostic>? diagnostics,
            OfficeRasterCanvas textMeasurer,
            System.Threading.CancellationToken cancellationToken,
            OfficeSvgUtf8CompositionBudget? budget) {
            int imageIndex = 0;
            foreach (ExcelVisualDrawingLayer layer in snapshot.DrawingLayers) {
                cancellationToken.ThrowIfCancellationRequested();
                int currentImageIndex = imageIndex;
                AppendSvgFragment(builder, budget, fragment => {
                    switch (layer.Kind) {
                        case ExcelVisualDrawingLayerKind.DrawingObject:
                            if (layer.DrawingObject != null) {
                                AppendSvgDrawingObject(fragment, layer.DrawingObject, options, diagnostics, cancellationToken);
                            }

                            break;
                        case ExcelVisualDrawingLayerKind.Image:
                            if (layer.Image != null) {
                                AppendSvgImage(fragment, snapshot, layer.Image, options, diagnostics, ref currentImageIndex);
                            }

                            break;
                        case ExcelVisualDrawingLayerKind.Chart:
                            if (layer.Chart != null) {
                                AppendSvgChart(fragment, snapshot, layer.Chart, options, diagnostics, cancellationToken);
                            }

                            break;
                        case ExcelVisualDrawingLayerKind.CommentBody:
                            if (layer.CommentBody != null) {
                                AppendSvgCommentBody(fragment, layer.CommentBody, options, textMeasurer);
                            }

                            break;
                    }
                });
                imageIndex = currentImageIndex;
            }
        }
    }
}
