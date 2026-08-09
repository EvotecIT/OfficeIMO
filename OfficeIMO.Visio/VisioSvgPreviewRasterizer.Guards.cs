using System.Collections.Generic;
using System.Xml.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio {
    internal static partial class VisioSvgPreviewRasterizer {
        private static bool RenderElement(
            OfficeRasterCanvas canvas,
            XElement element,
            SvgPaint inherited,
            SvgTransform transform,
            SvgRenderContext context) {
            if (!context.TryEnterRenderElement()) return false;
            try {
                return RenderElementWithinBudget(canvas, element, inherited, transform, context);
            } finally {
                context.ExitRenderElement();
            }
        }

        private static void AddSvgLossDiagnostic(
            ICollection<OfficeImageExportDiagnostic>? diagnostics,
            string? source,
            string message,
            OfficeConversionLossKind lossKind) {
            diagnostics?.Add(new OfficeImageExportDiagnostic(
                OfficeImageExportDiagnosticSeverity.Warning,
                OfficeImageExportDiagnosticCodes.SourceSvgPreviewLoss,
                message,
                source,
                lossKind));
        }
    }
}
