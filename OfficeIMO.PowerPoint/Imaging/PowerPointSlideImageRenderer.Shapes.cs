using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint;

internal static partial class PowerPointSlideImageRenderer {
    private static void AddAutoShape(OfficeDrawing drawing, PowerPointShape shape, List<OfficeImageExportDiagnostic> diagnostics, PowerPointShapeBoundsMapping mapping, A.ColorScheme? colorScheme) {
        string? presetName = GetAutoShapePresetName(shape);
        if (IsZeroThicknessLinePreset(presetName) && TryAddZeroThicknessLine(drawing, shape, diagnostics, mapping, presetName, colorScheme)) {
            return;
        }

        if (IsBentConnectorPreset(presetName) && TryAddBentConnector(drawing, shape, diagnostics, mapping, colorScheme)) {
            return;
        }

        if (IsCurvedConnectorPreset(presetName) && TryAddCurvedConnector(drawing, shape, diagnostics, mapping, colorScheme)) {
            return;
        }

        if (!TryGetBounds(shape, drawing, diagnostics, mapping, out double left, out double top, out double width, out double height)) {
            return;
        }

        if (TryAddCustomGeometryShape(drawing, shape, left, top, width, height, diagnostics, mapping, colorScheme)) {
            return;
        }

        if (!OfficeShapePresets.TryCreate(presetName, width, height, out OfficeShape? drawingShape) || drawingShape == null) {
            AddUnsupportedShapeDiagnostic(diagnostics, shape, "Skipped a PowerPoint auto shape geometry that is not yet projected through OfficeIMO.Drawing.");
            return;
        }

        ApplyShapeStyle(drawingShape, shape, colorScheme, mapping, diagnostics);
        ApplyShapeTransform(drawingShape, shape, width, height);
        drawing.AddShape(drawingShape, left, top);
    }

    private static string? GetAutoShapePresetName(PowerPointShape shape) {
        return GetOpenXmlShapeProperties(shape)?.GetFirstChild<A.PresetGeometry>()?.Preset?.InnerText;
    }
}
