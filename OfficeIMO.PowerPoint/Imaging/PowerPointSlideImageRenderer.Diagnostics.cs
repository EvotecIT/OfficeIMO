using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint;

internal static partial class PowerPointSlideImageRenderer {
    private static void AddDiagnostic(
        List<OfficeImageExportDiagnostic> diagnostics,
        string code,
        string message,
        OfficeConversionLossKind lossKind) {
        diagnostics.Add(new OfficeImageExportDiagnostic(
            OfficeImageExportDiagnosticSeverity.Warning,
            code,
            message,
            "PowerPoint slide",
            lossKind));
    }
}
