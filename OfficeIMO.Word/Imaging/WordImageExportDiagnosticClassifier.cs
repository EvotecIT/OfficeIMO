using System;
using OfficeIMO.Drawing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Classifies Word image-export diagnostics by their stable diagnostic-code family.
    /// </summary>
    internal static class WordImageExportDiagnosticClassifier {
        internal static OfficeImageExportDiagnostic Create(
            OfficeImageExportDiagnosticSeverity severity,
            string code,
            string message,
            string? source) =>
            new OfficeImageExportDiagnostic(
                severity,
                code,
                message,
                source,
                Classify(code));

        internal static OfficeImageExportLossKind Classify(string code) {
            if (string.IsNullOrWhiteSpace(code)) {
                throw new ArgumentException("Word image-export diagnostics require a stable code.", nameof(code));
            }

            switch (code) {
                case WordImageExportDiagnosticCodes.LimitedFloatingImageWrap:
                case WordImageExportDiagnosticCodes.LimitedFloatingShapeWrap:
                case WordImageExportDiagnosticCodes.LimitedFloatingTextBoxWrap:
                case WordImageExportDiagnosticCodes.LimitedSmartArt:
                    return OfficeImageExportLossKind.Approximation;

                case WordImageExportDiagnosticCodes.UnsupportedBodyElement:
                case WordImageExportDiagnosticCodes.UnsupportedExternalImage:
                case WordImageExportDiagnosticCodes.UnsupportedFloatingImage:
                case WordImageExportDiagnosticCodes.UnsupportedFooterElement:
                case WordImageExportDiagnosticCodes.UnsupportedFooterOverflow:
                case WordImageExportDiagnosticCodes.UnsupportedHeaderFooterMeasurementOverflow:
                case WordImageExportDiagnosticCodes.UnsupportedHeaderElement:
                case WordImageExportDiagnosticCodes.UnsupportedHeaderOverflow:
                case WordImageExportDiagnosticCodes.UnsupportedImage:
                case WordImageExportDiagnosticCodes.UnsupportedKeepMeasurementOverflow:
                case WordImageExportDiagnosticCodes.UnsupportedNestedTableOverflow:
                case WordImageExportDiagnosticCodes.UnsupportedNestedTable:
                case WordImageExportDiagnosticCodes.UnsupportedPageIndex:
                case WordImageExportDiagnosticCodes.UnsupportedPagination:
                case WordImageExportDiagnosticCodes.UnsupportedShape:
                case WordImageExportDiagnosticCodes.UnsupportedSmartArt:
                case WordImageExportDiagnosticCodes.UnsupportedTableCellMeasurementOverflow:
                case WordImageExportDiagnosticCodes.UnsupportedTableCellTextOverflow:
                case WordImageExportDiagnosticCodes.UnsupportedTableHeaderPagination:
                case WordImageExportDiagnosticCodes.UnsupportedTableImageOverflow:
                case WordImageExportDiagnosticCodes.UnsupportedTableRowPagination:
                case WordImageExportDiagnosticCodes.UnsupportedTextBox:
                    return OfficeImageExportLossKind.Omission;

                default:
                    throw new ArgumentOutOfRangeException(
                        nameof(code),
                        code,
                        "Word image-export diagnostic codes require an explicit fidelity-loss classification.");
            }
        }
    }
}
