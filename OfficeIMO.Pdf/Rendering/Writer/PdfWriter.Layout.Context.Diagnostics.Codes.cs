namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private readonly struct PaintDiagnosticCodes {
            private PaintDiagnosticCodes(string clipped, string ancestorClipped, string pageClipped, string overflow, string unmeasured) {
                Clipped = clipped; AncestorClipped = ancestorClipped; PageClipped = pageClipped;
                Overflow = overflow; Unmeasured = unmeasured;
            }
            internal string Clipped { get; }
            internal string AncestorClipped { get; }
            internal string PageClipped { get; }
            internal string Overflow { get; }
            internal string Unmeasured { get; }

            internal static readonly PaintDiagnosticCodes CanvasText = new(
                OfficeDiagnosticCodes.CanvasTextPaintClipped,
                OfficeDiagnosticCodes.CanvasTextCanvasClipPaintClipped,
                OfficeDiagnosticCodes.CanvasTextPageBoundsPaintClipped,
                OfficeDiagnosticCodes.CanvasTextPaintOverflow,
                OfficeDiagnosticCodes.CanvasTextPaintBoundsUnmeasured);

            internal static readonly PaintDiagnosticCodes CanvasTextBox = new(
                OfficeDiagnosticCodes.CanvasTextBoxPaintClipped,
                OfficeDiagnosticCodes.CanvasTextBoxCanvasClipPaintClipped,
                OfficeDiagnosticCodes.CanvasTextBoxPageBoundsPaintClipped,
                OfficeDiagnosticCodes.CanvasTextBoxPaintOverflow,
                OfficeDiagnosticCodes.CanvasTextBoxPaintBoundsUnmeasured);

            internal static readonly PaintDiagnosticCodes FreeTextAnnotation = new(
                OfficeDiagnosticCodes.FreeTextAnnotationPaintClipped,
                OfficeDiagnosticCodes.FreeTextAnnotationCanvasClipPaintClipped,
                OfficeDiagnosticCodes.FreeTextAnnotationPageBoundsPaintClipped,
                OfficeDiagnosticCodes.FreeTextAnnotationPaintOverflow,
                OfficeDiagnosticCodes.FreeTextAnnotationPaintBoundsUnmeasured);

            internal static readonly PaintDiagnosticCodes Paragraph = new(
                OfficeDiagnosticCodes.ParagraphPaintClipped,
                OfficeDiagnosticCodes.ParagraphCanvasClipPaintClipped,
                OfficeDiagnosticCodes.ParagraphPageBoundsPaintClipped,
                OfficeDiagnosticCodes.ParagraphPaintOverflow,
                OfficeDiagnosticCodes.ParagraphPaintBoundsUnmeasured);

            internal static readonly PaintDiagnosticCodes TableCell = new(
                OfficeDiagnosticCodes.TableCellPaintClipped,
                OfficeDiagnosticCodes.TableCellCanvasClipPaintClipped,
                OfficeDiagnosticCodes.TableCellPageBoundsPaintClipped,
                OfficeDiagnosticCodes.TableCellPaintOverflow,
                OfficeDiagnosticCodes.TableCellPaintBoundsUnmeasured);
        }
    }
}
