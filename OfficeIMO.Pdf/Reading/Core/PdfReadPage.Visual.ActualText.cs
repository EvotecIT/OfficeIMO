using OfficeIMO.Drawing;
using System.Threading;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private sealed class ActualTextReplayFrame {
        internal ActualTextReplayFrame(TextContentParser.DrawingActualTextScope scope, OfficeDrawing drawing) {
            Scope = scope; Drawing = drawing;
        }
        internal TextContentParser.DrawingActualTextScope Scope { get; }
        internal OfficeDrawing Drawing { get; }
    }

    /// <summary>Replays closed path-only replacement scopes in their original paint order.</summary>
    private static void ReplayDrawingElements(OfficeDrawing drawing, List<PdfPageDrawingElement> elements,
        List<TextContentParser.DrawingActualTextScope> scopes,
        Action<OfficeDrawing, PdfPageDrawingElement> paint, CancellationToken cancellationToken) {
        if (scopes.Count == 0 || elements.Any(static element => element.ContentOrderKey == null)) {
            foreach (PdfPageDrawingElement element in elements) {
                cancellationToken.ThrowIfCancellationRequested();
                paint(drawing, element);
            }
            return;
        }

        scopes.Sort(static (left, right) => {
            int order = left.Start.CompareTo(right.Start);
            return order != 0 ? order : right.End.CompareTo(left.End);
        });
        var stack = new Stack<ActualTextReplayFrame>();
        int nextScope = 0;
        foreach (PdfPageDrawingElement element in elements) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfContentOrderKey key = element.ContentOrderKey!;
            while (stack.Count > 0 && stack.Peek().Scope.End.CompareTo(key) <= 0) CloseScope();
            while (nextScope < scopes.Count && scopes[nextScope].Start.CompareTo(key) < 0) {
                TextContentParser.DrawingActualTextScope scope = scopes[nextScope++];
                if (scope.End.CompareTo(key) <= 0 || scope.Owner.HasVisibleDrawingText) continue;
                // Locally balanced scopes are nested or disjoint. Never relabel crossing/unowned paint.
                if (stack.Count > 0 && scope.End.CompareTo(stack.Peek().Scope.End) > 0) continue;
                OfficeDrawing parent = stack.Count > 0 ? stack.Peek().Drawing : drawing;
                var child = new OfficeDrawing(drawing.Width, drawing.Height);
                child.Fonts.AddRange(parent.Fonts);
                stack.Push(new ActualTextReplayFrame(scope, child));
            }
            paint(stack.Count > 0 ? stack.Peek().Drawing : drawing, element);
        }
        while (stack.Count > 0) CloseScope();

        void CloseScope() {
            cancellationToken.ThrowIfCancellationRequested();
            ActualTextReplayFrame frame = stack.Pop();
            OfficeDrawing parent = stack.Count > 0 ? stack.Peek().Drawing : drawing;
            if (frame.Drawing.Elements.Count == 0) return;
            if (!frame.Scope.Owner.HasVisibleDrawingText && IsPathOnlyActualTextDrawing(frame.Drawing, cancellationToken)) {
                parent.AddActualTextDrawing(frame.Scope.Text, frame.Drawing,
                    frame.Scope.Owner.DrawingAnchorX, frame.Scope.Owner.DrawingAnchorY);
            } else {
                parent.AddDrawingForClippedRendering(frame.Drawing, 0D, 0D, null);
            }
        }
    }

    private static bool IsPathOnlyActualTextDrawing(OfficeDrawing drawing, CancellationToken cancellationToken) {
        bool hasPaint = false;
        foreach (OfficeDrawingElement element in drawing.Elements) {
            cancellationToken.ThrowIfCancellationRequested();
            switch (element) {
                case OfficeDrawingShape:
                    hasPaint = true;
                    break;
                case OfficeDrawingGroup group:
                    if (!IsPathOnlyActualTextDrawing(group.InnerDrawing, cancellationToken)) return false;
                    hasPaint = true;
                    break;
                case OfficeDrawingEffectGroup effect:
                    if (!IsPathOnlyActualTextDrawing(effect.InnerDrawing, cancellationToken)) return false;
                    hasPaint = true;
                    break;
                default:
                    return false;
            }
        }
        return hasPaint;
    }
}
