using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfRedactionApplier {
    private static Dictionary<int, TextObjectContext> CollectTextObjectContexts(
        string content,
        TextScrubGraphicsState? graphicsState,
        PdfReadLimits limits) {
        var contexts = new Dictionary<int, TextObjectContext>();
        TextScrubGraphicsState state = graphicsState ?? new TextScrubGraphicsState();
        Stack<TextScrubStateSnapshot> stack = state.Stack;
        Matrix2D current = state.Current;
        PdfTextStateSnapshot textState = state.TextState;
        PdfContentStreamInterpreter.Interpret(
            content,
            limits.MaxContentOperations,
            operation => {
                switch (operation.Name) {
                    case "q":
                        stack.Push(new TextScrubStateSnapshot(current, textState));
                        break;
                    case "Q":
                        if (stack.Count > 0) {
                            TextScrubStateSnapshot restored = stack.Pop();
                            current = restored.Transform;
                            textState = restored.TextState;
                        } else {
                            current = Matrix2D.Identity;
                            textState = textState.WithTextRenderingMode(0);
                        }
                        break;
                    case "cm" when operation.Operands.Count >= 6:
                        int start = operation.Operands.Count - 6;
                        current = Matrix2D.Multiply(current, new Matrix2D(
                            Convert.ToDouble(operation.Operands[start], CultureInfo.InvariantCulture),
                            Convert.ToDouble(operation.Operands[start + 1], CultureInfo.InvariantCulture),
                            Convert.ToDouble(operation.Operands[start + 2], CultureInfo.InvariantCulture),
                            Convert.ToDouble(operation.Operands[start + 3], CultureInfo.InvariantCulture),
                            Convert.ToDouble(operation.Operands[start + 4], CultureInfo.InvariantCulture),
                            Convert.ToDouble(operation.Operands[start + 5], CultureInfo.InvariantCulture)));
                        break;
                    case "BT":
                        contexts[operation.OperatorOffset] = new TextObjectContext(current, textState);
                        break;
                    case "Tf" when operation.Operands.Count >= 2:
                        string font = operation.Operands[operation.Operands.Count - 2] as string ?? textState.FontResource;
                        textState = textState.WithFont(font, ReadTextStateNumber(operation.Operands, operation.Operands.Count - 1, textState.FontSize));
                        break;
                    case "Tc" when operation.Operands.Count >= 1:
                        textState = textState.WithCharacterSpacing(ReadTextStateNumber(operation.Operands, operation.Operands.Count - 1, textState.CharacterSpacing));
                        break;
                    case "Tw" when operation.Operands.Count >= 1:
                        textState = textState.WithWordSpacing(ReadTextStateNumber(operation.Operands, operation.Operands.Count - 1, textState.WordSpacing));
                        break;
                    case "Tz" when operation.Operands.Count >= 1:
                        textState = textState.WithHorizontalScaling(ReadTextStateNumber(operation.Operands, operation.Operands.Count - 1, textState.HorizontalScaling * 100D) / 100D);
                        break;
                    case "TL" when operation.Operands.Count >= 1:
                        textState = textState.WithLeading(ReadTextStateNumber(operation.Operands, operation.Operands.Count - 1, textState.Leading));
                        break;
                    case "TD" when operation.Operands.Count >= 2:
                        textState = textState.WithLeading(-ReadTextStateNumber(operation.Operands, operation.Operands.Count - 1, -textState.Leading));
                        break;
                    case "Ts" when operation.Operands.Count >= 1:
                        textState = textState.WithTextRise(ReadTextStateNumber(operation.Operands, operation.Operands.Count - 1, textState.TextRise));
                        break;
                    case "Tr" when operation.Operands.Count >= 1:
                        textState = textState.WithTextRenderingMode((int)ReadTextStateNumber(operation.Operands, operation.Operands.Count - 1, textState.TextRenderingMode));
                        break;
                    case "\"" when operation.Operands.Count >= 3:
                        textState = textState
                            .WithWordSpacing(ReadTextStateNumber(operation.Operands, operation.Operands.Count - 3, textState.WordSpacing))
                            .WithCharacterSpacing(ReadTextStateNumber(operation.Operands, operation.Operands.Count - 2, textState.CharacterSpacing));
                        break;
                }
            },
            maxNestingDepth: limits.MaxContentNestingDepth,
            maxOperands: limits.MaxContentOperands);
        state.Current = current;
        state.TextState = textState;
        return contexts;
    }

    private static double ReadTextStateNumber(IReadOnlyList<object> operands, int index, double fallback) =>
        index >= 0 && index < operands.Count && operands[index] is double value ? value : fallback;

    private sealed class TextScrubGraphicsState {
        internal Matrix2D Current { get; set; } = Matrix2D.Identity;
        internal PdfTextStateSnapshot TextState { get; set; } = PdfTextStateSnapshot.Default;
        internal Stack<TextScrubStateSnapshot> Stack { get; } = new Stack<TextScrubStateSnapshot>();

        internal void Reset() {
            Current = Matrix2D.Identity;
            TextState = PdfTextStateSnapshot.Default;
            Stack.Clear();
        }
    }

    private readonly struct TextObjectContext {
        internal TextObjectContext(Matrix2D transform, PdfTextStateSnapshot textState) {
            Transform = transform;
            TextState = textState;
        }

        internal Matrix2D Transform { get; }
        internal PdfTextStateSnapshot TextState { get; }
    }

    private readonly struct TextScrubStateSnapshot {
        internal TextScrubStateSnapshot(Matrix2D transform, PdfTextStateSnapshot textState) {
            Transform = transform;
            TextState = textState;
        }

        internal Matrix2D Transform { get; }
        internal PdfTextStateSnapshot TextState { get; }
    }
}
