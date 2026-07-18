namespace OfficeIMO.Pdf;

internal static class PdfPageGraphicsEffectTimelineParser {
    public static IReadOnlyList<PdfPageDrawingEffectTransition> Parse(
        string content,
        IReadOnlyDictionary<string, PdfPageGraphicsStateResource>? graphicsStates,
        PdfPageDrawingEffect initialEffect,
        double paintOrderBase = 0D,
        double paintOrderScale = 1D,
        double paintOrderOffset = 0D,
        int maxOperations = PdfReadLimits.DefaultMaxContentOperations,
        int maxNestingDepth = PdfReadLimits.DefaultMaxContentNestingDepth,
        int maxOperands = PdfReadLimits.DefaultMaxContentOperands) {
        if (string.IsNullOrEmpty(content)) {
            return Array.Empty<PdfPageDrawingEffectTransition>();
        }

        var transitions = new List<PdfPageDrawingEffectTransition>();
        var stack = new Stack<PdfPageDrawingEffect>();
        PdfPageDrawingEffect state = initialEffect;
        PdfContentStreamInterpreter.Interpret(
            content,
            maxOperations,
            operation => {
                double paintOrder = paintOrderBase +
                    ((operation.OperatorOffset + paintOrderOffset) * paintOrderScale);
                switch (operation.Name) {
                    case "q":
                        stack.Push(state);
                        break;
                    case "Q":
                        ApplyState(stack.Count > 0 ? stack.Pop() : initialEffect, paintOrder);
                        break;
                    case "gs":
                        string? resourceName = operation.Operands.Count == 0
                            ? null
                            : operation.Operands[operation.Operands.Count - 1] as string;
                        if (resourceName is not null &&
                            graphicsStates is not null &&
                            graphicsStates.TryGetValue(resourceName, out PdfPageGraphicsStateResource resource)) {
                            ApplyState(state.Apply(resource), paintOrder);
                        }
                        break;
                }
            },
            maxNestingDepth: maxNestingDepth,
            maxOperands: maxOperands);
        return transitions.Count == 0
            ? Array.Empty<PdfPageDrawingEffectTransition>()
            : transitions.AsReadOnly();

        void ApplyState(PdfPageDrawingEffect updated, double paintOrder) {
            if (SameEffect(state, updated)) {
                return;
            }

            state = updated;
            transitions.Add(new PdfPageDrawingEffectTransition(paintOrder, state));
        }
    }

    private static bool SameEffect(PdfPageDrawingEffect left, PdfPageDrawingEffect right) =>
        left.BlendMode == right.BlendMode &&
        ReferenceEquals(left.SoftMask, right.SoftMask);
}
