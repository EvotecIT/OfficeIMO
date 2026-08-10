namespace OfficeIMO.Pdf;

internal static class PdfPageGraphicsEffectTimelineParser {
    public static IReadOnlyList<PdfPageDrawingEffectTransition> Parse(
        string content,
        IReadOnlyDictionary<string, PdfPageGraphicsStateResource>? graphicsStates,
        PdfPageDrawingEffect initialEffect,
        Matrix2D initialTransform,
        PdfContentOrderKey? contentOrderPrefix = null,
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
        var stack = new Stack<(PdfPageDrawingEffect Effect, Matrix2D Transform)>();
        PdfPageDrawingEffect state = initialEffect;
        Matrix2D transform = initialTransform;
        PdfContentStreamInterpreter.Interpret(
            content,
            maxOperations,
            operation => {
                double paintOrder = paintOrderBase +
                    ((operation.OperatorOffset + paintOrderOffset) * paintOrderScale);
                PdfContentOrderKey? contentOrderKey = contentOrderPrefix?.Append(operation.OperatorOffset);
                switch (operation.Name) {
                    case "q":
                        stack.Push((state, transform));
                        break;
                    case "Q":
                        (PdfPageDrawingEffect Effect, Matrix2D Transform) restored = stack.Count > 0
                            ? stack.Pop()
                            : (initialEffect, initialTransform);
                        transform = restored.Transform;
                        ApplyState(restored.Effect, paintOrder, contentOrderKey);
                        break;
                    case "cm":
                        if (operation.Operands.Count >= 6) {
                            int start = operation.Operands.Count - 6;
                            var matrix = new Matrix2D(
                                NumberAt(operation.Operands, start),
                                NumberAt(operation.Operands, start + 1),
                                NumberAt(operation.Operands, start + 2),
                                NumberAt(operation.Operands, start + 3),
                                NumberAt(operation.Operands, start + 4),
                                NumberAt(operation.Operands, start + 5));
                            transform = Matrix2D.Multiply(transform, matrix);
                        }
                        break;
                    case "gs":
                        string? resourceName = operation.Operands.Count == 0
                            ? null
                            : operation.Operands[operation.Operands.Count - 1] as string;
                        if (resourceName is not null &&
                            graphicsStates is not null &&
                            graphicsStates.TryGetValue(resourceName, out PdfPageGraphicsStateResource resource)) {
                            PdfPageDrawingEffect updated = state.Apply(resource);
                            if (resource.HasSoftMask && updated.SoftMask != null) {
                                updated = updated.WithSoftMaskTransform(transform);
                            }
                            ApplyState(updated, paintOrder, contentOrderKey);
                        }
                        break;
                }
            },
            maxNestingDepth: maxNestingDepth,
            maxOperands: maxOperands);
        return transitions.Count == 0
            ? Array.Empty<PdfPageDrawingEffectTransition>()
            : transitions.AsReadOnly();

        void ApplyState(PdfPageDrawingEffect updated, double paintOrder, PdfContentOrderKey? contentOrderKey) {
            if (SameEffect(state, updated)) {
                return;
            }

            state = updated;
            transitions.Add(new PdfPageDrawingEffectTransition(paintOrder, state, contentOrderKey));
        }
    }

    private static double NumberAt(IReadOnlyList<object> operands, int index) =>
        operands[index] is double value ? value : 0D;

    private static bool SameEffect(PdfPageDrawingEffect left, PdfPageDrawingEffect right) =>
        left.BlendMode == right.BlendMode &&
        ReferenceEquals(left.SoftMask, right.SoftMask) &&
        left.HasBlendMode == right.HasBlendMode &&
        left.HasSoftMask == right.HasSoftMask &&
        Nullable.Equals(left.SoftMaskTransform, right.SoftMaskTransform);
}
