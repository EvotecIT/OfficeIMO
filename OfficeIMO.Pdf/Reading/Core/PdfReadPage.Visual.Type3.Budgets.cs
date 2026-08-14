namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private sealed class Type3GlyphBudget {
        private readonly int _maximum;
        private int _count;
        private int _failureVersion;
        private Type3SoftMaskValidationContext? _softMaskValidationContext;

        internal Type3GlyphBudget(int maximum) {
            _maximum = maximum;
        }

        internal void Consume(int count) {
            long next = (long)_count + count;
            if (next > _maximum) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.Type3GlyphInvocations, _maximum, next);
            }
            _count = (int)next;
        }

        internal int FailureVersion => _failureVersion;

        internal VisualGeometryBudget VisibilityGeometryBudget { get; } = new VisualGeometryBudget();

        internal Type3SoftMaskValidationContext GetOrCreateSoftMaskValidationContext(
            PdfReadPage owner,
            PageContentBudget pageContentBudget,
            PdfTextClippingBudget? invocationTextClippingBudget = null,
            PdfTextClippingBudget? patternTextClippingBudget = null) =>
            _softMaskValidationContext ??= new Type3SoftMaskValidationContext(
                this,
                owner.CreateTextOutputBudget(),
                pageContentBudget,
                invocationTextClippingBudget ?? new PdfTextClippingBudget(),
                patternTextClippingBudget ?? new PdfTextClippingBudget());

        internal void AttachSoftMaskValidationContext(Type3SoftMaskValidationContext context) =>
            _softMaskValidationContext = context;

        internal void RecordFailure() {
            _failureVersion++;
        }
    }

    private sealed class Type3SoftMaskValidationContext {
        internal Type3SoftMaskValidationContext(
            Type3GlyphBudget type3GlyphBudget,
            TextContentParser.TextOutputBudget textOutputBudget,
            PageContentBudget pageContentBudget,
            PdfTextClippingBudget invocationTextClippingBudget,
            PdfTextClippingBudget patternTextClippingBudget) {
            PageContentBudget = pageContentBudget;
            Type3GlyphBudget = type3GlyphBudget;
            TextOutputBudget = textOutputBudget;
            TransparencyProofPageContentBudget = pageContentBudget;
            TransparencyProofType3GlyphBudget = type3GlyphBudget;
            InvocationTextClippingBudget = invocationTextClippingBudget;
            PatternTextClippingBudget = patternTextClippingBudget;
        }

        internal PageContentBudget PageContentBudget { get; }

        internal Dictionary<(PdfStream Group, PdfDictionary? ParentResources, Matrix2D Transform, double Width, double Height), int> ValidatedGroups { get; } =
            new Dictionary<(PdfStream Group, PdfDictionary? ParentResources, Matrix2D Transform, double Width, double Height), int>();

        internal Type3GlyphBudget Type3GlyphBudget { get; }

        internal TextContentParser.TextOutputBudget TextOutputBudget { get; }

        internal PageContentBudget TransparencyProofPageContentBudget { get; }

        internal Type3GlyphBudget TransparencyProofType3GlyphBudget { get; }

        internal PdfTextClippingBudget InvocationTextClippingBudget { get; }

        internal PdfTextClippingBudget PatternTextClippingBudget { get; }

    }
}
