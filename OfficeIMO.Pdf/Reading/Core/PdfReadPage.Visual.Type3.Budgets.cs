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

        internal Type3SoftMaskValidationContext GetOrCreateSoftMaskValidationContext(
            PdfReadPage owner,
            TextContentParser.TextOutputBudget? textOutputBudget) =>
            _softMaskValidationContext ??= new Type3SoftMaskValidationContext(
                owner,
                _maximum,
                textOutputBudget ?? owner.CreateTextOutputBudget());

        internal void AttachSoftMaskValidationContext(Type3SoftMaskValidationContext context) =>
            _softMaskValidationContext = context;

        internal void RecordFailure() {
            _failureVersion++;
        }
    }

    private sealed class Type3SoftMaskValidationContext {
        internal Type3SoftMaskValidationContext(
            PdfReadPage owner,
            int maximumType3GlyphInvocations,
            TextContentParser.TextOutputBudget textOutputBudget) {
            PageContentBudget = new PageContentBudget(owner);
            Type3GlyphBudget = new Type3GlyphBudget(maximumType3GlyphInvocations);
            Type3GlyphBudget.AttachSoftMaskValidationContext(this);
            TextOutputBudget = textOutputBudget;
        }

        internal PageContentBudget PageContentBudget { get; }

        internal Dictionary<(PdfStream Group, Matrix2D Transform, double Width, double Height), int> ValidatedGroups { get; } =
            new Dictionary<(PdfStream Group, Matrix2D Transform, double Width, double Height), int>();

        internal Type3GlyphBudget Type3GlyphBudget { get; }

        internal TextContentParser.TextOutputBudget TextOutputBudget { get; }
    }
}
