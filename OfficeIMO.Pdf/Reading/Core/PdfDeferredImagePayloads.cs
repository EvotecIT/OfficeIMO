namespace OfficeIMO.Pdf;

/// <summary>
/// Builds the payloads of one page's extracted images when they are first needed instead of while reading.
/// </summary>
/// <remarks>
/// The payloads of a page draw on one colour-function budget and function cache, so whether a payload can be
/// normalized depends on the payloads built before it. Payloads are therefore built in the order an eager read
/// builds them, from a fresh page budget whenever an earlier one has to be built again, and each one is the
/// payload an eager read produces. Built payloads are held weakly: a reader that keeps the bytes keeps them alive,
/// and payloads nobody holds can be collected and are built again on the next request.
/// </remarks>
internal sealed class PdfDeferredImagePayloads {
    internal delegate PdfExtractedImagePayload Builder(
        Func<int, long, bool>? colorFunctionEvaluationBudget,
        PdfColorFunctionResolutionContext? functionResolutionContext,
        CancellationToken cancellationToken);

    internal readonly struct Budget {
        internal Budget(Func<int, long, bool>? colorFunctionEvaluationBudget, PdfColorFunctionResolutionContext? functionResolutionContext) {
            ColorFunctionEvaluationBudget = colorFunctionEvaluationBudget;
            FunctionResolutionContext = functionResolutionContext;
        }

        internal Func<int, long, bool>? ColorFunctionEvaluationBudget { get; }
        internal PdfColorFunctionResolutionContext? FunctionResolutionContext { get; }
    }

    private readonly object _gate = new();
    private readonly Func<Budget> _createBudget;
    private readonly List<Builder> _builders = new();
    private readonly List<WeakReference<PdfExtractedImagePayload>?> _built = new();
    private Budget? _budget;
    private int _next;

    internal PdfDeferredImagePayloads(Func<Budget> createBudget) {
        _createBudget = createBudget ?? throw new ArgumentNullException(nameof(createBudget));
    }

    internal int Add(Builder builder) {
        lock (_gate) {
            _builders.Add(builder);
            _built.Add(null);
            return _builders.Count - 1;
        }
    }

    internal PdfExtractedImagePayload Get(int index, CancellationToken cancellationToken) {
        lock (_gate) {
            if (_built[index] is { } built && built.TryGetTarget(out PdfExtractedImagePayload? cached)) return cached;
            if (_budget is null || _next > index) {
                _budget = _createBudget();
                _next = 0;
            }

            try {
                while (true) {
                    Budget budget = _budget.Value;
                    PdfExtractedImagePayload payload = _builders[_next](
                        budget.ColorFunctionEvaluationBudget,
                        budget.FunctionResolutionContext,
                        cancellationToken);
                    _built[_next] = new WeakReference<PdfExtractedImagePayload>(payload);
                    if (_next++ == index) return payload;
                }
            } catch {
                // A build that did not finish may have spent part of the budget, so the next request starts over.
                _budget = null;
                throw;
            }
        }
    }
}
