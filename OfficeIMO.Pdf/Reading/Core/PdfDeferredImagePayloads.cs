using System.Runtime.ExceptionServices;

namespace OfficeIMO.Pdf;

/// <summary>
/// Builds the payloads of one page's extracted images when they are first needed instead of while reading.
/// </summary>
/// <remarks>
/// The payloads of a page draw on one colour-function budget and function cache, so whether a payload can be
/// normalized depends on the payloads built before it. Payloads are therefore built in the order an eager read
/// builds them, from a fresh page budget whenever an earlier one has to be built again, and each one is the
/// payload an eager read produces. Built bytes are held weakly: a reader that keeps them keeps them alive, and
/// bytes nobody holds can be collected and are built again on the next request. A payload that fails, for example
/// on a read limit, keeps failing with the same exception while the payloads after it stay available.
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

    private sealed class Slot {
        internal Slot(Builder builder) => Builder = builder;
        internal Builder Builder { get; }
        internal PdfExtractedImagePayloadFacts? Facts;
        internal WeakReference<byte[]>? Bytes;
        internal ExceptionDispatchInfo? Failure;
    }

    private readonly object _gate = new();
    private readonly Func<Budget> _createBudget;
    private readonly List<Slot> _slots = new();
    private Budget? _budget;
    private int _next;

    internal PdfDeferredImagePayloads(Func<Budget> createBudget) {
        _createBudget = createBudget ?? throw new ArgumentNullException(nameof(createBudget));
    }

    internal int Add(Builder builder) {
        lock (_gate) {
            _slots.Add(new Slot(builder));
            return _slots.Count - 1;
        }
    }

    internal PdfExtractedImagePayloadFacts GetFacts(int index, CancellationToken cancellationToken) {
        lock (_gate) {
            Slot slot = _slots[index];
            slot.Failure?.Throw();
            if (slot.Facts is { } facts) return facts;
        }
        return Get(index, cancellationToken).Facts;
    }

    internal PdfExtractedImagePayload Get(int index, CancellationToken cancellationToken) {
        lock (_gate) {
            Slot requested = _slots[index];
            requested.Failure?.Throw();
            if (requested.Facts is { } knownFacts && requested.Bytes is { } weak && weak.TryGetTarget(out byte[]? cached)) {
                return new PdfExtractedImagePayload(cached, knownFacts);
            }
            if (_budget is null || _next > index) {
                _budget = _createBudget();
                _next = 0;
            }

            while (true) {
                Slot slot = _slots[_next];
                Budget budget = _budget.Value;
                PdfExtractedImagePayload payload;
                try {
                    payload = slot.Builder(budget.ColorFunctionEvaluationBudget, budget.FunctionResolutionContext, cancellationToken);
                } catch (Exception exception) when (exception is not (OperationCanceledException or OutOfMemoryException)) {
                    // The build fails the same way every time, so later payloads continue from the budget it left.
                    slot.Failure = ExceptionDispatchInfo.Capture(exception);
                    if (_next++ == index) throw;
                    continue;
                } catch {
                    // An interrupted build may have spent part of the budget, so the next request starts over.
                    _budget = null;
                    throw;
                }
                slot.Facts = payload.Facts;
                slot.Bytes = new WeakReference<byte[]>(payload.Bytes);
                if (_next++ == index) return payload;
            }
        }
    }
}
