using System.Runtime.ExceptionServices;

namespace OfficeIMO.Pdf;

internal sealed partial class PdfReadDocument {
    // One gate serializes every deferred feature. Extractors share the decoded-stream budget and
    // destination caches, so two features must never run at the same time on different threads.
    private readonly object _featureGate = new();
    private readonly DocumentFeature<PdfTaggedContentInfo?> _taggedContentFeature;
    private readonly DocumentFeature<IReadOnlyList<PdfOutlineItem>> _outlinesFeature;
    private readonly DocumentFeature<NamedDestinationIndex> _namedDestinationsFeature;
    private readonly DocumentFeature<PdfXmpMetadataInfo?> _xmpMetadataFeature;
    private readonly DocumentFeature<IReadOnlyList<PdfAttachmentInfo>> _attachmentsFeature;
    private readonly DocumentFeature<FormState> _formsFeature;
    private readonly DocumentFeature<PdfPortfolioInfo?> _portfolioFeature;
    // Orphaned-object analysis walks every reachable reference, so it waits for a repair report reader.
    private readonly DocumentFeature<PdfRepairReport> _repairReportFeature;

    private IReadOnlyList<PdfNamedDestination> _namedDestinations => _namedDestinationsFeature.GetValue().Destinations;
    private IReadOnlyList<PdfAttachmentInfo> _attachments => _attachmentsFeature.GetValue();
    private PdfXmpMetadataInfo? _xmpMetadata => _xmpMetadataFeature.GetValue();
    private PdfTaggedContentInfo? _taggedContent => _taggedContentFeature.GetValue();
    private IReadOnlyList<PdfOutlineItem> _outlines => _outlinesFeature.GetValue();
    private PdfPortfolioInfo? _portfolio => _portfolioFeature.GetValue();
    private IReadOnlyList<PdfFormField> _formFields => _formsFeature.GetValue().Fields;
    private int _formWidgetJavaScriptCount => _formsFeature.GetValue().WidgetJavaScriptCount;
    private long _formWidgetJavaScriptBytes => _formsFeature.GetValue().WidgetJavaScriptBytes;
    private PdfAcroFormXfaInfo? _acroFormXfa => _formsFeature.GetValue().Xfa;

    private DocumentFeature<T> Defer<T>(Func<CancellationToken, T> extract) => new(_featureGate, extract);

    private FormState ExtractFormState(CancellationToken cancellationToken) {
        PdfAcroFormXfaInfo? xfa = ExtractAcroFormXfaInfo(cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        IReadOnlyList<PdfFormField> fields = ExtractFormFields(out int javaScriptCount, out long javaScriptBytes, cancellationToken);
        return new FormState(fields, javaScriptCount, javaScriptBytes, xfa);
    }

    /// <summary>Strict reads keep their load-time failure contract by completing every deferred feature in the original order.</summary>
    private void CompleteDeferredFeatures(CancellationToken cancellationToken) {
        _namedDestinationsFeature.Complete(cancellationToken);
        _attachmentsFeature.Complete(cancellationToken);
        _xmpMetadataFeature.Complete(cancellationToken);
        _taggedContentFeature.Complete(cancellationToken);
        _outlinesFeature.Complete(cancellationToken);
        _portfolioFeature.Complete(cancellationToken);
        _formsFeature.Complete(cancellationToken);
    }

    private sealed class FormState {
        internal FormState(IReadOnlyList<PdfFormField> fields, int widgetJavaScriptCount, long widgetJavaScriptBytes, PdfAcroFormXfaInfo? xfa) {
            Fields = fields;
            WidgetJavaScriptCount = widgetJavaScriptCount;
            WidgetJavaScriptBytes = widgetJavaScriptBytes;
            Xfa = xfa;
        }

        internal IReadOnlyList<PdfFormField> Fields { get; }
        internal int WidgetJavaScriptCount { get; }
        internal long WidgetJavaScriptBytes { get; }
        internal PdfAcroFormXfaInfo? Xfa { get; }
    }

    /// <summary>
    /// A document-level read model extracted on first access. The first result, or the first failure,
    /// is retained and returned to every later caller. A caller's own cancellation is never retained.
    /// </summary>
    private sealed class DocumentFeature<T> {
        private readonly object _gate;
        private Func<CancellationToken, T>? _extract;
        private T _value = default!;
        private ExceptionDispatchInfo? _failure;
        private volatile bool _completed;
        private bool _extracting;

        internal DocumentFeature(object gate, Func<CancellationToken, T> extract) {
            _gate = gate;
            _extract = extract;
        }

        internal T GetValue() {
            if (!_completed) Complete(CancellationToken.None);
            _failure?.Throw();
            return _value;
        }

        internal void Complete(CancellationToken cancellationToken) {
            if (_completed) {
                _failure?.Throw();
                return;
            }

            lock (_gate) {
                if (!_completed) {
                    if (_extracting) {
                        throw new InvalidOperationException("A PDF document feature depends on itself while it is being read.");
                    }

                    _extracting = true;
                    try {
                        _value = _extract!(cancellationToken);
                    } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
                        throw;
                    } catch (Exception exception) when (exception is not OutOfMemoryException) {
                        _failure = ExceptionDispatchInfo.Capture(exception);
                    } finally {
                        _extracting = false;
                    }

                    _extract = null;
                    _completed = true;
                }
            }

            _failure?.Throw();
        }
    }
}
