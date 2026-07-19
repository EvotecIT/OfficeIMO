using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Reader;

/// <summary>Immutable processor pipeline that executes processors in registration order.</summary>
public sealed class OfficeDocumentProcessorPipeline {
    private readonly IOfficeDocumentProcessor[] _processors;
    private readonly ReadOnlyCollection<IOfficeDocumentProcessor> _processorView;

    internal OfficeDocumentProcessorPipeline(IOfficeDocumentProcessor[] processors) {
        _processors = processors ?? throw new ArgumentNullException(nameof(processors));
        _processorView = Array.AsReadOnly(_processors);
    }

    /// <summary>Empty reusable pipeline.</summary>
    public static OfficeDocumentProcessorPipeline Empty { get; } = new OfficeDocumentProcessorPipeline(Array.Empty<IOfficeDocumentProcessor>());

    /// <summary>Configured processors in deterministic execution order.</summary>
    public IReadOnlyList<IOfficeDocumentProcessor> Processors => _processorView;

    /// <summary>Configured processor count.</summary>
    public int Count => _processors.Length;

    /// <summary>Processes a document synchronously in configured order.</summary>
    public OfficeDocumentProcessingResult Process(
        OfficeDocumentReadResult document,
        OfficeDocumentProcessingOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        OfficeDocumentProcessorFailureBehavior failureBehavior = Normalize(options).FailureBehavior;
        var steps = new OfficeDocumentProcessorStepResult[_processors.Length];
        var retainedFailureDiagnostics = new List<OfficeDocumentDiagnostic>();
        OfficeDocumentReadResult current = document;

        for (int index = 0; index < _processors.Length; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            IOfficeDocumentProcessor processor = _processors[index];
            var context = new OfficeDocumentProcessorContext(processor.Id, index, _processors.Length, cancellationToken);
            try {
                if (!(processor is ISynchronousOfficeDocumentProcessor synchronousProcessor)) {
                    throw new InvalidOperationException(
                        $"Document processor '{processor.Id}' is asynchronous-only. Use ProcessAsync(...).");
                }
                current = synchronousProcessor.Process(current, context)
                    ?? throw new InvalidOperationException("Processor returned a null document result.");
                AppendDiagnostics(current, retainedFailureDiagnostics);
                steps[index] = Completed(processor, index);
            } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
                throw;
            } catch (Exception exception) {
                if (failureBehavior == OfficeDocumentProcessorFailureBehavior.Throw) {
                    throw new OfficeDocumentProcessorException(processor.Id, index, exception);
                }
                OfficeDocumentDiagnostic diagnostic = BuildFailureDiagnostic(processor, index, exception, failureBehavior);
                AppendDiagnostic(current, diagnostic);
                retainedFailureDiagnostics.Add(diagnostic);
                steps[index] = Failed(processor, index, diagnostic);
                if (failureBehavior == OfficeDocumentProcessorFailureBehavior.StopWithDiagnostic) {
                    MarkRemainingSkipped(steps, index + 1);
                    break;
                }
            }
        }

        return new OfficeDocumentProcessingResult(current, steps);
    }

    /// <summary>Processes a document asynchronously in configured order.</summary>
    public async Task<OfficeDocumentProcessingResult> ProcessAsync(
        OfficeDocumentReadResult document,
        OfficeDocumentProcessingOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        OfficeDocumentProcessorFailureBehavior failureBehavior = Normalize(options).FailureBehavior;
        var steps = new OfficeDocumentProcessorStepResult[_processors.Length];
        var retainedFailureDiagnostics = new List<OfficeDocumentDiagnostic>();
        OfficeDocumentReadResult current = document;

        for (int index = 0; index < _processors.Length; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            IOfficeDocumentProcessor processor = _processors[index];
            var context = new OfficeDocumentProcessorContext(processor.Id, index, _processors.Length, cancellationToken);
            try {
                if (processor is IAsyncOfficeDocumentProcessor asynchronousProcessor) {
                    Task<OfficeDocumentReadResult> task = asynchronousProcessor.ProcessAsync(current, context)
                        ?? throw new InvalidOperationException("Processor returned a null asynchronous operation.");
                    current = await task.ConfigureAwait(false)
                        ?? throw new InvalidOperationException("Processor returned a null document result.");
                } else if (processor is ISynchronousOfficeDocumentProcessor synchronousProcessor) {
                    current = synchronousProcessor.Process(current, context)
                        ?? throw new InvalidOperationException("Processor returned a null document result.");
                } else {
                    throw new InvalidOperationException(
                        $"Document processor '{processor.Id}' implements neither the synchronous nor asynchronous processing contract.");
                }
                AppendDiagnostics(current, retainedFailureDiagnostics);
                steps[index] = Completed(processor, index);
            } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
                throw;
            } catch (Exception exception) {
                if (failureBehavior == OfficeDocumentProcessorFailureBehavior.Throw) {
                    throw new OfficeDocumentProcessorException(processor.Id, index, exception);
                }
                OfficeDocumentDiagnostic diagnostic = BuildFailureDiagnostic(processor, index, exception, failureBehavior);
                AppendDiagnostic(current, diagnostic);
                retainedFailureDiagnostics.Add(diagnostic);
                steps[index] = Failed(processor, index, diagnostic);
                if (failureBehavior == OfficeDocumentProcessorFailureBehavior.StopWithDiagnostic) {
                    MarkRemainingSkipped(steps, index + 1);
                    break;
                }
            }
        }

        return new OfficeDocumentProcessingResult(current, steps);
    }

    internal static string ValidateProcessorId(string? id, string paramName) {
        if (string.IsNullOrWhiteSpace(id)) {
            throw new ArgumentException("Processor id cannot be empty.", paramName);
        }
        string normalized = id!.Trim();
        if (!string.Equals(id, normalized, StringComparison.Ordinal)) {
            throw new ArgumentException("Processor id cannot contain leading or trailing whitespace.", paramName);
        }
        return normalized;
    }

    private static OfficeDocumentProcessingOptions Normalize(OfficeDocumentProcessingOptions? options) {
        OfficeDocumentProcessingOptions effective = options?.Clone() ?? new OfficeDocumentProcessingOptions();
        if (!Enum.IsDefined(typeof(OfficeDocumentProcessorFailureBehavior), effective.FailureBehavior)) {
            throw new ArgumentOutOfRangeException(nameof(options), effective.FailureBehavior, "Unknown processor failure behavior.");
        }
        return effective;
    }

    private OfficeDocumentProcessorStepResult Completed(IOfficeDocumentProcessor processor, int index) =>
        new OfficeDocumentProcessorStepResult(processor.Id, index, OfficeDocumentProcessorStepStatus.Completed);

    private OfficeDocumentProcessorStepResult Failed(
        IOfficeDocumentProcessor processor,
        int index,
        OfficeDocumentDiagnostic diagnostic) =>
        new OfficeDocumentProcessorStepResult(processor.Id, index, OfficeDocumentProcessorStepStatus.Failed, diagnostic);

    private void MarkRemainingSkipped(OfficeDocumentProcessorStepResult[] steps, int startIndex) {
        for (int index = startIndex; index < _processors.Length; index++) {
            steps[index] = new OfficeDocumentProcessorStepResult(
                _processors[index].Id,
                index,
                OfficeDocumentProcessorStepStatus.Skipped);
        }
    }

    private static OfficeDocumentDiagnostic BuildFailureDiagnostic(
        IOfficeDocumentProcessor processor,
        int index,
        Exception exception,
        OfficeDocumentProcessorFailureBehavior failureBehavior) {
        return new OfficeDocumentDiagnostic {
            Severity = OfficeDocumentDiagnosticSeverity.Error,
            Category = OfficeDocumentDiagnosticCategory.General,
            Code = "processor-failed",
            Message = $"Document processor '{processor.Id}' failed: {exception.Message}",
            Source = "officeimo.reader.processor." + processor.Id,
            IsRecoverable = failureBehavior == OfficeDocumentProcessorFailureBehavior.ContinueWithDiagnostic,
            Attributes = new SortedDictionary<string, string>(StringComparer.Ordinal) {
                ["exceptionType"] = exception.GetType().FullName ?? exception.GetType().Name,
                ["processorId"] = processor.Id,
                ["processorIndex"] = index.ToString(CultureInfo.InvariantCulture)
            }
        };
    }

    private static void AppendDiagnostic(OfficeDocumentReadResult document, OfficeDocumentDiagnostic diagnostic) {
        IReadOnlyList<OfficeDocumentDiagnostic>? existing = document.Diagnostics;
        int count = existing?.Count ?? 0;
        var diagnostics = new OfficeDocumentDiagnostic[count + 1];
        for (int index = 0; index < count; index++) diagnostics[index] = existing![index];
        diagnostics[count] = diagnostic;
        document.Diagnostics = diagnostics;
    }

    private static void AppendDiagnostics(
        OfficeDocumentReadResult document,
        IReadOnlyList<OfficeDocumentDiagnostic> diagnostics) {
        for (int index = 0; index < diagnostics.Count; index++) {
            OfficeDocumentDiagnostic diagnostic = diagnostics[index];
            if (!ContainsProcessorFailure(document.Diagnostics, diagnostic)) {
                AppendDiagnostic(document, diagnostic);
            }
        }
    }

    private static bool ContainsProcessorFailure(
        IReadOnlyList<OfficeDocumentDiagnostic>? diagnostics,
        OfficeDocumentDiagnostic expected) {
        if (diagnostics == null) return false;
        expected.Attributes.TryGetValue("processorId", out string? expectedProcessorId);
        expected.Attributes.TryGetValue("processorIndex", out string? expectedProcessorIndex);
        for (int index = 0; index < diagnostics.Count; index++) {
            OfficeDocumentDiagnostic? candidate = diagnostics[index];
            if (candidate == null) continue;
            if (ReferenceEquals(candidate, expected)) return true;
            if (!string.Equals(candidate.Code, expected.Code, StringComparison.Ordinal) ||
                candidate.Attributes == null ||
                !candidate.Attributes.TryGetValue("processorId", out string? processorId) ||
                !candidate.Attributes.TryGetValue("processorIndex", out string? processorIndex)) {
                continue;
            }
            if (string.Equals(processorId, expectedProcessorId, StringComparison.Ordinal) &&
                string.Equals(processorIndex, expectedProcessorIndex, StringComparison.Ordinal)) {
                return true;
            }
        }
        return false;
    }
}
