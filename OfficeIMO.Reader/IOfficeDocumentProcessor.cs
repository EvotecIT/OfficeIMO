using System;
using System.Threading.Tasks;

namespace OfficeIMO.Reader;

/// <summary>
/// Processes one rich document result as an ordered step in an <see cref="OfficeDocumentProcessorPipeline"/>.
/// Implementations used by an <see cref="OfficeDocumentReader"/> must be safe for concurrent calls.
/// </summary>
public interface IOfficeDocumentProcessor {
    /// <summary>Stable processor identifier used in diagnostics and execution reports.</summary>
    string Id { get; }

    /// <summary>Processes a document synchronously.</summary>
    OfficeDocumentReadResult Process(
        OfficeDocumentReadResult document,
        OfficeDocumentProcessorContext context);

    /// <summary>Processes a document asynchronously.</summary>
    Task<OfficeDocumentReadResult> ProcessAsync(
        OfficeDocumentReadResult document,
        OfficeDocumentProcessorContext context);
}

/// <summary>
/// Base class for synchronous processors. The asynchronous path delegates to <see cref="Process"/>.
/// </summary>
public abstract class OfficeDocumentProcessorBase : IOfficeDocumentProcessor {
    /// <summary>Creates a processor with a stable identifier.</summary>
    protected OfficeDocumentProcessorBase(string id) {
        if (string.IsNullOrWhiteSpace(id)) {
            throw new ArgumentException("Processor id cannot be empty.", nameof(id));
        }
        Id = id.Trim();
    }

    /// <inheritdoc />
    public string Id { get; }

    /// <inheritdoc />
    public abstract OfficeDocumentReadResult Process(
        OfficeDocumentReadResult document,
        OfficeDocumentProcessorContext context);

    /// <inheritdoc />
    public virtual Task<OfficeDocumentReadResult> ProcessAsync(
        OfficeDocumentReadResult document,
        OfficeDocumentProcessorContext context) {
        return Task.FromResult(Process(document, context));
    }
}

/// <summary>
/// Adapts caller-owned delegates into a typed processor without requiring a custom class.
/// </summary>
public sealed class DelegateOfficeDocumentProcessor : IOfficeDocumentProcessor {
    private readonly Func<OfficeDocumentReadResult, OfficeDocumentProcessorContext, OfficeDocumentReadResult> _process;
    private readonly Func<OfficeDocumentReadResult, OfficeDocumentProcessorContext, Task<OfficeDocumentReadResult>> _processAsync;

    /// <summary>Creates a processor backed by a synchronous delegate.</summary>
    public DelegateOfficeDocumentProcessor(
        string id,
        Func<OfficeDocumentReadResult, OfficeDocumentProcessorContext, OfficeDocumentReadResult> process) {
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Processor id cannot be empty.", nameof(id));
        Id = id.Trim();
        _process = process ?? throw new ArgumentNullException(nameof(process));
        _processAsync = (document, context) => Task.FromResult(_process(document, context));
    }

    /// <summary>Creates a processor with explicit synchronous and asynchronous delegates.</summary>
    public DelegateOfficeDocumentProcessor(
        string id,
        Func<OfficeDocumentReadResult, OfficeDocumentProcessorContext, OfficeDocumentReadResult> process,
        Func<OfficeDocumentReadResult, OfficeDocumentProcessorContext, Task<OfficeDocumentReadResult>> processAsync) {
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Processor id cannot be empty.", nameof(id));
        Id = id.Trim();
        _process = process ?? throw new ArgumentNullException(nameof(process));
        _processAsync = processAsync ?? throw new ArgumentNullException(nameof(processAsync));
    }

    /// <inheritdoc />
    public string Id { get; }

    /// <inheritdoc />
    public OfficeDocumentReadResult Process(
        OfficeDocumentReadResult document,
        OfficeDocumentProcessorContext context) {
        return _process(document, context);
    }

    /// <inheritdoc />
    public Task<OfficeDocumentReadResult> ProcessAsync(
        OfficeDocumentReadResult document,
        OfficeDocumentProcessorContext context) {
        return _processAsync(document, context);
    }
}
