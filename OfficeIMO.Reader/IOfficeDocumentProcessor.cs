using System;
using System.Threading.Tasks;

namespace OfficeIMO.Reader;

/// <summary>
/// Identifies one rich-document processing step in an <see cref="OfficeDocumentProcessorPipeline"/>.
/// Processors implement the synchronous contract, asynchronous contract, or both.
/// </summary>
public interface IOfficeDocumentProcessor {
    /// <summary>Stable processor identifier used in diagnostics and execution reports.</summary>
    string Id { get; }
}

/// <summary>Synchronous document processor capability.</summary>
public interface ISynchronousOfficeDocumentProcessor : IOfficeDocumentProcessor {
    /// <summary>Processes a document synchronously.</summary>
    OfficeDocumentReadResult Process(
        OfficeDocumentReadResult document,
        OfficeDocumentProcessorContext context);
}

/// <summary>Asynchronous document processor capability for real asynchronous work.</summary>
public interface IAsyncOfficeDocumentProcessor : IOfficeDocumentProcessor {
    /// <summary>Processes a document asynchronously.</summary>
    Task<OfficeDocumentReadResult> ProcessAsync(
        OfficeDocumentReadResult document,
        OfficeDocumentProcessorContext context);
}

/// <summary>
/// Base class for synchronous processors.
/// </summary>
public abstract class OfficeDocumentProcessorBase : ISynchronousOfficeDocumentProcessor {
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

}

/// <summary>
/// Adapts caller-owned delegates into a typed processor without requiring a custom class.
/// </summary>
public sealed class DelegateOfficeDocumentProcessor : ISynchronousOfficeDocumentProcessor {
    private readonly Func<OfficeDocumentReadResult, OfficeDocumentProcessorContext, OfficeDocumentReadResult> _process;

    /// <summary>Creates a processor backed by a synchronous delegate.</summary>
    public DelegateOfficeDocumentProcessor(
        string id,
        Func<OfficeDocumentReadResult, OfficeDocumentProcessorContext, OfficeDocumentReadResult> process) {
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Processor id cannot be empty.", nameof(id));
        Id = id.Trim();
        _process = process ?? throw new ArgumentNullException(nameof(process));
    }

    /// <inheritdoc />
    public string Id { get; }

    /// <inheritdoc />
    public OfficeDocumentReadResult Process(
        OfficeDocumentReadResult document,
        OfficeDocumentProcessorContext context) {
        return _process(document, context);
    }

}

/// <summary>Adapts a caller-owned asynchronous delegate into a typed processor.</summary>
public sealed class DelegateAsyncOfficeDocumentProcessor : IAsyncOfficeDocumentProcessor {
    private readonly Func<OfficeDocumentReadResult, OfficeDocumentProcessorContext, Task<OfficeDocumentReadResult>> _process;

    /// <summary>Creates an asynchronous processor backed by a real asynchronous delegate.</summary>
    public DelegateAsyncOfficeDocumentProcessor(
        string id,
        Func<OfficeDocumentReadResult, OfficeDocumentProcessorContext, Task<OfficeDocumentReadResult>> process) {
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Processor id cannot be empty.", nameof(id));
        Id = id.Trim();
        _process = process ?? throw new ArgumentNullException(nameof(process));
    }

    /// <inheritdoc />
    public string Id { get; }

    /// <inheritdoc />
    public Task<OfficeDocumentReadResult> ProcessAsync(
        OfficeDocumentReadResult document,
        OfficeDocumentProcessorContext context) => _process(document, context);
}
