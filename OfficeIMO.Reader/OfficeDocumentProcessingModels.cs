using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Reader;

/// <summary>Controls how a processor pipeline responds to processor failures.</summary>
public enum OfficeDocumentProcessorFailureBehavior {
    /// <summary>Throw an <see cref="OfficeDocumentProcessorException"/> immediately.</summary>
    Throw = 0,
    /// <summary>Add a structured diagnostic and continue with the next processor.</summary>
    ContinueWithDiagnostic,
    /// <summary>Add a structured diagnostic and skip the remaining processors.</summary>
    StopWithDiagnostic
}

/// <summary>Execution options for an ordered document processor pipeline.</summary>
public sealed class OfficeDocumentProcessingOptions {
    /// <summary>Failure behavior. The default preserves fail-fast behavior.</summary>
    public OfficeDocumentProcessorFailureBehavior FailureBehavior { get; set; } = OfficeDocumentProcessorFailureBehavior.Throw;

    internal OfficeDocumentProcessingOptions Clone() => new OfficeDocumentProcessingOptions {
        FailureBehavior = FailureBehavior
    };
}

/// <summary>Immutable context for one processor invocation.</summary>
public sealed class OfficeDocumentProcessorContext {
    internal OfficeDocumentProcessorContext(
        string processorId,
        int processorIndex,
        int processorCount,
        CancellationToken cancellationToken) {
        ProcessorId = processorId;
        ProcessorIndex = processorIndex;
        ProcessorCount = processorCount;
        CancellationToken = cancellationToken;
    }

    /// <summary>Stable processor identifier.</summary>
    public string ProcessorId { get; }

    /// <summary>Zero-based processor index in configured execution order.</summary>
    public int ProcessorIndex { get; }

    /// <summary>Total configured processor count.</summary>
    public int ProcessorCount { get; }

    /// <summary>Cancellation token for this processing operation.</summary>
    public CancellationToken CancellationToken { get; }
}

/// <summary>Status of one configured processor step.</summary>
public enum OfficeDocumentProcessorStepStatus {
    /// <summary>The processor completed and returned a document.</summary>
    Completed = 0,
    /// <summary>The processor failed and failure policy retained a diagnostic.</summary>
    Failed,
    /// <summary>The processor was not invoked because a previous step stopped the pipeline.</summary>
    Skipped
}

/// <summary>Deterministic execution record for one processor step.</summary>
public sealed class OfficeDocumentProcessorStepResult {
    internal OfficeDocumentProcessorStepResult(
        string processorId,
        int processorIndex,
        OfficeDocumentProcessorStepStatus status,
        OfficeDocumentDiagnostic? diagnostic = null) {
        ProcessorId = processorId;
        ProcessorIndex = processorIndex;
        Status = status;
        Diagnostic = diagnostic;
    }

    /// <summary>Stable processor identifier.</summary>
    public string ProcessorId { get; }

    /// <summary>Zero-based configured execution index.</summary>
    public int ProcessorIndex { get; }

    /// <summary>Execution status.</summary>
    public OfficeDocumentProcessorStepStatus Status { get; }

    /// <summary>Structured failure diagnostic when <see cref="Status"/> is <see cref="OfficeDocumentProcessorStepStatus.Failed"/>.</summary>
    public OfficeDocumentDiagnostic? Diagnostic { get; }
}

/// <summary>Processed document plus deterministic per-step execution evidence.</summary>
public sealed class OfficeDocumentProcessingResult {
    internal OfficeDocumentProcessingResult(
        OfficeDocumentReadResult document,
        IReadOnlyList<OfficeDocumentProcessorStepResult> steps) {
        Document = document;
        Steps = steps;
    }

    /// <summary>Final document returned by the pipeline.</summary>
    public OfficeDocumentReadResult Document { get; }

    /// <summary>One record per configured processor, in execution order.</summary>
    public IReadOnlyList<OfficeDocumentProcessorStepResult> Steps { get; }

    /// <summary>True when every configured processor completed.</summary>
    public bool Succeeded {
        get {
            for (int index = 0; index < Steps.Count; index++) {
                if (Steps[index].Status != OfficeDocumentProcessorStepStatus.Completed) return false;
            }
            return true;
        }
    }
}

/// <summary>Wraps a fail-fast processor failure with stable processor identity and order.</summary>
public sealed class OfficeDocumentProcessorException : Exception {
    internal OfficeDocumentProcessorException(string processorId, int processorIndex, Exception innerException)
        : base($"Document processor '{processorId}' failed at index {processorIndex}: {innerException.Message}", innerException) {
        ProcessorId = processorId;
        ProcessorIndex = processorIndex;
    }

    /// <summary>Processor that failed.</summary>
    public string ProcessorId { get; }

    /// <summary>Zero-based configured processor index.</summary>
    public int ProcessorIndex { get; }
}
