namespace OfficeIMO.Pdf;

/// <summary>
/// Immutable evidence for one create, open, mutation, or output stage.
/// </summary>
public sealed class PdfPipelineStep {
    internal PdfPipelineStep(
        PdfPipelineStepKind kind,
        string operation,
        bool succeeded,
        PdfArtifactSnapshot? input,
        PdfArtifactSnapshot? output,
        TimeSpan? duration,
        PdfMutationOperation? mutationOperation,
        PdfMutationExecutionMode? executionMode,
        IReadOnlyList<string>? diagnostics = null) {
        Kind = kind;
        Operation = operation;
        Succeeded = succeeded;
        Input = input;
        Output = output;
        Duration = duration;
        MutationOperation = mutationOperation;
        ExecutionMode = executionMode;
        Diagnostics = Array.AsReadOnly((diagnostics ?? Array.Empty<string>()).ToArray());
    }

    /// <summary>Pipeline stage category.</summary>
    public PdfPipelineStepKind Kind { get; }

    /// <summary>Short operation name suitable for logs and reports.</summary>
    public string Operation { get; }

    /// <summary>True when the stage completed.</summary>
    public bool Succeeded { get; }

    /// <summary>Input artifact evidence when the stage consumed a PDF.</summary>
    public PdfArtifactSnapshot? Input { get; }

    /// <summary>Output artifact evidence when the stage produced a PDF.</summary>
    public PdfArtifactSnapshot? Output { get; }

    /// <summary>Measured execution duration, or null when an older operation path could not capture it.</summary>
    public TimeSpan? Duration { get; }

    /// <summary>Shared mutation family selected for an existing-document operation.</summary>
    public PdfMutationOperation? MutationOperation { get; }

    /// <summary>Observed full-rewrite or append-only execution mode for an existing-document operation.</summary>
    public PdfMutationExecutionMode? ExecutionMode { get; }

    /// <summary>Diagnostics captured for this stage.</summary>
    public IReadOnlyList<string> Diagnostics { get; }
}
