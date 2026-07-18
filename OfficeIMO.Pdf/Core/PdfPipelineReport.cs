namespace OfficeIMO.Pdf;

/// <summary>
/// Immutable end-to-end history for a <see cref="PdfDocument"/> and its output operation.
/// </summary>
public sealed class PdfPipelineReport {
    private readonly IReadOnlyList<PdfPipelineStep> _steps;

    internal PdfPipelineReport(IEnumerable<PdfPipelineStep> steps) {
        Guard.NotNull(steps, nameof(steps));
        _steps = Array.AsReadOnly(steps.ToArray());
    }

    /// <summary>Ordered create/open, mutation, and output stages.</summary>
    public IReadOnlyList<PdfPipelineStep> Steps => _steps;

    /// <summary>True when every recorded stage completed.</summary>
    public bool Succeeded => _steps.All(step => step.Succeeded);

    /// <summary>True when at least one recorded stage failed.</summary>
    public bool HasFailures => !Succeeded;

    /// <summary>First captured PDF artifact, when the pipeline opened existing bytes.</summary>
    public PdfArtifactSnapshot? Input => _steps.Select(step => step.Input).FirstOrDefault(snapshot => snapshot is not null);

    /// <summary>Most recent captured PDF artifact.</summary>
    public PdfArtifactSnapshot? Output => _steps.Reverse().Select(step => step.Output).FirstOrDefault(snapshot => snapshot is not null);

    /// <summary>Total duration of stages whose execution time was measured.</summary>
    public TimeSpan TotalDuration => TimeSpan.FromTicks(_steps
        .Where(step => step.Duration.HasValue)
        .Sum(step => step.Duration!.Value.Ticks));

    /// <summary>Returns this report or throws when a recorded stage failed.</summary>
    public PdfPipelineReport RequireSuccess() {
        if (Succeeded) {
            return this;
        }

        string[] diagnostics = _steps
            .Where(step => !step.Succeeded)
            .SelectMany(step => step.Diagnostics.Count == 0
                ? new[] { step.Operation + " failed." }
                : step.Diagnostics)
            .ToArray();
        throw new InvalidOperationException(
            diagnostics.Length == 0
                ? "The PDF pipeline did not complete."
                : "The PDF pipeline did not complete. " + string.Join(" ", diagnostics));
    }

    internal PdfPipelineReport Append(PdfPipelineStep step) {
        Guard.NotNull(step, nameof(step));
        return new PdfPipelineReport(_steps.Concat(new[] { step }));
    }

    internal static PdfPipelineReport Created() {
        return new PdfPipelineReport(new[] {
            new PdfPipelineStep(
                PdfPipelineStepKind.Create,
                "Create",
                succeeded: true,
                input: null,
                output: null,
                duration: null,
                mutationOperation: null,
                executionMode: null)
        });
    }

    internal static PdfPipelineReport Empty() => new PdfPipelineReport(Array.Empty<PdfPipelineStep>());

    internal static PdfPipelineReport Opened(byte[] bytes, PdfReadOptions? readOptions) {
        PdfArtifactSnapshot artifact = PdfArtifactSnapshot.Capture(bytes, readOptions);
        return new PdfPipelineReport(new[] {
            new PdfPipelineStep(
                PdfPipelineStepKind.Open,
                "Open",
                succeeded: true,
                input: artifact,
                output: artifact,
                duration: null,
                mutationOperation: null,
                executionMode: null)
        });
    }
}
