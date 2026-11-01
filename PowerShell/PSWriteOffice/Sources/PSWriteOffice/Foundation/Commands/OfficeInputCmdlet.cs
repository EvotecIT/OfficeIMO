using System;
using System.Collections.Generic;
using System.Management.Automation;
using System.Threading.Tasks;

namespace PSWriteOffice.Foundation;

/// <summary>Provides provider-resolved batch inputs and the existing async pipeline/cancellation bridge.</summary>
public abstract class OfficeInputCmdlet : AsyncPSCmdlet {
    /// <summary>Gets or sets input paths with provider wildcard expansion.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    [SupportsWildcards]
    [ValidateNotNullOrEmpty]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>Gets or sets input paths without wildcard expansion, binding FileInfo.FullName and provider PSPath.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("PSPath", "FullName")]
    [ValidateNotNullOrEmpty]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();
    /// <summary>Gets or sets the maximum number of concurrent background operations.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [ValidateRange(1, 1024)]
    public int ThrottleLimit { get; set; } = 4;

    /// <summary>Resolves each input before any await, reporting missing items independently.</summary>
    protected IReadOnlyList<string> ResolveInputs(bool allowDirectory = false) {
        bool literal = ParameterSetName == "LiteralPath";
        return ResolveInputs(literal ? LiteralPath : Path, literal, allowDirectory);
    }

    /// <summary>Resolves captured pipeline path properties with the same per-item errors as named path parameters.</summary>
    internal IReadOnlyList<string> ResolveInputs(IEnumerable<string> inputs, bool literal, bool allowDirectory = false) {
        var result = new List<string>();
        foreach (string input in inputs) {
            try { result.AddRange(OfficePath.ResolveInput(SessionState, input, literal, allowDirectory,
                (error, item) => WriteError(OfficeError.Create(error, "input-unreadable", item)))); }
            catch (ArgumentException error) { ThrowTerminatingError(OfficeError.Create(error, "invalid-argument", input, ErrorCategory.InvalidArgument)); }
            catch (Exception error) when ((error is System.IO.IOException or RuntimeException or UnauthorizedAccessException) &&
                error is not (ActionPreferenceStopException or PipelineStoppedException)) {
                WriteError(OfficeError.Create(error, "input-unreadable", input));
            }
        }
        return result;
    }
    /// <summary>Runs bounded background work and streams completions through the canonical pipeline bridge.</summary>
    /// <remarks>Call after provider resolution and confirmation. Workers must not invoke script blocks, SessionState, or raw PSCmdlet APIs.</remarks>
    protected Task RunBatchAsync<T>(IEnumerable<T> inputs, Func<T, System.Threading.CancellationToken, Task<object?>> worker,
        string activity, string diagnosticCode = "operation-failed") => OfficeBatch.RunAsync(inputs, worker, ThrottleLimit,
            CancelToken, this, activity, diagnosticCode);
}
