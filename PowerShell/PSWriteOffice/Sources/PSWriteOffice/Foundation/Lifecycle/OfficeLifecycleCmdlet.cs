using System;
using System.IO;
using System.Management.Automation;
using System.Threading;
using OfficeIMO;

namespace PSWriteOffice.Foundation;

/// <summary>Provides the New-/Open- content-block contract without adding content-item confirmation.</summary>
/// <remarks>Concrete lifecycle commands declare SupportsShouldProcess and their format-specific input sets.</remarks>
public abstract class OfficeLifecycleCmdlet<TDocument> : OfficeCmdlet, IDisposable where TDocument : class, IDisposable {
    private readonly CancellationTokenSource _stop = new();
    /// <summary>Gets or sets a content block; omission returns the live document.</summary>
    [Parameter(Position = 1)]
    [ValidateNotNull]
    public ScriptBlock? Content { get; set; }
    /// <summary>Gets or sets whether a completed content block returns the saved FileInfo.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }
    /// <summary>Gets or sets whether native-save loss is explicitly accepted.</summary>
    [Parameter]
    public SwitchParameter AllowLoss { get; set; }
    /// <summary>Gets or sets whether the saved file opens in the system default application.</summary>
    [Parameter]
    public SwitchParameter Show { get; set; }

    /// <summary>Creates/loads a live document or runs and atomically publishes a content block on the pipeline thread.</summary>
    /// <param name="path">The provider path. A path is optional only for a new live document.</param>
    /// <param name="isNew">Whether the operation creates a document instead of opening an existing input.</param>
    /// <param name="force">Whether New may replace an existing file.</param>
    /// <param name="factory">Creates an in-memory/new staged document or loads without autosave.</param>
    /// <param name="save">Saves natively to the private path, honoring AllowLoss and returning any save diagnostics.</param>
    protected void RunDocument(string? path, bool isNew, bool force, Func<string?, TDocument> factory,
        Func<TDocument, string, bool, OfficeConversionReport?> save) {
        string? resolved = null;
        try {
            if (path != null) resolved = ResolveOutputPath(path);
            if ((!isNew || Content != null) && resolved == null) throw new ArgumentException("A path is required when opening a document or using a content block.");
            if (Content == null && (PassThru || Show)) throw new ArgumentException("PassThru and Show require a content block; without a block the live document is returned.");
            if (resolved != null) {
                if (!isNew && !File.Exists(resolved)) throw new FileNotFoundException("The input does not exist.", resolved);
                OfficePath.ValidateDestination(resolved, isNew && !force);
            }
        } catch (ArgumentException error) { InvalidArgument(error, path); return; }
        // Opening a live document and editing in memory have no filesystem effect.
        if (Content == null) { WriteObject(factory(resolved)); return; }
        if (!ShouldProcess(resolved!, isNew ? "Create document" : "Save document")) return;
        try {
            FileInfo result = OfficeContentLifecycle.Run(resolved!, isNew, isNew && !force, AllowLoss,
                p => factory(p), _ => Content.Invoke(), save, _stop.Token, error => ThrowTerminatingError(OfficeError.NativeSave(error, resolved)));
            if (Show) OfficeShow.Open(result.FullName);
            if (PassThru) WriteObject(result);
        } catch (PipelineStoppedException) { throw; }
        catch (OperationCanceledException) when (_stop.IsCancellationRequested) { throw new PipelineStoppedException(); }
        // Block failures propagate unchanged; never turn a terminating authoring error into a success result.
    }
    /// <summary>Releases the lifecycle cancellation source after the command finishes.</summary>
    public void Dispose() => _stop.Dispose();
    /// <inheritdoc />
    protected override void StopProcessing() => _stop.Cancel();
}
