using System;
using System.Management.Automation;

namespace PSWriteOffice.Foundation;

/// <summary>Binds live creation separately from atomic content-block publication, using the shared lifecycle.</summary>
public abstract class OfficeNewLifecycleCmdlet<TDocument> : OfficeLifecycleCmdlet<TDocument> where TDocument : class, IDisposable {
    /// <summary>Optional associated destination for a live document; required when a content block saves a file.</summary>
    [Parameter(Position = 0, ParameterSetName = "Live")]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Content")]
    [ValidateNotNullOrEmpty]
    public string? Path { get; set; }

    /// <summary>Authoring block to complete before the document is atomically saved and closed.</summary>
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "Content")]
    [ValidateNotNull]
    public new ScriptBlock? Content { get => base.Content; set => base.Content = value; }

    /// <summary>Return the saved FileInfo after a completed content block.</summary>
    [Parameter(ParameterSetName = "Content")]
    public new SwitchParameter PassThru { get => base.PassThru; set => base.PassThru = value; }

    /// <summary>Explicitly accept reductions reported by the content block's native save.</summary>
    [Parameter(ParameterSetName = "Content")]
    public new SwitchParameter AllowLoss { get => base.AllowLoss; set => base.AllowLoss = value; }

    /// <summary>Open the saved result in its registered application after the content block succeeds.</summary>
    [Parameter(ParameterSetName = "Content")]
    public new SwitchParameter Show { get => base.Show; set => base.Show = value; }
}
