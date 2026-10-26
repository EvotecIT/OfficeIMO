using System;
using System.Management.Automation;

namespace PSWriteOffice.Foundation;

/// <summary>Shared synchronous binding helpers. Content commands deliberately do not declare SupportsShouldProcess.</summary>
public abstract class OfficeCmdlet : PSCmdlet {
    /// <summary>Returns null for an omitted switch, preserving an explicitly false value.</summary>
    protected bool? BoundSwitch(string name) => OfficeSwitch.Get(MyInvocation.BoundParameters, name);
    /// <summary>Reads a feature enabled by default through its NoFeature switch.</summary>
    protected bool FeatureEnabled(string feature) => OfficeSwitch.Enabled(MyInvocation.BoundParameters, feature);
    /// <summary>Resolves one output file, rejecting a directory.</summary>
    protected string ResolveOutputPath(string path) => OfficePath.ResolveOutput(SessionState, path, directory: false);
    /// <summary>Resolves an output folder without creating it.</summary>
    protected string ResolveOutputDirectory(string path) => OfficePath.ResolveOutput(SessionState, path, directory: true);
    /// <summary>Stops the invocation for an invalid command argument.</summary>
    protected void InvalidArgument(Exception exception, object? input, string code = "invalid-argument") => ThrowTerminatingError(OfficeError.Create(exception, code, input, ErrorCategory.InvalidArgument));
    /// <summary>Reports an item failure while allowing subsequent pipeline input.</summary>
    protected void ItemError(Exception exception, string code, object? input) => WriteError(OfficeError.Create(exception, code, input));
}

/// <summary>Binds a live document explicitly or from its enclosing content scope.</summary>
public abstract class OfficeDocumentCmdlet<TDocument> : OfficeCmdlet where TDocument : class {
    /// <summary>Gets or sets an explicit live document; omission uses the enclosing content block.</summary>
    [Parameter(ValueFromPipeline = true)]
    [ValidateNotNull]
    public TDocument? Document { get; set; }
    /// <summary>Resolves the explicit or enclosing document, failing clearly when neither exists.</summary>
    protected TDocument GetDocument() {
        var document = Document ?? OfficeDocumentContext.Find<TDocument>();
        if (document != null) return document;
        InvalidArgument(new ArgumentException("Supply -Document, pipe a live document, or invoke this command inside a New-/Open- content block."), Document, "document-required");
        throw new InvalidOperationException();
    }
}
