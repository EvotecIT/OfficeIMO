using System;
using System.Collections;
using System.Management.Automation;
using System.Security;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Creates a proof report for user-visible signals preserved by a PDF rewrite.</summary>
[Cmdlet(VerbsDiagnostic.Test, "OfficePdfRewrite", DefaultParameterSetName = "Path")]
[OutputType(typeof(PdfRewritePreservationReport))]
public sealed class TestOfficePdfRewriteCommand : OfficeCmdlet
{
    /// <summary>Original PDF path.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Path")]
    public string ReferencePath { get; set; } = string.Empty;

    /// <summary>Rewritten PDF path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string DifferencePath { get; set; } = string.Empty;

    /// <summary>Optional required preservation signals and limits.</summary>
    [Parameter]
    public object[]? Options { get; set; }

    /// <summary>Throw when preservation checks find a mismatch.</summary>
    [Parameter]
    public SwitchParameter FailOnLoss { get; set; }

    /// <summary>Password used to authenticate the original PDF.</summary>
    [Parameter]
    [OfficePasswordTransformation]
    public SecureString? Password { get; set; }

    /// <summary>After authentication, explicitly ignore restrictions on the original PDF.</summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <summary>Password used to authenticate the rewritten PDF.</summary>
    [Parameter]
    [OfficePasswordTransformation]
    public SecureString? DifferencePassword { get; set; }

    /// <summary>After authentication, explicitly ignore restrictions on the rewritten PDF.</summary>
    [Parameter]
    public SwitchParameter IgnoreDifferencePermissionRestrictions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var binder = new OfficeOptionsTransformationAttribute(typeof(PdfRewritePreservationOptions));
        var options = new PdfRewritePreservationOptions();
        foreach (object item in Options ?? Array.Empty<object>()) {
            object? value = item is PSObject wrapped ? wrapped.BaseObject : item;
            if (value is PdfRewritePreservationOptions native) options = native.Clone();
            else if (value is IDictionary fields) options = (PdfRewritePreservationOptions)binder.ApplyTo(
                (EngineIntrinsics)SessionState.PSVariable.GetValue("ExecutionContext"), options, fields);
            else throw new PSArgumentException("Options accepts PdfRewritePreservationOptions or its property hashtables.");
        }
        var originalReadOptions = PdfCommandUtilities.CreateReadOptions(
            options.OriginalReadOptions,
            Password,
            IgnorePermissionRestrictions.IsPresent);
        var rewrittenReadOptions = PdfCommandUtilities.CreateReadOptions(
            options.RewrittenReadOptions,
            DifferencePassword,
            IgnoreDifferencePermissionRestrictions.IsPresent);
        using var original = PdfCommandUtilities.LoadDocument(
            SessionState.Path.GetUnresolvedProviderPathFromPSPath(ReferencePath),
            originalReadOptions);
        using var rewritten = PdfCommandUtilities.LoadDocument(
            SessionState.Path.GetUnresolvedProviderPathFromPSPath(DifferencePath),
            rewrittenReadOptions);

        options.OriginalReadOptions = originalReadOptions;
        options.RewrittenReadOptions = rewrittenReadOptions;
        PdfRewritePreservationReport report = original.Proof.AssessRewritePreservation(rewritten, options);

        if (FailOnLoss.IsPresent) report.ThrowIfFailed();
        WriteObject(report);
    }
}
