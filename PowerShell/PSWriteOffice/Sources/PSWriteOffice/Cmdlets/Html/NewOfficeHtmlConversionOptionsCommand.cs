using System;
using System.Management.Automation;
using OfficeIMO.Html;

namespace PSWriteOffice.Cmdlets.Html;

/// <summary>Creates discoverable parsing, trust, and document settings for HTML conversion.</summary>
/// <example>
///   <summary>Resolve relative resources from a trusted report directory.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$document = New-OfficeHtmlConversionOptions -BaseUri (Resolve-Path .\Assets) -NoBodyContentsOnly
/// Export-OfficeHtmlImage -Path .\Report.html -OutputPath .\Report.svg -DocumentOptions $document</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeHtmlConversionOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(HtmlConversionDocumentOptions))]
public sealed class NewOfficeHtmlConversionOptionsCommand : PSCmdlet {
    /// <summary>Built-in conversion profile.</summary>
    [Parameter(ParameterSetName = "Default")] public HtmlConversionProfile? Profile { get; set; }
    /// <summary>Input trust level.</summary>
    [Parameter] public HtmlInputTrust? Trust { get; set; }
    /// <summary>Base URI used to resolve relative references.</summary>
    [Parameter] public string? BaseUri { get; set; }
    /// <summary>Convert only body contents.</summary>
    [Parameter] public SwitchParameter NoBodyContentsOnly { get; set; }
    /// <summary>Retain normalized HTML in the conversion document.</summary>
    [Parameter] public SwitchParameter NoNormalizedHtml { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new HtmlConversionDocumentOptions();
        if (Profile.HasValue) options.Profile = Profile.Value;
        if (Trust.HasValue) options.Trust = Trust.Value;
        if (!string.IsNullOrWhiteSpace(BaseUri)) options.BaseUri = HtmlOptionsCommandUtilities.NormalizeBaseUri(SessionState, BaseUri!);
        if (IsBound(nameof(NoBodyContentsOnly))) options.UseBodyContentsOnly = !NoBodyContentsOnly.IsPresent;
        if (IsBound(nameof(NoNormalizedHtml))) options.IncludeNormalizedHtml = !NoNormalizedHtml.IsPresent;
        WriteObject(options);
    }
    private bool IsBound(string name) => MyInvocation.BoundParameters.ContainsKey(name);
}
