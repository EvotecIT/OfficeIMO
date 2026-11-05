using System.Management.Automation;
using OfficeIMO.OpenDocument;
using OfficeIMO.Word.OpenDocument;

namespace PSWriteOffice.Cmdlets.OpenDocument;

/// <summary>Creates Word/OpenDocument conversion settings.</summary>
/// <example>
///   <summary>Include Word images and headers during conversion.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeWordOpenDocumentOptions -NoImages -NoHeadersAndFooters
/// Convert-OfficeDocument -Path .\Report.docx -To OpenDocument -OutputPath .\Report.odt -Options $options</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeWordOpenDocumentOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(WordOpenDocumentConversionOptions))]
public sealed class NewOfficeWordOpenDocumentOptionsCommand : PSCmdlet {
    /// <summary>Whether conversion loss is reported or rejected.</summary>
    [Parameter(ParameterSetName = "Default")] public OdfConversionLossPolicy? LossPolicy { get; set; }
    /// <summary>Copy supported inline images.</summary>
    [Parameter] public SwitchParameter NoImages { get; set; }
    /// <summary>Copy default headers and footers.</summary>
    [Parameter] public SwitchParameter NoHeadersAndFooters { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new WordOpenDocumentConversionOptions();
        if (LossPolicy.HasValue) options.LossPolicy = LossPolicy.Value;
        if (IsBound(nameof(NoImages))) options.IncludeImages = !NoImages.IsPresent;
        if (IsBound(nameof(NoHeadersAndFooters))) options.IncludeHeadersAndFooters = !NoHeadersAndFooters.IsPresent;
        WriteObject(options);
    }
    private bool IsBound(string name) => MyInvocation.BoundParameters.ContainsKey(name);
}
