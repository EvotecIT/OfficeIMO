using System.Management.Automation;
using OfficeIMO.OpenDocument;
using OfficeIMO.Word.OpenDocument;

namespace PSWriteOffice.Cmdlets.OpenDocument;

/// <summary>Creates Word/OpenDocument conversion settings.</summary>
/// <example>
///   <summary>Include Word images and headers during conversion.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeWordOpenDocumentOptions -IncludeImages -IncludeHeadersAndFooters
/// ConvertTo-OfficeOpenDocument -Path .\Report.docx -OutputPath .\Report.odt -WordOptions $options</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeWordOpenDocumentOptions")]
[OutputType(typeof(WordOpenDocumentConversionOptions))]
public sealed class NewOfficeWordOpenDocumentOptionsCommand : PSCmdlet {
    /// <summary>Whether conversion loss is reported or rejected.</summary>
    [Parameter] public OdfConversionLossPolicy? LossPolicy { get; set; }
    /// <summary>Copy supported inline images.</summary>
    [Parameter] public SwitchParameter IncludeImages { get; set; }
    /// <summary>Copy default headers and footers.</summary>
    [Parameter] public SwitchParameter IncludeHeadersAndFooters { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new WordOpenDocumentConversionOptions();
        if (LossPolicy.HasValue) options.LossPolicy = LossPolicy.Value;
        if (IsBound(nameof(IncludeImages))) options.IncludeImages = IncludeImages.IsPresent;
        if (IsBound(nameof(IncludeHeadersAndFooters))) options.IncludeHeadersAndFooters = IncludeHeadersAndFooters.IsPresent;
        WriteObject(options);
    }
    private bool IsBound(string name) => MyInvocation.BoundParameters.ContainsKey(name);
}
