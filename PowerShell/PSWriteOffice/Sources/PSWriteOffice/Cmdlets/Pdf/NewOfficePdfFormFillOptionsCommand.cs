using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Creates native XFDF field/value limits and appearance behavior for Import-OfficePdfXfdf.</summary>
/// <example><summary>Bound an imported form data set.</summary>
/// <code>$options = New-OfficePdfFormFillOptions -MaxFields 100 -MaxValueCharacters 10000
/// Import-OfficePdfXfdf -Path ./Form.pdf -XfdfPath ./Values.xfdf -OutputPath ./Filled.pdf -Options $options</code></example>
[Cmdlet(VerbsCommon.New, "OfficePdfFormFillOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(PdfFormFillerOptions))]
public sealed class NewOfficePdfFormFillOptionsCommand : OfficeCmdlet {
    /// <summary>Maximum fields admitted from one XFDF document.</summary>
    [Parameter(ParameterSetName = "Default"), ValidateRange(1, int.MaxValue)]
    public int? MaxFields { get; set; }
    /// <summary>Maximum aggregate XFDF value characters.</summary>
    [Parameter, ValidateRange(1, int.MaxValue)]
    public int? MaxValueCharacters { get; set; }
    /// <summary>Keep the native NeedAppearances flag set for legacy viewers.</summary>
    [Parameter]
    public SwitchParameter KeepNeedAppearances { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new PdfFormFillerOptions { KeepNeedAppearances = KeepNeedAppearances.IsPresent };
        if (MaxFields.HasValue) options.MaxXfdfFields = MaxFields.Value;
        if (MaxValueCharacters.HasValue) options.MaxXfdfValueCharacters = MaxValueCharacters.Value;
        WriteObject(options);
    }
}
