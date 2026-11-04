using System.Management.Automation;
using OfficeIMO.Latex;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Latex;

/// <summary>Creates native LatexParseOptions settings for document operations.</summary>
/// <example><summary>Create default settings.</summary><code>$options = New-OfficeLatexReadOptions</code></example>
[Cmdlet(VerbsCommon.New, "OfficeLatexReadOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(OfficeIMO.Latex.LatexParseOptions))]
public sealed class NewOfficeLatexReadOptionsCommand : OfficeCmdlet {
    /// <summary>Profile override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public LatexDocumentProfile? Profile { get; set; }
    /// <summary>MacroExpansion override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public LatexMacroExpansion? MacroExpansion { get; set; }
    /// <summary>MaximumInputLength override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public int? MaximumInputLength { get; set; }
    /// <summary>MaximumInputBytes override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public long? MaximumInputBytes { get; set; }
    /// <summary>MaximumTokenCount override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public int? MaximumTokenCount { get; set; }
    /// <summary>MaximumNestingDepth override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public int? MaximumNestingDepth { get; set; }
    /// <summary>MaximumExpansionDepth override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public int? MaximumExpansionDepth { get; set; }
    /// <summary>MaximumExpansionLength override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public int? MaximumExpansionLength { get; set; }
    /// <summary>MaximumExpansionInputLength override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public int? MaximumExpansionInputLength { get; set; }
    /// <summary>MaximumExpansionTokenCount override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public int? MaximumExpansionTokenCount { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new OfficeIMO.Latex.LatexParseOptions();
        if (Profile != null) options.Profile = Profile.Value;
        if (MacroExpansion != null) options.MacroExpansion = MacroExpansion.Value;
        if (MaximumInputLength != null) options.MaxInputLength = MaximumInputLength.Value;
        if (MaximumInputBytes != null) options.MaxInputBytes = MaximumInputBytes.Value;
        if (MaximumTokenCount != null) options.MaxTokenCount = MaximumTokenCount.Value;
        if (MaximumNestingDepth != null) options.MaxNestingDepth = MaximumNestingDepth.Value;
        if (MaximumExpansionDepth != null) options.MaxExpansionDepth = MaximumExpansionDepth.Value;
        if (MaximumExpansionLength != null) options.MaxExpansionLength = MaximumExpansionLength.Value;
        if (MaximumExpansionInputLength != null) options.MaxExpansionInputLength = MaximumExpansionInputLength.Value;
        if (MaximumExpansionTokenCount != null) options.MaxExpansionTokenCount = MaximumExpansionTokenCount.Value;
        WriteObject(options);
    }
}
