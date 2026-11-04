using System.Management.Automation;
using OfficeIMO.AsciiDoc;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.AsciiDoc;

/// <summary>Creates native AsciiDocParseOptions settings for document operations.</summary>
/// <example><summary>Create default settings.</summary><code>$options = New-OfficeAsciiDocReadOptions</code></example>
[Cmdlet(VerbsCommon.New, "OfficeAsciiDocReadOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(OfficeIMO.AsciiDoc.AsciiDocParseOptions))]
public sealed class NewOfficeAsciiDocReadOptionsCommand : OfficeCmdlet {
    /// <summary>Profile override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public AsciiDocDocumentProfile? Profile { get; set; }
    /// <summary>MaximumInputLength override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public int? MaximumInputLength { get; set; }
    /// <summary>MaximumBlockCount override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public int? MaximumBlockCount { get; set; }
    /// <summary>MaximumInlineNestingDepth override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public int? MaximumInlineNestingDepth { get; set; }
    /// <summary>MaximumInlineNodeCount override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public int? MaximumInlineNodeCount { get; set; }
    /// <summary>MaximumTableCellCount override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public int? MaximumTableCellCount { get; set; }
    /// <summary>MaximumTableColumnCount override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public int? MaximumTableColumnCount { get; set; }
    /// <summary>MaximumTableSpan override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public int? MaximumTableSpan { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new OfficeIMO.AsciiDoc.AsciiDocParseOptions();
        if (Profile != null) options.Profile = Profile.Value;
        if (MaximumInputLength != null) options.MaxInputLength = MaximumInputLength.Value;
        if (MaximumBlockCount != null) options.MaxBlockCount = MaximumBlockCount.Value;
        if (MaximumInlineNestingDepth != null) options.MaxInlineNestingDepth = MaximumInlineNestingDepth.Value;
        if (MaximumInlineNodeCount != null) options.MaxInlineNodeCount = MaximumInlineNodeCount.Value;
        if (MaximumTableCellCount != null) options.MaxTableCellCount = MaximumTableCellCount.Value;
        if (MaximumTableColumnCount != null) options.MaxTableColumnCount = MaximumTableColumnCount.Value;
        if (MaximumTableSpan != null) options.MaxTableSpan = MaximumTableSpan.Value;
        WriteObject(options);
    }
}
