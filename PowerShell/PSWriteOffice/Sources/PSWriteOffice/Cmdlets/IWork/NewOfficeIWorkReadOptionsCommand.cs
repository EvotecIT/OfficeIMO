using System.Management.Automation;
using OfficeIMO.IWork;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.IWork;

/// <summary>Creates native IWorkLoadOptions settings for document operations.</summary>
/// <example><summary>Create default settings.</summary><code>$options = New-OfficeIWorkReadOptions</code></example>
[Cmdlet(VerbsCommon.New, "OfficeIWorkReadOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(OfficeIMO.IWork.IWorkLoadOptions))]
public sealed class NewOfficeIWorkReadOptionsCommand : OfficeCmdlet {
    /// <summary>NoSourceRecords override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public SwitchParameter NoSourceRecords { get; set; }
    /// <summary>MaximumPackageBytes override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public long? MaximumPackageBytes { get; set; }
    /// <summary>MaximumEntryCount override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public int? MaximumEntryCount { get; set; }
    /// <summary>MaximumEntryBytes override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public long? MaximumEntryBytes { get; set; }
    /// <summary>MaximumTotalEntryBytes override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public long? MaximumTotalEntryBytes { get; set; }
    /// <summary>MaximumTableRows override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public int? MaximumTableRows { get; set; }
    /// <summary>MaximumTableColumns override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public int? MaximumTableColumns { get; set; }
    /// <summary>MaximumMaterializedCells override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public int? MaximumMaterializedCells { get; set; }
    /// <summary>MaximumProjectedImageBytes override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public long? MaximumProjectedImageBytes { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new OfficeIMO.IWork.IWorkLoadOptions();
        if (BoundSwitch(nameof(NoSourceRecords)) is bool valueNoSourceRecords) options.PreserveSourceRecords = !valueNoSourceRecords;
        if (MaximumPackageBytes != null) options.MaxPackageBytes = MaximumPackageBytes.Value;
        if (MaximumEntryCount != null) options.MaxEntryCount = MaximumEntryCount.Value;
        if (MaximumEntryBytes != null) options.MaxEntryBytes = MaximumEntryBytes.Value;
        if (MaximumTotalEntryBytes != null) options.MaxTotalEntryBytes = MaximumTotalEntryBytes.Value;
        if (MaximumTableRows != null) options.MaxTableRows = MaximumTableRows.Value;
        if (MaximumTableColumns != null) options.MaxTableColumns = MaximumTableColumns.Value;
        if (MaximumMaterializedCells != null) options.MaxMaterializedCells = MaximumMaterializedCells.Value;
        if (MaximumProjectedImageBytes != null) options.MaxProjectedImageBytes = MaximumProjectedImageBytes.Value;
        WriteObject(options);
    }
}
