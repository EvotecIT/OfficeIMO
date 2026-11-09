using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Access;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Access;
namespace PSWriteOffice.Cmdlets.Access;

/// <summary>Appends PowerShell objects or dictionaries using the declared field types.</summary>
/// <example>
///   <summary>Appends PowerShell objects or dictionaries using the declared field types.</summary>
///   <code>$rows | Add-OfficeAccessRow -Table $table</code>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeAccessRow", DefaultParameterSetName = "Table")]
[OutputType(typeof(AccessTable))]
[Alias("AccessRow") ]
public sealed class AddOfficeAccessRowCommand : OfficeCmdlet {
    /// <summary>Destination table in a newly created model.</summary>
    [Parameter(ParameterSetName = "Table")]
    public AccessTable? Table { get; set; }

    /// <summary>One row; omitted properties remain absent, and explicit null remains null.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
    public PSObject[] InputObject { get; set; } = Array.Empty<PSObject>();

    /// <summary>Return the created or changed object.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var table = ResolveDocument(Table);
        foreach (var input in InputObject) {
            var values = AccessPowerShellInput.ToRow(table, input);
            table.AppendRow(values);
        }
        if (PassThru.IsPresent) WriteObject(table);
    }
}
