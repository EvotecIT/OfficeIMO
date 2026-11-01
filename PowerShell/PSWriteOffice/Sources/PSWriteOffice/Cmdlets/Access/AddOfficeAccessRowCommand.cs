using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Access;
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
public sealed class AddOfficeAccessRowCommand : PSCmdlet {
    /// <summary>Destination table in a newly created model.</summary>
    [Parameter(Mandatory = true)]
    public AccessTable Table { get; set; } = null!;

    /// <summary>One row; omitted properties remain absent, and explicit null remains null.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
    public PSObject[] InputObject { get; set; } = Array.Empty<PSObject>();

    /// <summary>Return the created or changed object.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        foreach (var input in InputObject) {
            var values = AccessPowerShellInput.ToRow(Table, input);
            Table.AppendRow(values);
        }
        if (PassThru.IsPresent) WriteObject(Table);
    }
}
