using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Access;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Access;

/// <summary>Adds a table to a newly created native database model.</summary>
/// <example>
///   <summary>Adds a table to a newly created native database model.</summary>
///   <code>$table = $database | Add-OfficeAccessTable -Name Contacts -PassThru</code>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeAccessTable", DefaultParameterSetName = "Document")]
[OutputType(typeof(AccessTable))]
public sealed class AddOfficeAccessTableCommand : PSCmdlet {
    /// <summary>Live database document.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public AccessDocument Document { get; set; } = null!;

    /// <summary>Table name.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Name { get; set; } = string.Empty;

    /// <summary>Return the created or changed object.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var table = Document.Tables.Add(Name);
        if (PassThru.IsPresent) WriteObject(table);
    }
}
