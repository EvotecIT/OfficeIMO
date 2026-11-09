using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Access;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Access;

/// <summary>Adds a table to a newly created native database model.</summary>
/// <example>
///   <summary>Adds a table to a newly created native database model.</summary>
///   <code>$table = $database | Add-OfficeAccessTable -Name Contacts -PassThru</code>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeAccessTable", DefaultParameterSetName = "Document")]
[OutputType(typeof(AccessTable))]
[Alias("AccessTable") ]
public sealed class AddOfficeAccessTableCommand : OfficeCmdlet {
    /// <summary>Live database document.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = "Document")]
    public AccessDocument? Document { get; set; }

    /// <summary>Table name.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Name { get; set; } = string.Empty;

    /// <summary>Content block whose AccessColumn and AccessRow commands target the new table.</summary>
    [Parameter(Position = 1)]
    public ScriptBlock? Content { get; set; }

    /// <summary>Return the created or changed object.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var table = ResolveDocument(Document).Tables.Add(Name);
        if (Content != null) { using (OfficeDocumentContext.Push(table)) Content.InvokeReturnAsIs(); }
        if (PassThru.IsPresent) WriteObject(table);
    }
}
