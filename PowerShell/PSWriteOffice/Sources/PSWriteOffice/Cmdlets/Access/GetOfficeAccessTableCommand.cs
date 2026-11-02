using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Access;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Access;

/// <summary>Lists native table definitions and schema without executing their content.</summary>
/// <example>
///   <summary>Lists native table definitions and schema without executing their content.</summary>
///   <code>$database | Get-OfficeAccessTable | Select-Object Name, RowCount, Columns</code>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeAccessTable", DefaultParameterSetName = "Document")]
[OutputType(typeof(AccessTable))]
public sealed class GetOfficeAccessTableCommand : PSCmdlet {
    /// <summary>Live database document.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public AccessDocument Document { get; set; } = null!;

    /// <summary>Optional table names.</summary>
    [Parameter(Position = 0)]
    public string[]? Name { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        foreach (var table in Document.Tables) {
            if (Name == null || Name.Contains(table.Name, StringComparer.OrdinalIgnoreCase)) WriteObject(table);
        }
    }
}
