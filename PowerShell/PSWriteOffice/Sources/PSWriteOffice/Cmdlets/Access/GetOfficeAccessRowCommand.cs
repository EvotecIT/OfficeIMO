using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Access;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Access;

/// <summary>Streams native or modeled table rows as PowerShell objects.</summary>
/// <example>
///   <summary>Streams native or modeled table rows as PowerShell objects.</summary>
///   <code>$database | Get-OfficeAccessRow -TableName Contacts | Export-Csv ./contacts.csv -NoTypeInformation</code>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeAccessRow")]
[OutputType(typeof(PSObject))]
public sealed class GetOfficeAccessRowCommand : PSCmdlet {
    /// <summary>Live database document.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public AccessDocument Document { get; set; } = null!;

    /// <summary>Table name.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string TableName { get; set; } = string.Empty;

    /// <summary>Maximum rows emitted; a partial result produces a warning.</summary>
    [Parameter]
    [ValidateRange(1L, long.MaxValue)]
    public long MaxRows { get; set; } = 1000000;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        using var reader = Document.Tables[TableName].OpenDataReader();
        long count = 0;
        while (reader.Read()) {
            if (count++ == MaxRows) {
                WriteWarning("The selected row limit was reached; the output is partial.");
                break;
            }
            var row = new PSObject();
            for (int index = 0; index < reader.FieldCount; index++)
                row.Properties.Add(new PSNoteProperty(reader.GetName(index), reader.IsDBNull(index) ? null : reader.GetValue(index)));
            WriteObject(row);
        }
    }
}
