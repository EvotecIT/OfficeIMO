using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Access;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Access;

/// <summary>Reads a native database without running queries, macros, or linked connections.</summary>
/// <example>
///   <summary>Reads a native database without running queries, macros, or linked connections.</summary>
///   <code>Get-OfficeAccess -Path ./contacts.accdb -TableName Contacts</code>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeAccess")]
[OutputType(typeof(AccessDocument))]
public sealed class GetOfficeAccessCommand : PSCmdlet {
    /// <summary>Literal database filename.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
    [Alias("FullName", "LiteralPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional user-table selection.</summary>
    [Parameter]
    public string[]? TableName { get; set; }

    /// <summary>Maximum source size in bytes.</summary>
    [Parameter]
    [ValidateRange(1L, int.MaxValue)]
    public long MaxInputBytes { get; set; } = 67108864;

    /// <summary>Maximum rows traversed by each native reader.</summary>
    [Parameter]
    [ValidateRange(1L, long.MaxValue)]
    public long MaxRows { get; set; } = 1000000;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        WriteObject(AccessDocument.Load(OfficeFileCommandUtilities.ResolvePath(this, Path), new AccessLoadOptions {
            AccessMode = DocumentAccessMode.ReadOnly, PersistenceMode = DocumentPersistenceMode.Explicit,
            TableNames = TableName, MaxInputBytes = MaxInputBytes, MaxRows = MaxRows
        }));
    }
}
