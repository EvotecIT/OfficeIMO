using PSWriteOffice.Foundation;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Access;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Access;

/// <summary>Reads a native database without running queries, macros, or linked connections.</summary>
/// <example>
///   <summary>Reads a native database without running queries, macros, or linked connections.</summary>
///   <code>Open-OfficeAccess -Path ./contacts.accdb -Table Contacts</code>
/// </example>
[Cmdlet(VerbsCommon.Open, "OfficeAccess", DefaultParameterSetName = "Path")]
[OutputType(typeof(AccessDocument))]
public sealed class OpenOfficeAccessCommand : PSCmdlet {
    /// <summary>Literal database filename.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Exact input path without wildcard expansion.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("PSPath", "FullName")]
    public string LiteralPath { get; set; } = string.Empty;


    /// <summary>Optional user-table selection.</summary>
    [Parameter]
    public string[]? Table { get; set; }

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
        foreach (var path in OfficePath.ResolveInput(SessionState, ParameterSetName == "LiteralPath" ? LiteralPath : Path, ParameterSetName == "LiteralPath", allowDirectory: false)) ReadPath(path);
    }
    private void ReadPath(string inputPath) {
        WriteObject(AccessDocument.Load(inputPath, new AccessLoadOptions {
            AccessMode = DocumentAccessMode.ReadOnly, PersistenceMode = DocumentPersistenceMode.Explicit,
            TableNames = Table, MaxInputBytes = MaxInputBytes, MaxRows = MaxRows
        }));
    }
}
