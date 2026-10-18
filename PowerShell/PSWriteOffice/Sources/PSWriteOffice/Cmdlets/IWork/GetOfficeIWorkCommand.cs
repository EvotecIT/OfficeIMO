using System.Management.Automation;
using OfficeIMO.IWork;

namespace PSWriteOffice.Cmdlets.IWork;

/// <summary>Reads a modern Apple Pages, Numbers, or Keynote package without launching iWork.</summary>
/// <example>
///   <summary>Inspect an iWork source before deciding how to convert it.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$source = Get-OfficeIWork -Path .\Quarterly.numbers
/// $source | Select-Object Kind, ContainerKind, BuildVersions
/// $source.ReadNumbers().Sheets | Select-Object Name</code>
///   <para>Returns OfficeIMO's bounded, loss-aware source model.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeIWork")]
[OutputType(typeof(IWorkSourceDocument))]
public sealed class GetOfficeIWorkCommand : PSCmdlet
{
    /// <summary>Path to a modern Pages, Numbers, or Keynote package or directory bundle.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional expected application kind; a mismatch is rejected.</summary>
    [Parameter]
    public IWorkDocumentKind? Kind { get; set; }

    /// <summary>Bounded OfficeIMO iWork read options.</summary>
    [Parameter]
    public IWorkReadOptions? Options { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var path = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        WriteObject(Kind.HasValue
            ? IWorkSourceDocument.Open(path, Kind.Value, Options)
            : IWorkSourceDocument.Open(path, Options));
    }
}
