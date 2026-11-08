using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using System.Management.Automation;
using OfficeIMO.IWork;

namespace PSWriteOffice.Cmdlets.IWork;

/// <summary>Reads a supported Apple Pages, Numbers, or Keynote package without launching iWork.</summary>
/// <para>Reads supported iWork packages, including the qualified iWork '09 XML profile, into OfficeIMO's bounded source model.</para>
/// <example>
///   <summary>Inspect an iWork source before deciding how to convert it.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$source = Open-OfficeIWork -Path .\Quarterly.numbers
/// $source | Select-Object Kind, ContainerKind, BuildVersions
/// $source.ReadNumbers().Sheets | Select-Object Name</code>
///   <para>Returns OfficeIMO's bounded, loss-aware source model.</para>
/// </example>
[Cmdlet(VerbsCommon.Open, "OfficeIWork", DefaultParameterSetName = "Path")]
[OutputType(typeof(IWorkSourceDocument))]
public sealed class OpenOfficeIWorkCommand : PSCmdlet
{
    /// <summary>Path to a supported Pages, Numbers, or Keynote package or directory bundle, including the qualified iWork '09 XML profile.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Exact input path without wildcard expansion.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("PSPath", "FullName")]
    public string LiteralPath { get; set; } = string.Empty;


    /// <summary>Optional expected application kind; a mismatch is rejected.</summary>
    [Parameter]
    public IWorkDocumentKind? Kind { get; set; }

    /// <summary>Bounded OfficeIMO iWork read options.</summary>
    [Parameter]
    public IWorkLoadOptions? Options { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        foreach (var path in OfficePath.ResolveInput(SessionState, ParameterSetName == "LiteralPath" ? LiteralPath : Path, ParameterSetName == "LiteralPath", allowDirectory: true)) ReadPath(path);
    }
    private void ReadPath(string inputPath) {
        var path = inputPath;
        WriteObject(Kind.HasValue
            ? IWorkSourceDocument.Load(path, Kind.Value, Options)
            : IWorkSourceDocument.Load(path, Options));
    }
}
