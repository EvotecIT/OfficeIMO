using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Epub;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Epub;

/// <summary>Loads a publication for package-preserving inspection or editing.</summary>
/// <example>
///   <summary>Loads a publication for package-preserving inspection or editing.</summary>
///   <code>Get-Item ./book.epub | Get-OfficeEpub</code>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeEpub")]
[OutputType(typeof(EpubPublication))]
public sealed class GetOfficeEpubCommand : PSCmdlet {
    /// <summary>Literal EPUB path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
    [Alias("FullName", "FilePath", "LiteralPath")]
    public string Path { get; set; } = string.Empty;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        WriteObject(EpubPublication.Load(OfficeFileCommandUtilities.ResolvePath(this, Path)));
    }
}
