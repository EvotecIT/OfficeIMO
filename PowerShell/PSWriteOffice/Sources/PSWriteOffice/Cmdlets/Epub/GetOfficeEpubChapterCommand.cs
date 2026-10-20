using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Epub;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Epub;

/// <summary>Projects chapters to searchable PowerShell objects.</summary>
/// <example>
///   <summary>Projects chapters to searchable PowerShell objects.</summary>
///   <code>$book | Get-OfficeEpubChapter | Select-Object Order, Title, Text</code>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeEpubChapter")]
[OutputType(typeof(EpubChapter))]
public sealed class GetOfficeEpubChapterCommand : PSCmdlet {
    /// <summary>Live editable publication.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public EpubPublication Publication { get; set; } = null!;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        WriteObject(Publication.Read().Chapters, enumerateCollection: true);
    }
}
