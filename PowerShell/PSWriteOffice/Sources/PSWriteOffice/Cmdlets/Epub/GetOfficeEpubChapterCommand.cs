using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Epub;
using PSWriteOffice.Services;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Epub;

/// <summary>Projects chapters to searchable PowerShell objects.</summary>
/// <example>
///   <summary>Projects chapters to searchable PowerShell objects.</summary>
///   <code>$book | Get-OfficeEpubChapter | Select-Object Order, Title, Text</code>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeEpubChapter")]
[OutputType(typeof(EpubChapter))]
public sealed class GetOfficeEpubChapterCommand : OfficeDocumentCmdlet<EpubPublication> {
    /// <inheritdoc />
    protected override void ProcessRecord() {
        EpubPublication publication = GetDocument();
        WriteObject(publication.Read().Chapters, enumerateCollection: true);
    }
}
