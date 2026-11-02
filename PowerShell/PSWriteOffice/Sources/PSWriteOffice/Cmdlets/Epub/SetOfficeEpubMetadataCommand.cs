using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Epub;
using PSWriteOffice.Services;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Epub;

/// <summary>Updates the primary title, language or creator through the package model.</summary>
/// <example>
///   <summary>Updates the primary title, language or creator through the package model.</summary>
///   <code>$book | Set-OfficeEpubMetadata -Title "Operations handbook"</code>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeEpubMetadata", DefaultParameterSetName = "Default")]
[OutputType(typeof(EpubPublication))]
public sealed class SetOfficeEpubMetadataCommand : OfficeDocumentCmdlet<EpubPublication> {
    /// <summary>New primary title.</summary>
    [Parameter]
    public string? Title { get; set; }

    /// <summary>New BCP 47 language.</summary>
    [Parameter]
    public string? Language { get; set; }

    /// <summary>New primary author.</summary>
    [Parameter]
    public string? Author { get; set; }

    /// <summary>Return the created or changed object.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        EpubPublication publication = GetDocument();
        if (Title != null) publication.Title = Title;
        if (Language != null) publication.Language = Language;
        if (Author != null) publication.Creator = Author;
        if (PassThru.IsPresent) WriteObject(publication);
    }
}
