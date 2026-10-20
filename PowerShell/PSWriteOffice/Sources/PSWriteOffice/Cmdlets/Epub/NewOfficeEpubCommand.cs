using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Epub;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Epub;

/// <summary>Creates an EPUB publication with native navigation and metadata.</summary>
/// <example>
///   <summary>Creates an EPUB publication with native navigation and metadata.</summary>
///   <code>$book = New-OfficeEpub -Title Operations -Language en -Author Example</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeEpub")]
[OutputType(typeof(EpubPublication))]
public sealed class NewOfficeEpubCommand : PSCmdlet {
    /// <summary>Publication title.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Title { get; set; } = string.Empty;

    /// <summary>BCP 47 language.</summary>
    [Parameter]
    public string Language { get; set; } = "en";

    /// <summary>Optional stable publication identifier.</summary>
    [Parameter]
    public string? Identifier { get; set; }

    /// <summary>Primary author.</summary>
    [Parameter]
    public string? Author { get; set; }

    /// <summary>EPUB package generation.</summary>
    [Parameter]
    public EpubVersion Version { get; set; } = EpubVersion.Epub3;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var publication = EpubPublication.Create(Title, Language, Identifier, Version);
        if (Author != null) publication.Creator = Author;
        WriteObject(publication);
    }
}
