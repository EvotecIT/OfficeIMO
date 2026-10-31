using System.IO;
using System.Management.Automation;
using OfficeIMO.Epub;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Epub;

/// <summary>Creates a live EPUB publication or saves a completed chapter/resource block.</summary>
/// <para>Without a block no file is written. A completed block saves atomically and requires Force to replace an existing destination.</para>
/// <example><summary>Create a handbook and save its completed content.</summary>
/// <code>New-OfficeEpub -Path ./Operations.epub -Title Operations -Author Example {
///     EpubChapter -Title Opening -Text 'Safety and care'
/// }</code></example>
/// <example><summary>Create a live publication and save it explicitly.</summary>
/// <code>$book = New-OfficeEpub -Path ./Operations.epub -Title Operations
/// $book | Add-OfficeEpubChapter -Title Opening -Text 'Safety and care'
/// Save-OfficeEpub -Document $book</code></example>
[Cmdlet(VerbsCommon.New, "OfficeEpub", SupportsShouldProcess = true)]
[OutputType(typeof(EpubPublication), typeof(FileInfo))]
public sealed class NewOfficeEpubCommand : OfficeLifecycleCmdlet<EpubPublication> {
    /// <summary>Optional associated destination; required with a content block.</summary>
    [Parameter(Position = 0), ValidateNotNullOrEmpty]
    public string? Path { get; set; }

    /// <summary>Publication title.</summary>
    [Parameter(Mandatory = true), ValidateNotNullOrEmpty]
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

    /// <summary>Allow a completed content block to replace an existing destination.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() => RunDocument(Path, isNew: true, force: Force.IsPresent,
        path => {
            var publication = EpubPublication.Create(Title, Language, Identifier, Version);
            publication.FilePath = path;
            if (Author != null) publication.Creator = Author;
            return publication;
        },
        (document, path, allowLoss) => {
            EpubWriteReport report = document.Save(path);
            if (!allowLoss) report.RequireNoLoss();
            return null;
        });
}
