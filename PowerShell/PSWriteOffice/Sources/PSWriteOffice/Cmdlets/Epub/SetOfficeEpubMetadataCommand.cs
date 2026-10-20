using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Epub;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Epub;

/// <summary>Updates the primary title, language or creator through the package model.</summary>
/// <example>
///   <summary>Updates the primary title, language or creator through the package model.</summary>
///   <code>$book | Set-OfficeEpubMetadata -Title "Operations handbook"</code>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeEpubMetadata")]
[OutputType(typeof(EpubPublication))]
public sealed class SetOfficeEpubMetadataCommand : PSCmdlet {
    /// <summary>Live editable publication.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public EpubPublication Publication { get; set; } = null!;

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
        if (Title != null) Publication.Title = Title;
        if (Language != null) Publication.Language = Language;
        if (Author != null) Publication.Creator = Author;
        if (PassThru.IsPresent) WriteObject(Publication);
    }
}
