using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Epub;
using PSWriteOffice.Services;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
namespace PSWriteOffice.Cmdlets.Epub;

/// <summary>Adds chapter content, reading order and a navigation entry.</summary>
/// <example>
///   <summary>Adds chapter content, reading order and a navigation entry.</summary>
///   <code>$book | Add-OfficeEpubChapter -Title Opening -Text "Welcome to this handbook."</code>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeEpubChapter", DefaultParameterSetName = "Text")]
[OutputType(typeof(EpubManifestItem))]
public sealed class AddOfficeEpubChapterCommand : PSCmdlet {
    /// <summary>Live editable publication.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public EpubPublication Publication { get; set; } = null!;

    /// <summary>Chapter title.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Title { get; set; } = string.Empty;

    /// <summary>Plain text, escaped into XHTML paragraphs.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Text")]
    [AllowEmptyString]
    public string? Text { get; set; }

    /// <summary>Well-formed XHTML body fragment.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Xhtml")]
    public string? Xhtml { get; set; }

    /// <summary>Markdown rendered by the existing Markdown engine. Unsupported raw HTML remains subject to EPUB validation.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Markdown")]
    public string? Markdown { get; set; }

    /// <summary>Optional stable manifest identifier; omission chooses an unused chapter number.</summary>
    [Parameter]
    public string? Id { get; set; }

    /// <summary>Optional package-relative XHTML path.</summary>
    [Parameter]
    public string? ContainerPath { get; set; }

    /// <summary>Return the created or changed object.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        string id = Id ?? "chapter-" + (Publication.Spine.Count + 1);
        if (Id == null) {
            int suffix = Publication.Spine.Count + 1;
            while (Publication.Manifest.Any(item => item.Id == id)) id = "chapter-" + (++suffix);
        }
        string body = ParameterSetName == "Xhtml" ? Xhtml!
            : ParameterSetName == "Markdown" ? MarkdownDocument.Parse(Markdown!).ToHtml(new MarkdownHtmlExportOptions { Kind = MarkdownHtmlKind.Fragment })
            : "<p>" + System.Security.SecurityElement.Escape(Text ?? string.Empty)?.Replace("\r\n", "\n").Replace("\n", "</p><p>") + "</p>";
        var chapter = Publication.AddChapter(id, ContainerPath ?? "EPUB/text/" + id + ".xhtml", Title, body);
        if (PassThru.IsPresent) WriteObject(chapter);
    }
}
