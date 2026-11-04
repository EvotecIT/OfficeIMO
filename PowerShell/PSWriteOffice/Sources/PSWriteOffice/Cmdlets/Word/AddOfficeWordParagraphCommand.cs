using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds a paragraph to the current section/header/footer context.</summary>
/// <para>Acts as the primary DSL container for inline content such as text runs, bold segments, and images.</para>
/// <example>
///   <summary>Write a formatted sentence.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeWordParagraph { Add-OfficeWordText -Text 'Hello '; Add-OfficeWordText -Text 'World' -Bold }</code>
///   <para>Outputs “Hello World” with the second word bolded.</para>
/// </example>
/// <example>
///   <summary>Apply a paragraph style by id.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>WordParagraph -Text 'Executive summary' -StyleId 'ReportHeading'</code>
///   <para>Applies a paragraph style id, including custom styles already present in a template document.</para>
/// </example>
/// <example>
///   <summary>Add mixed-format text through explicit objects.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$paragraph = $document | Add-OfficeWordParagraph -PassThru
/// $paragraph | Add-OfficeWordText 'Owner: ' -Bold
/// $paragraph | Add-OfficeWordText 'Platform'</code>
///   <para>Creates a paragraph on a live document and appends two differently formatted runs.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordParagraph", DefaultParameterSetName = ParameterSetText)]
[OutputType(typeof(WordParagraph))]
[Alias("WordParagraph") ]
public sealed class AddOfficeWordParagraphCommand : PSCmdlet
{
    private const string ParameterSetText = "Text";
    private const string ParameterSetContent = "Content";

    /// <summary>Live document to receive the paragraph; omitted in a content block.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetText)]
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetContent)]
    public WordDocument? Document { get; set; }

    /// <summary>Section to receive the paragraph.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "SectionText")]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "SectionContent")]
    public WordSection? Section { get; set; }

    /// <summary>Initial paragraph text; omission creates an empty paragraph.</summary>
    [Parameter(Position = 0, ParameterSetName = ParameterSetText)]
    [Parameter(Position = 0, ParameterSetName = "SectionText")]
    [AllowEmptyString]
    public string? Text { get; set; }

    /// <summary>Nested WordText and inline content.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetContent)]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "SectionContent")]
    public ScriptBlock? Content { get; set; }

    /// <summary>Paragraph justification.</summary>
    [Parameter]
    public OfficeAlignment? Alignment { get; set; }

    /// <summary>Paragraph style.</summary>
    [Parameter]
    public WordParagraphStyles? Style { get; set; }

    /// <summary>Paragraph style id, including custom style ids from a template document.</summary>
    [Parameter]
    public string? StyleId { get; set; }

    /// <summary>Emit the <see cref="WordParagraph"/> for further use.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        WordDslContext? context = null;
        WordParagraph paragraph;
        WordDocument? document = null;
        var activeContext = Content != null ? WordDslContext.Current : null;
        object? target = (object?)Document ?? Section;
        if (target is WordDocument targetDocument)
        {
            if (activeContext != null && !ReferenceEquals(activeContext.Document, targetDocument))
            {
                throw new PSInvalidOperationException("The active Word DSL context belongs to a different document.");
            }

            document = targetDocument;
            context = activeContext;
            paragraph = targetDocument.AddParagraph(Text ?? string.Empty);
        }
        else if (target is WordSection targetSection)
        {
            if (Content != null)
            {
                var sectionBelongsToActiveDocument = false;
                foreach (var section in activeContext?.Document.Sections ?? new System.Collections.Generic.List<WordSection>())
                {
                    if (ReferenceEquals(section, targetSection))
                    {
                        sectionBelongsToActiveDocument = true;
                        break;
                    }
                }

                if (activeContext != null && !sectionBelongsToActiveDocument)
                {
                    throw new PSInvalidOperationException("The active Word DSL context belongs to a different document.");
                }

                context = activeContext;
            }

            paragraph = targetSection.AddParagraph(Text ?? string.Empty);
        }
        else if (PSWriteOffice.Foundation.OfficeDocumentContext.Find<WordTableCell>() is WordTableCell cell)
        {
            paragraph = cell.AddParagraph(Text ?? string.Empty);
        }
        else
        {
            context = WordDslContext.Require(this);
            paragraph = context.RequireParagraphHost().AddParagraph(Text);
        }

        if (Alignment.HasValue)
        {
            paragraph.ParagraphAlignment = Alignment.Value switch {
                OfficeAlignment.Left or OfficeAlignment.Start => WordParagraphAlignment.Left,
                OfficeAlignment.Right or OfficeAlignment.End => WordParagraphAlignment.Right,
                OfficeAlignment.Center => WordParagraphAlignment.Center,
                OfficeAlignment.Justify => WordParagraphAlignment.Both,
                _ => throw new PSArgumentException("Unsupported paragraph alignment.", nameof(Alignment))
            };
        }

        if (Style.HasValue)
        {
            paragraph.Style = Style.Value;
        }

        if (!string.IsNullOrWhiteSpace(StyleId))
        {
            paragraph.SetStyleId(StyleId!);
        }

        if (Content != null)
        {
            InvokeContent(document, context, paragraph);
        }

        if (PassThru.IsPresent)
        {
            WriteObject(paragraph);
        }
    }

    private void InvokeContent(WordDocument? document, WordDslContext? context, WordParagraph paragraph)
    {
        if (context != null)
        {
            using (context.Push(paragraph))
            {
                Content!.InvokeReturnAsIs();
            }
            return;
        }

        if (document == null)
        {
            using (PSWriteOffice.Foundation.OfficeDocumentContext.Push(paragraph)) Content!.InvokeReturnAsIs();
            return;
        }

        var current = WordDslContext.Current;
        if (current != null && !ReferenceEquals(current.Document, document))
        {
            throw new PSInvalidOperationException("The active Word DSL context belongs to a different document.");
        }

        if (current != null)
        {
            using (current.Push(paragraph))
            {
                Content!.InvokeReturnAsIs();
            }
            return;
        }

        using var ownedContext = WordDslContext.Enter(document);
        using (ownedContext.Push(paragraph))
        {
            Content!.InvokeReturnAsIs();
        }
    }
}
