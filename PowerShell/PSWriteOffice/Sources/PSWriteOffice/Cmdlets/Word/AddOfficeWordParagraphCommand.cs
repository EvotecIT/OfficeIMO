using System.Management.Automation;
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
/// $paragraph | Add-OfficeWordText -Run @{ Text = 'Owner: ', 'Platform'; Bold = $true, $false }</code>
///   <para>Creates a paragraph on a live document and appends two differently formatted runs.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordParagraph", DefaultParameterSetName = ParameterSetText)]
[Alias("WordParagraph")]
[OutputType(typeof(WordParagraph))]
public sealed class AddOfficeWordParagraphCommand : PSCmdlet
{
    private const string ParameterSetText = "Text";
    private const string ParameterSetContent = "Content";

    /// <summary>Document or section that will receive the paragraph.</summary>
    [Parameter(ValueFromPipeline = true)]
    [Alias("Document", "Section")]
    public object? Target { get; set; }

    /// <summary>Optional initial paragraph text.</summary>
    [Parameter(Position = 0, ParameterSetName = ParameterSetText)]
    [Parameter(ParameterSetName = ParameterSetContent)]
    public string? Text { get; set; }

    /// <summary>Rich text runs. Each run can be created with TextRun/WordTextRun or provided as a hashtable/object.</summary>
    [Parameter(ParameterSetName = ParameterSetText)]
    [Alias("Runs")]
    public object[]? Run { get; set; }

    /// <summary>Nested DSL content (runs, lists, images).</summary>
    [Parameter(Position = 0, ParameterSetName = ParameterSetContent)]
    public ScriptBlock? Content { get; set; }

    /// <summary>Paragraph justification.</summary>
    [Parameter]
    public WordParagraphAlignment? Alignment { get; set; }

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
        if (!string.IsNullOrEmpty(Text) && Run is { Length: > 0 })
        {
            throw new PSArgumentException("Use either -Text or -Run, not both.");
        }

        WordDslContext? context = null;
        WordParagraph paragraph;
        WordDocument? document = null;
        var activeContext = Content != null ? WordDslContext.Current : null;
        var target = Target is PSObject psObject ? psObject.BaseObject : Target;
        if (target is WordDocument targetDocument)
        {
            if (activeContext != null && !ReferenceEquals(activeContext.Document, targetDocument))
            {
                throw new PSInvalidOperationException("The active Word DSL context belongs to a different document.");
            }

            document = targetDocument;
            context = activeContext;
            paragraph = targetDocument.AddParagraph(Run is { Length: > 0 } ? string.Empty : Text ?? string.Empty);
        }
        else if (target is WordSection targetSection)
        {
            if (Content != null)
            {
                if (activeContext == null)
                {
                    throw new PSInvalidOperationException("Nested paragraph content targeting a WordSection requires the section's active Word DSL context. Use -PassThru and target the returned paragraph for object-style composition.");
                }

                var sectionBelongsToActiveDocument = false;
                foreach (var section in activeContext.Document.Sections)
                {
                    if (ReferenceEquals(section, targetSection))
                    {
                        sectionBelongsToActiveDocument = true;
                        break;
                    }
                }

                if (!sectionBelongsToActiveDocument)
                {
                    throw new PSInvalidOperationException("The active Word DSL context belongs to a different document.");
                }

                context = activeContext;
            }

            paragraph = targetSection.AddParagraph(Run is { Length: > 0 } ? string.Empty : Text ?? string.Empty);
        }
        else if (target != null)
        {
            throw new PSArgumentException("-Target accepts a WordDocument or WordSection.", nameof(Target));
        }
        else
        {
            context = WordDslContext.Require(this);
            paragraph = context.RequireParagraphHost().AddParagraph(Run is { Length: > 0 } ? null : Text);
        }

        if (Alignment.HasValue)
        {
            paragraph.ParagraphAlignment = Alignment.Value;
        }

        if (Style.HasValue)
        {
            paragraph.Style = Style.Value;
        }

        if (!string.IsNullOrWhiteSpace(StyleId))
        {
            paragraph.SetStyleId(StyleId!);
        }

        if (Run is { Length: > 0 })
        {
            WordTextRunService.ApplyRuns(paragraph, Run);
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
            throw new PSInvalidOperationException("Nested paragraph content requires -Document or an active Word DSL context. Use -PassThru and target the returned paragraph for object-style composition.");
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
