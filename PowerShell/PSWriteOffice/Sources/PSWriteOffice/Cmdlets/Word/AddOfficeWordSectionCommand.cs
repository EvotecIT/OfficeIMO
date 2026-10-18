using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds or reuses a section inside the current Word document.</summary>
/// <para>Provides the DSL entry point for section-level operations inside <c>New-OfficeWord</c>.</para>
/// <example>
///   <summary>Create a section with a paragraph.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeWord -Path .\doc.docx { Add-OfficeWordSection { Add-OfficeWordParagraph -Text 'Hello' } }</code>
///   <para>Creates a document and inserts a section that contains a single paragraph.</para>
/// </example>
/// <example>
///   <summary>Add a section to a live document.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$section = $document | Add-OfficeWordSection -BreakType NextPage -PassThru
/// $section | Add-OfficeWordParagraph -Text 'Appendix' -Style Heading1</code>
///   <para>Adds a section through the document pipeline and uses the returned section as the next explicit target.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordSection", DefaultParameterSetName = ParameterSetContext)]
[Alias("WordSection")]
[OutputType(typeof(WordSection))]
public sealed class AddOfficeWordSectionCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Document that will receive a new section.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public WordDocument? Document { get; set; }

    /// <summary>DSL scriptblock executed within the section scope.</summary>
    [Parameter(Position = 0, ParameterSetName = ParameterSetContext)]
    [Parameter(Position = 0, ParameterSetName = ParameterSetDocument)]
    public ScriptBlock? Content { get; set; }

    /// <summary>Optional section break type.</summary>
    [Parameter]
    public WordSectionBreakType? BreakType { get; set; }

    /// <summary>Emit the created <see cref="WordSection"/>.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Document == null)
        {
            var context = WordDslContext.Require(this);
            var section = context.AcquireSection(BreakType);

            using (context.Push(section))
            {
                Content?.InvokeReturnAsIs();
            }

            if (PassThru.IsPresent)
            {
                WriteObject(section);
            }
            return;
        }

        var current = Content != null ? WordDslContext.Current : null;
        if (current != null && !ReferenceEquals(current.Document, Document))
        {
            throw new PSInvalidOperationException("The active Word DSL context belongs to a different document.");
        }

        var createdSection = BreakType.HasValue
            ? Document.AddSection(BreakType.Value)
            : Document.AddSection();
        if (Content != null)
        {
            if (current != null)
            {
                using (current.Push(createdSection))
                {
                    Content.InvokeReturnAsIs();
                }
            }
            else
            {
                using var context = WordDslContext.Enter(Document);
                using (context.Push(createdSection))
                {
                    Content.InvokeReturnAsIs();
                }
            }
        }

        if (PassThru.IsPresent)
        {
            WriteObject(createdSection);
        }
    }
}
