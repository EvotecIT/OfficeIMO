using System;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds a single list item.</summary>
/// <para>Can be called within <c>Add-OfficeWordList</c>/<c>WordList</c> or against an existing <see cref="WordList"/> from the pipeline; supports nesting via <c>-Level</c>.</para>
/// <example>
///   <summary>Add bullet text.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>WordList { Add-OfficeWordListItem -Text 'First task' }</code>
///   <para>Creates a bullet with the text “First task”.</para>
/// </example>
/// <example>
///   <summary>Append to an existing list.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Find-OfficeWordList -Document $doc -Text 'Initial review' | Add-OfficeWordListItem -Text 'Final approval'</code>
///   <para>Finds a list in an existing document and appends a new item.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordListItem", DefaultParameterSetName = "Text")]
[OutputType(typeof(WordParagraph))]
[Alias("WordListItem") ]
public sealed class AddOfficeWordListItemCommand : PSCmdlet
{
    /// <summary>Existing list to append to. When omitted, the current DSL list is used.</summary>
    [Parameter(ValueFromPipeline = true)]
    public WordList? List { get; set; }

    /// <summary>List item text.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Text"), AllowEmptyString]
    public string Text { get; set; } = string.Empty;

    /// <summary>Rich item content composed with WordText.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Content")]
    public ScriptBlock? Content { get; set; }

    /// <summary>One-based list level, from 1 to 9.</summary>
    [Parameter, ValidateRange(1, 9)]
    public int Level { get; set; } = 1;

    /// <summary>Emit the created <see cref="WordParagraph"/>.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var context = WordDslContext.Current;
        var list = List ?? context?.CurrentList ?? throw new InvalidOperationException("WordListItem must be used inside WordList or receive an existing WordList from the pipeline.");
        var anchor = context?.CurrentList == list
            ? context.ConsumeListAnchor(list)
            : null;
        var paragraph = list.AddItem(string.IsNullOrEmpty(Text) ? null : Text, Level - 1, anchor);

        if (anchor != null && !ReferenceEquals(paragraph, anchor) && string.IsNullOrWhiteSpace(anchor.Text))
        {
            anchor.Remove();
        }

        if (Content != null)
        {
            using (OfficeDocumentContext.Push(paragraph)) Content.InvokeReturnAsIs();
        }

        if (PassThru.IsPresent)
        {
            WriteObject(paragraph);
        }
    }
}
