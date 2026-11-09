using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Starts a list inside the current section or paragraph anchor.</summary>
/// <para>Creates a temporary anchor paragraph, spawns an OfficeIMO list, and runs child <c>WordListItem</c> commands.</para>
/// <example>
///   <summary>Numbered checklist.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeWordList -Style 'Numbered' { Add-OfficeWordListItem -Text 'Plan'; Add-OfficeWordListItem -Text 'Execute' }</code>
///   <para>Creates a numbered list with two steps.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordList", DefaultParameterSetName = "Content")]
[OutputType(typeof(WordList))]
[Alias("WordList") ]
public sealed class AddOfficeWordListCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Built-in list style or custom numbering scheme.</summary>
    [Parameter]
    public WordListStyle Style { get; set; } = WordListStyle.Bulleted;

    /// <summary>Scriptblock executed within the list scope.</summary>
    [Parameter(Position = 0, ParameterSetName = "Content")]
    public ScriptBlock? Content { get; set; }

    /// <summary>Plain list items.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Items")]
    public string[] Items { get; set; } = Array.Empty<string>();
    /// <summary>Use the native numbered list style.</summary>
    [Parameter]
    public SwitchParameter Numbered { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var context = WordDslContext.Require(this);
        var host = context.RequireParagraphHost();
        var anchor = host.AddParagraph();

        var list = anchor.AddList(Numbered ? WordListStyle.Numbered : Style);
        context.RegisterListAnchor(list, anchor);

        using (context.Push(list)) {
            if (Content != null) Content.InvokeReturnAsIs();
            else foreach (string text in Items) {
                var placeholder = context.ConsumeListAnchor(list);
                var item = list.AddItem(text, 0, placeholder);
                if (placeholder != null && string.IsNullOrWhiteSpace(placeholder.Text)) placeholder.Remove();
            }
        }

        var leftoverAnchor = context.ConsumeListAnchor(list);
        if (leftoverAnchor != null && string.IsNullOrWhiteSpace(leftoverAnchor.Text)) {
            leftoverAnchor.Remove();
        }
        WritePassThru(list);
    }
}