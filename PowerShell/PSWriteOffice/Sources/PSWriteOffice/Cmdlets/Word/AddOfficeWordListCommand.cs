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
[Cmdlet(VerbsCommon.Add, "OfficeWordList")]
[Alias("WordList")]
public sealed class AddOfficeWordListCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Built-in list style or custom numbering scheme.</summary>
    [Parameter(Position = 1)]
    [Alias("Type")]
    public WordListStyle Style { get; set; } = WordListStyle.Bulleted;

    /// <summary>Scriptblock executed within the list scope.</summary>
    [Parameter(Position = 0)]
    public ScriptBlock? Content { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var context = WordDslContext.Require(this);
        var host = context.RequireParagraphHost();
        var anchor = host.AddParagraph();

        var list = anchor.AddList(Style);
        context.RegisterListAnchor(list, anchor);

        using (context.Push(list)) {
            Content?.InvokeReturnAsIs();
        }

        var leftoverAnchor = context.ConsumeListAnchor(list);
        if (leftoverAnchor != null && string.IsNullOrWhiteSpace(leftoverAnchor.Text)) {
            leftoverAnchor.Remove();
        }
        WritePassThru(list);
    }
}