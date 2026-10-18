using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds a dropdown list content control to the current paragraph.</summary>
/// <example>
///   <summary>Add a dropdown list.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeWordParagraph { Add-OfficeWordDropDownList -Items 'Low','Medium','High' -Alias 'Priority' }</code>
///   <para>Creates a dropdown list control with the supplied items.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordDropDownList")]
[Alias("WordDropDownList")]
[OutputType(typeof(WordDropDownList))]
public sealed class AddOfficeWordDropDownListCommand : PSCmdlet
{
    /// <summary>Items to include in the dropdown list.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string[] Items { get; set; } = Array.Empty<string>();

    /// <summary>Optional alias for the control.</summary>
    [Parameter]
    public string? Alias { get; set; }

    /// <summary>Optional tag for the control.</summary>
    [Parameter]
    public string? Tag { get; set; }

    /// <summary>Explicit paragraph to receive the control.</summary>
    [Parameter(ValueFromPipeline = true)]
    public WordParagraph? Paragraph { get; set; }

    /// <summary>Emit the created control.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var values = NormalizeItems(Items);
        var paragraph = ResolveParagraph();
        var control = paragraph.AddDropDownList(values, Alias, Tag);

        if (PassThru.IsPresent)
        {
            WriteObject(control);
        }
    }

    private static List<string> NormalizeItems(string[]? items)
    {
        if (items == null || items.Length == 0)
        {
            throw new PSArgumentException("Items cannot be empty.", nameof(Items));
        }

        var list = items
            .Select(item => item?.Trim())
            .Where(item => !string.IsNullOrWhiteSpace(item))
            .Cast<string>()
            .ToList();

        if (list.Count == 0)
        {
            throw new PSArgumentException("Items cannot be empty.", nameof(Items));
        }

        return list;
    }

    private WordParagraph ResolveParagraph()
    {
        if (Paragraph != null)
        {
            return Paragraph;
        }

        var context = WordDslContext.Require(this);
        return context.CurrentParagraph ?? context.RequireParagraphHost().AddParagraph();
    }
}
