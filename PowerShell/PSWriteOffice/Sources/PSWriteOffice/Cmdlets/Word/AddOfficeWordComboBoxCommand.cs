using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds a combo box content control to the current paragraph.</summary>
/// <example>
///   <summary>Add a combo box with a default value.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeWordParagraph { Add-OfficeWordComboBox -Items 'Red','Blue' -DefaultValue 'Blue' -Alias 'Color' }</code>
///   <para>Creates a combo box control with the selected default value.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordComboBox")]
[Alias("WordComboBox")]
[OutputType(typeof(WordComboBox))]
public sealed class AddOfficeWordComboBoxCommand : PSCmdlet
{
    /// <summary>Items to include in the combo box.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string[] Items { get; set; } = Array.Empty<string>();

    /// <summary>Optional default value (must match one of the items).</summary>
    [Parameter]
    public string? DefaultValue { get; set; }

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
        var control = paragraph.AddComboBox(values, Alias, Tag, DefaultValue);

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
