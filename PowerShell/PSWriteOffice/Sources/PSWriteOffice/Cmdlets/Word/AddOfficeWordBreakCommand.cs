using System;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds a break to a Word paragraph.</summary>
/// <para>By default this creates a soft line break, equivalent to Shift+Enter in Word.</para>
/// <example>
///   <summary>Add a same-paragraph line break.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeWordParagraph { Add-OfficeWordText 'Line 1'; Add-OfficeWordBreak; Add-OfficeWordText 'Line 2' }</code>
///   <para>Writes both lines in the same paragraph separated by a soft break.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordBreak")]
[Alias("WordBreak")]
[OutputType(typeof(WordParagraph))]
public sealed class AddOfficeWordBreakCommand : PSCmdlet
{
    /// <summary>Target paragraph. When omitted, the current DSL paragraph is used or created.</summary>
    [Parameter(ValueFromPipeline = true)]
    public WordParagraph? Paragraph { get; set; }

    /// <summary>Optional break type, for example Page or Column.</summary>
    [Parameter]
    public WordBreakType? BreakType { get; set; }

    /// <summary>Number of breaks to add.</summary>
    [Parameter]
    [ValidateRange(1, 1000)]
    public int Count { get; set; } = 1;

    /// <summary>Emit the paragraph returned by the final break for additional native OfficeIMO chaining.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var paragraph = Paragraph ?? ResolveContextParagraph();
        var current = paragraph;

        for (var index = 0; index < Count; index++)
        {
            current = BreakType.HasValue
                ? current.AddBreak(BreakType.Value)
                : current.AddBreak();
        }

        if (PassThru.IsPresent)
        {
            WriteObject(current);
        }
    }

    private WordParagraph ResolveContextParagraph()
    {
        var context = WordDslContext.Require(this);
        return context.CurrentParagraph ?? context.RequireParagraphHost().AddParagraph();
    }
}
