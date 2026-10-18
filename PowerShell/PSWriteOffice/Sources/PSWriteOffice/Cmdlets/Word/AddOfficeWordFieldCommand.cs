using System;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds a field to the current paragraph.</summary>
/// <para>Supports standard Word field codes such as Page, Date, or NumPages.</para>
/// <example>
///   <summary>Add a page number field.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeWordParagraph { Add-OfficeWordField -Type Page }</code>
///   <para>Inserts a PAGE field into the paragraph.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordField")]
[Alias("WordField")]
public sealed class AddOfficeWordFieldCommand : PSCmdlet
{
    /// <summary>Field type to insert.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public WordFieldType Type { get; set; }

    /// <summary>Optional field format switch.</summary>
    [Parameter]
    public WordFieldFormat? Format { get; set; }

    /// <summary>Custom format string (date/time fields).</summary>
    [Parameter]
    public string? CustomFormat { get; set; }

    /// <summary>Use advanced field representation.</summary>
    [Parameter]
    public SwitchParameter Advanced { get; set; }

    /// <summary>Additional field parameters.</summary>
    [Parameter]
    public string[]? Parameters { get; set; }

    /// <summary>Explicit paragraph to receive the field.</summary>
    [Parameter(ValueFromPipeline = true)]
    public WordParagraph? Paragraph { get; set; }

    /// <summary>Emit the paragraph after adding the field.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var paragraph = Paragraph;
        if (paragraph == null)
        {
            var context = WordDslContext.Require(this);
            paragraph = context.CurrentParagraph ?? context.RequireParagraphHost().AddParagraph();
        }

        List<string>? parameters = null;
        if (Parameters != null && Parameters.Length > 0)
        {
            parameters = new List<string>(Parameters);
        }

        paragraph.AddField(Type, Format, CustomFormat, Advanced.IsPresent, parameters);

        if (PassThru.IsPresent)
        {
            WriteObject(paragraph);
        }
    }
}
