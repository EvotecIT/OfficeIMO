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
[Cmdlet(VerbsCommon.Add, "OfficeWordField", DefaultParameterSetName = "Type")]
[Alias("WordField") ]
public sealed class AddOfficeWordFieldCommand : PSCmdlet
{
    /// <summary>Field type to insert.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Type")]
    public WordFieldType Type { get; set; }

    /// <summary>Mail-merge field name.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "MergeField"), ValidateNotNullOrEmpty]
    public string? MergeField { get; set; }
    /// <summary>Built-in or custom document property to display.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "DocumentProperty"), ValidateNotNullOrEmpty]
    public string? DocumentProperty { get; set; }
    /// <summary>Bookmark to reference.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Bookmark"), ValidateNotNullOrEmpty]
    public string? Bookmark { get; set; }

    /// <summary>Optional field format switch.</summary>
    [Parameter]
    public WordFieldFormat? Format { get; set; }

    /// <summary>Custom format string (date/time fields).</summary>
    [Parameter]
    public string? CustomFormat { get; set; }

    /// <summary>Use advanced field representation.</summary>
    [Parameter(ParameterSetName = "Type")]
    public SwitchParameter Advanced { get; set; }

    /// <summary>Additional field parameters.</summary>
    [Parameter(ParameterSetName = "Type")]
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

        if (ParameterSetName == "Type") paragraph.AddField(Type, Format, CustomFormat, Advanced.IsPresent, parameters);
        else {
            var fieldType = ParameterSetName == "MergeField" ? WordFieldType.MergeField :
                ParameterSetName == "DocumentProperty" ? WordFieldType.DocProperty : WordFieldType.Ref;
            var field = new WordFieldBuilder(fieldType).AddInstruction(MergeField ?? DocumentProperty ?? Bookmark!, quoted: true);
            paragraph.AddField(field, Format, CustomFormat);
        }

        if (PassThru.IsPresent)
        {
            WriteObject(paragraph);
        }
    }
}
