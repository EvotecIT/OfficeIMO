using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds a simple AcroForm field to a generated PDF document.</summary>
/// <example>
///   <summary>Create a simple PDF form.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePdf -Path .\Examples\Documents\PdfForm.pdf {
///     Add-OfficePdfHeading -Text 'Access request'
///     Add-OfficePdfParagraph -Text 'Requester'
///     Add-OfficePdfFormField -Name 'Requester' -Type Text -Width 240
///     Add-OfficePdfParagraph -Text 'Priority'
///     Add-OfficePdfFormField -Name 'Priority' -Type Choice -Options 'Low','Normal','High' -Value 'Normal'
///     Add-OfficePdfFormField -Name 'Approved' -Type CheckBox
///   }</code>
///   <para>Adds text, choice, and checkbox form fields to a generated PDF.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePdfFormField", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfFormField")]
[OutputType(typeof(PdfDocument))]
public sealed class AddOfficePdfFormFieldCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>PDF document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Form field name.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Name { get; set; } = string.Empty;

    /// <summary>Field type to add.</summary>
    [Parameter]
    public OfficePdfFormFieldType Type { get; set; } = OfficePdfFormFieldType.Text;

    /// <summary>Initial text, selected choice, or selected radio value.</summary>
    [Parameter]
    public string? Value { get; set; }

    /// <summary>Initial selected values for multi-select choice fields.</summary>
    [Parameter]
    public string[] Values { get; set; } = System.Array.Empty<string>();

    /// <summary>Choice or radio options.</summary>
    [Parameter]
    public string[] Options { get; set; } = System.Array.Empty<string>();

    /// <summary>Initial check-box state.</summary>
    [Parameter]
    public SwitchParameter Checked { get; set; }

    /// <summary>Rendered field width in PDF points.</summary>
    [Parameter]
    public double Width { get; set; } = 180;

    /// <summary>Rendered field height in PDF points.</summary>
    [Parameter]
    public double Height { get; set; } = 22;

    /// <summary>Field alignment.</summary>
    [Parameter]
    public PdfAlign Align { get; set; } = PdfAlign.Left;

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Type is OfficePdfFormFieldType.Choice or OfficePdfFormFieldType.MultiSelectChoice or OfficePdfFormFieldType.RadioButton)
        {
            EnsureOptions();
        }

        var document = PdfCommandUtilities.ComposeContent(this, Document, ParameterSetName, ParameterSetDocument, content =>
        {
            switch (Type)
            {
                case OfficePdfFormFieldType.CheckBox:
                    content.CheckBox(Name, Checked.IsPresent, align: Align);
                    break;
                case OfficePdfFormFieldType.Choice:
                    content.ChoiceField(Name, Options, Value, Width, Height, Align);
                    break;
                case OfficePdfFormFieldType.MultiSelectChoice:
                    content.MultiSelectChoiceField(Name, Options, Values.Length > 0 ? Values : Value == null ? null : new[] { Value }, Width, Height, Align);
                    break;
                case OfficePdfFormFieldType.RadioButton:
                    content.RadioButtonGroup(Name, Options, Value, align: Align);
                    break;
                default:
                    content.TextField(Name, Width, Height, Value ?? string.Empty, Align);
                    break;
            }
        });

        if (PassThru.IsPresent && document != null)
        {
            WriteObject(document);
        }
    }

    private void EnsureOptions()
    {
        if (Options.Length == 0)
        {
            throw new PSArgumentException("Choice and radio button form fields require -Options.", nameof(Options));
        }
    }
}
