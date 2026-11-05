using PSWriteOffice.Foundation;
using OfficeIMO;
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
///     Add-OfficePdfFormField -Name 'Priority' -Type Choice -Option 'Low','Normal','High' -Value 'Normal'
///     Add-OfficePdfFormField -Name 'Approved' -Type CheckBox
///   }</code>
///   <para>Adds text, choice, and checkbox form fields to a generated PDF.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePdfFormField", DefaultParameterSetName = ParameterSetContext)]
[OutputType(typeof(PdfDocument))]
[Alias("PdfFormField") ]
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

    /// <summary>Initial value; MultiSelectChoice accepts multiple selected values, other fields accept at most one.</summary>
    [Parameter]
    public string[]? Value { get; set; }

    /// <summary>Export name of the checked state for a check box.</summary>
    [Parameter]
    public string CheckedValue { get; set; } = "Yes";

    /// <summary>Choice or radio options.</summary>
    [Parameter]
    public string[] Option { get; set; } = System.Array.Empty<string>();

    /// <summary>Render scalar choices as a list box rather than a combo box.</summary>
    [Parameter]
    public SwitchParameter ListBox { get; set; }

    /// <summary>Initial check-box state.</summary>
    [Parameter]
    public SwitchParameter Checked { get; set; }

    /// <summary>Rendered field width in PDF points.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength Width { get; set; } = OfficeLength.Points(180);

    /// <summary>Rendered field height in PDF points.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength Height { get; set; } = OfficeLength.Points(22);

    /// <summary>Native field text size or check-box size.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? Size { get; set; }

    /// <summary>Field alignment.</summary>
    [Parameter]
    public OfficeAlignment Alignment { get; set; } = OfficeAlignment.Left;

    /// <summary>Emit the updated document.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Type != OfficePdfFormFieldType.MultiSelectChoice && Value is { Length: > 1 })
            throw new PSArgumentException("Only MultiSelectChoice accepts more than one initial Value.");
        string? scalarValue = Value is { Length: > 0 } ? Value[0] : null;
        if (Type is OfficePdfFormFieldType.Choice or OfficePdfFormFieldType.MultiSelectChoice or OfficePdfFormFieldType.RadioButton)
        {
            EnsureOptions();
        }

        if (ListBox && Type != OfficePdfFormFieldType.Choice) throw new PSArgumentException("ListBox applies to scalar Choice fields.");
        var cell = Document == null ? OfficeDocumentContext.Find<PdfCellContent>() : null;
        if (cell != null) {
            switch (Type) {
                case OfficePdfFormFieldType.CheckBox:
                    cell.CheckBoxes.Add(new PdfTableCellCheckBox(Name, Checked, Size, CheckedValue));
                    break;
                case OfficePdfFormFieldType.Choice:
                    cell.Fields.Add(PdfTableCellFormField.ChoiceField(Name, Option, scalarValue, MyInvocation.BoundParameters.ContainsKey(nameof(Width)) ? Width : null, MyInvocation.BoundParameters.ContainsKey(nameof(Height)) ? Height : null, Size, !ListBox));
                    break;
                case OfficePdfFormFieldType.Text:
                    cell.Fields.Add(PdfTableCellFormField.TextField(Name, scalarValue, MyInvocation.BoundParameters.ContainsKey(nameof(Width)) ? Width : null, MyInvocation.BoundParameters.ContainsKey(nameof(Height)) ? Height : null, Size));
                    break;
                default: throw new PSNotSupportedException("Native table-cell fields support text, scalar choice and check boxes; use a flowing field for this kind.");
            }
            return;
        }

        var document = PdfCommandUtilities.ComposeContent(this, Document, ParameterSetName, ParameterSetDocument, content =>
        {
            switch (Type)
            {
                case OfficePdfFormFieldType.CheckBox:
                    content.CheckBox(Name, Checked.IsPresent, Size, align: Alignment, checkedValueName: CheckedValue);
                    break;
                case OfficePdfFormFieldType.Choice:
                    content.ChoiceField(Name, Option, scalarValue, Width, Height, Alignment, fontSize: Size, isComboBox: !ListBox);
                    break;
                case OfficePdfFormFieldType.MultiSelectChoice:
                    content.MultiSelectChoiceField(Name, Option, Value, Width, Height, Alignment, fontSize: Size);
                    break;
                case OfficePdfFormFieldType.RadioButton:
                    content.RadioButtonGroup(Name, Option, scalarValue, align: Alignment);
                    break;
                default:
                    content.TextField(Name, Width, Height, scalarValue ?? string.Empty, Alignment, fontSize: Size);
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
        if (Option.Length == 0)
        {
            throw new PSArgumentException("Choice and radio button form fields require -Option.", nameof(Option));
        }
    }
}
