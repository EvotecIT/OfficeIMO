using System;
using System.Management.Automation;
using OfficeIMO.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Creates a typed text or choice field for a PDF table cell.</summary>
/// <example>
///   <summary>Create a reviewer choice field for a typed PDF table cell.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$reviewer = New-OfficePdfTableCellField -Name Reviewer -Option 'Unassigned', 'Alice', 'Bob' -Value 'Unassigned'
/// $cell = New-OfficePdfTableCell -Text 'Reviewer' -FormField $reviewer</code>
///   <para>The choice field is positioned by the OfficeIMO PDF table renderer.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficePdfTableCellField", DefaultParameterSetName = ParameterSetText)]
[Alias("PdfTableCellField")]
[OutputType(typeof(PdfTableCellFormField))]
public sealed class NewOfficePdfTableCellFieldCommand : PSCmdlet
{
    private const string ParameterSetText = "Text";
    private const string ParameterSetChoice = "Choice";

    /// <summary>Unique AcroForm field name.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Name { get; set; } = string.Empty;

    /// <summary>Initial field value.</summary>
    [Parameter(Position = 1)]
    public string? Value { get; set; }

    /// <summary>Available values for a choice field.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetChoice)]
    public string[] Option { get; set; } = Array.Empty<string>();

    /// <summary>Rendered field width in PDF points.</summary>
    [Parameter]
    [ValidateRange(0.1D, double.MaxValue)]
    public double Width { get; set; } = 120D;

    /// <summary>Rendered field height in PDF points.</summary>
    [Parameter]
    [ValidateRange(0.1D, double.MaxValue)]
    public double Height { get; set; } = 18D;

    /// <summary>Field font size in PDF points.</summary>
    [Parameter]
    [ValidateRange(0.1D, double.MaxValue)]
    public double FontSize { get; set; } = 10D;

    /// <summary>Render a choice field as a list box instead of a combo box.</summary>
    [Parameter(ParameterSetName = ParameterSetChoice)]
    public SwitchParameter ListBox { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var field = ParameterSetName == ParameterSetChoice
            ? PdfTableCellFormField.ChoiceField(Name, Option, Value, Width, Height, FontSize, !ListBox.IsPresent)
            : PdfTableCellFormField.TextField(Name, Value, Width, Height, FontSize);
        WriteObject(field);
    }
}
