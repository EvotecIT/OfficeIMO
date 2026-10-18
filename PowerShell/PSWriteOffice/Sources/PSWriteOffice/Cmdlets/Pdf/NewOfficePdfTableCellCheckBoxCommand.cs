using System.Management.Automation;
using OfficeIMO.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Creates a typed check box for a PDF table cell.</summary>
/// <example>
///   <summary>Create a checked table-cell field.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$approved = New-OfficePdfTableCellCheckBox -Name Approved -Checked
/// $cell = New-OfficePdfTableCell -Text 'Approved' -CheckBox $approved</code>
///   <para>The check box remains an AcroForm field positioned by the OfficeIMO table renderer.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficePdfTableCellCheckBox")]
[Alias("PdfTableCellCheckBox")]
[OutputType(typeof(PdfTableCellCheckBox))]
public sealed class NewOfficePdfTableCellCheckBoxCommand : PSCmdlet
{
    /// <summary>Unique AcroForm field name.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Name { get; set; } = string.Empty;

    /// <summary>Create the check box in its checked state.</summary>
    [Parameter]
    public SwitchParameter Checked { get; set; }

    /// <summary>Visual square size in PDF points.</summary>
    [Parameter]
    [ValidateRange(0.1D, double.MaxValue)]
    public double Size { get; set; } = 12D;

    /// <summary>PDF appearance-state name written when checked.</summary>
    [Parameter]
    public string CheckedValueName { get; set; } = "Yes";

    /// <inheritdoc />
    protected override void ProcessRecord()
        => WriteObject(new PdfTableCellCheckBox(Name, Checked.IsPresent, Size, CheckedValueName));
}
