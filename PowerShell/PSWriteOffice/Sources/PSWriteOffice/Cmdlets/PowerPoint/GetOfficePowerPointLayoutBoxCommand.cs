using System;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Computes native content, column or row boxes for manual placement.</summary>
/// <example><summary>Get a content box with a physical margin.</summary>
/// <code>$box = Get-OfficePowerPointLayoutBox -Document $deck -Margin 1.5cm</code></example>
/// <example><summary>Place text in equal columns.</summary>
/// <code>$columns = Get-OfficePowerPointLayoutBox -Document $deck -Columns 2 -Margin 1.5cm -Gutter 1cm
/// PowerPointSlide -Document $deck -Title 'Status' {
///     PowerPointTextBox 'Healthy' -X $columns[0].Left -Y $columns[0].Top -Width $columns[0].Width -Height 40
/// }</code></example>
[Cmdlet(VerbsCommon.Get, "OfficePowerPointLayoutBox", DefaultParameterSetName = "Content")]
[OutputType(typeof(PowerPointLayoutBox))]
public sealed class GetOfficePowerPointLayoutBoxCommand : OfficeDocumentCmdlet<PowerPointDocument> {
    /// <summary>Gets or sets the number of columns.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Columns"), ValidateRange(1, int.MaxValue)]
    public int Columns { get; set; }
    /// <summary>Gets or sets the number of rows.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Rows"), ValidateRange(1, int.MaxValue)]
    public int Rows { get; set; }
    /// <summary>Gets or sets the outer margin; numbers mean points.</summary>
    [Parameter(ParameterSetName = "Content"), Parameter(ParameterSetName = "Columns"),
        Parameter(ParameterSetName = "Rows"), OfficeLengthTransformation]
    public OfficeLength Margin { get; set; }
    /// <summary>Gets or sets the space between rows or columns; numbers mean points.</summary>
    [Parameter(ParameterSetName = "Columns"), Parameter(ParameterSetName = "Rows"), OfficeLengthTransformation]
    public OfficeLength Gutter { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var document = GetDocument();
        try {
            if (Margin < OfficeLength.Zero || Gutter < OfficeLength.Zero)
                throw new ArgumentException("Margin and Gutter cannot be negative.");
            if (ParameterSetName == "Columns") WriteObject(document.SlideSize.GetColumns(Columns, Margin, Gutter), enumerateCollection: true);
            else if (ParameterSetName == "Rows") WriteObject(document.SlideSize.GetRows(Rows, Margin, Gutter), enumerateCollection: true);
            else WriteObject(document.SlideSize.GetContentBox(Margin));
        } catch (ArgumentException error) { InvalidArgument(error, document); }
    }
}
