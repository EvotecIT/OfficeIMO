using System.Management.Automation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Updates OfficeIMO Word table-cell content, layout, and merge settings.</summary>
/// <example>
///   <summary>Replace text in a cell after finding a table.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc = Get-OfficeWord -Path .\Handover.docx
/// $table = Find-OfficeWordTable -Document $doc -Text 'Risk marker' | Select-Object -First 1
/// $table |
///     Get-OfficeWordTableCell -Row 2 -Column 2 |
///     Set-OfficeWordTableCell -Text 'Investigating' -ShadingFillColor '#fff2cc' -ShadingPattern Clear
/// $doc | Close-OfficeWord -Save</code>
///   <para>Finds an existing table by text, replaces a target cell value, applies shading, and saves the document.</para>
/// </example>
/// <example>
///   <summary>Highlight a status column in the first report table.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc = Get-OfficeWord -Path .\Report.docx
/// $table = $doc | Get-OfficeWordTable | Select-Object -First 1
/// $table |
///     Get-OfficeWordTableCell -Column 2 |
///     Set-OfficeWordTableCell -ShadingFillColor '#fff1f0' -ShadingPattern Clear -Width 2400 -WidthType Dxa
/// $doc | Save-OfficeWord -Path .\Report-StatusCells.docx</code>
///   <para>Reads cells from an OfficeIMO table object, applies cell shading and width, and saves the updated document.</para>
/// </example>
/// <example>
///   <summary>Add readable padding to every cell in a report table.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$table | Get-OfficeWordTableCell | Set-OfficeWordTableCell -MarginTopCentimeters 0.12 -MarginBottomCentimeters 0.12 -MarginLeftCentimeters 0.16 -MarginRightCentimeters 0.16</code>
///   <para>Sets independent cell margins in centimeters. Omitted margins retain their values; pass null to remove an explicit margin.</para>
/// </example>
/// <example>
///   <summary>Merge a heading row across columns.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc = Get-OfficeWord -Path .\Report.docx
/// $table = $doc | Get-OfficeWordTable | Select-Object -First 1
/// $table |
///     Get-OfficeWordTableCell -Row 0 -Column 0 |
///     Set-OfficeWordTableCell -MergeRight 2 -CopyParagraphs
/// $doc | Save-OfficeWord -Path .\Report-MergedHeader.docx</code>
///   <para>Uses the OfficeIMO merge operation exposed by the thin table-cell wrapper.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeWordTableCell")]
[Alias("WordTableCellStyle")]
[OutputType(typeof(WordTableCell))]
public sealed class SetOfficeWordTableCellCommand : PSCmdlet
{
    /// <summary>Table cell to update.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
    public WordTableCell Cell { get; set; } = null!;

    /// <summary>Replace the visible cell text.</summary>
    [Parameter] public string? Text { get; set; }

    /// <summary>Top cell padding in centimeters. Pass null to remove the explicit margin.</summary>
    [Parameter] public double? MarginTopCentimeters { get; set; }

    /// <summary>Bottom cell padding in centimeters. Pass null to remove the explicit margin.</summary>
    [Parameter] public double? MarginBottomCentimeters { get; set; }

    /// <summary>Left cell padding in centimeters. Pass null to remove the explicit margin.</summary>
    [Parameter] public double? MarginLeftCentimeters { get; set; }

    /// <summary>Right cell padding in centimeters. Pass null to remove the explicit margin.</summary>
    [Parameter] public double? MarginRightCentimeters { get; set; }

    /// <summary>Cell shading fill color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? ShadingFillColor { get; set; }

    /// <summary>Cell shading pattern.</summary>
    [Parameter]
    public WordShadingPattern? ShadingPattern { get; set; }

    /// <summary>Cell width value.</summary>
    [Parameter] public int? Width { get; set; }

    /// <summary>Cell width unit type.</summary>
    [Parameter]
    public WordTableWidthUnit? WidthType { get; set; }

    /// <summary>Cell text direction.</summary>
    [Parameter] public WordTextDirection? TextDirection { get; set; }

    /// <summary>Whether text wraps in the cell.</summary>
    [Parameter] public bool? WrapText { get; set; }

    /// <summary>Whether text should fit within the cell.</summary>
    [Parameter] public bool? FitText { get; set; }

    /// <summary>Number of cells to merge to the right.</summary>
    [Parameter] public int? MergeRight { get; set; }

    /// <summary>Number of cells to merge downward.</summary>
    [Parameter] public int? MergeDown { get; set; }

    /// <summary>Number of columns to split the cell into.</summary>
    [Parameter] public int? SplitHorizontal { get; set; }

    /// <summary>Number of rows to split the cell into.</summary>
    [Parameter] public int? SplitVertical { get; set; }

    /// <summary>Copy paragraphs while merging cells.</summary>
    [Parameter] public SwitchParameter CopyParagraphs { get; set; }

    /// <summary>Emit the updated table cell.</summary>
    [Parameter] public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Cell == null)
        {
            return;
        }

        ValidateMargin(MarginTopCentimeters, nameof(MarginTopCentimeters));
        ValidateMargin(MarginBottomCentimeters, nameof(MarginBottomCentimeters));
        ValidateMargin(MarginLeftCentimeters, nameof(MarginLeftCentimeters));
        ValidateMargin(MarginRightCentimeters, nameof(MarginRightCentimeters));

        if (MyInvocation.BoundParameters.ContainsKey(nameof(Text))) Cell.AddParagraph(Text ?? string.Empty, removeExistingParagraphs: true);
        if (MyInvocation.BoundParameters.ContainsKey(nameof(MarginTopCentimeters))) Cell.MarginTop= ((MarginTopCentimeters) is { } wordLength6025 ? global::OfficeIMO.OfficeLength.Centimeters(wordLength6025) : (global::OfficeIMO.OfficeLength?)null);
        if (MyInvocation.BoundParameters.ContainsKey(nameof(MarginBottomCentimeters))) Cell.MarginBottom= ((MarginBottomCentimeters) is { } wordLength6165 ? global::OfficeIMO.OfficeLength.Centimeters(wordLength6165) : (global::OfficeIMO.OfficeLength?)null);
        if (MyInvocation.BoundParameters.ContainsKey(nameof(MarginLeftCentimeters))) Cell.MarginLeft= ((MarginLeftCentimeters) is { } wordLength6304 ? global::OfficeIMO.OfficeLength.Centimeters(wordLength6304) : (global::OfficeIMO.OfficeLength?)null);
        if (MyInvocation.BoundParameters.ContainsKey(nameof(MarginRightCentimeters))) Cell.MarginRight= ((MarginRightCentimeters) is { } wordLength6443 ? global::OfficeIMO.OfficeLength.Centimeters(wordLength6443) : (global::OfficeIMO.OfficeLength?)null);
        if (MyInvocation.BoundParameters.ContainsKey(nameof(ShadingFillColor))) Cell.ShadingFillColorHex = ShadingFillColor ?? string.Empty;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(ShadingPattern))) Cell.ShadingPattern = ShadingPattern;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Width))) Cell.Width = Width;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(WidthType))) Cell.WidthType = WidthType;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(TextDirection))) Cell.TextDirection = TextDirection;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(WrapText))) Cell.WrapText = WrapText ?? false;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(FitText))) Cell.FitText = FitText ?? false;
        if (MergeRight.HasValue) Cell.MergeHorizontally(MergeRight.Value, CopyParagraphs.IsPresent);
        if (MergeDown.HasValue) Cell.MergeVertically(MergeDown.Value, CopyParagraphs.IsPresent);
        if (SplitHorizontal.HasValue) Cell.SplitHorizontally(SplitHorizontal.Value);
        if (SplitVertical.HasValue) Cell.SplitVertically(SplitVertical.Value);

        if (PassThru.IsPresent)
        {
            WriteObject(Cell);
        }
    }

    private static void ValidateMargin(double? value, string name)
    {
        // OfficeIMO stores cell margins as signed 16-bit twips. Null removes an override.
        if (value.HasValue && (double.IsNaN(value.Value) || value.Value < 0 || value.Value > 57.7))
        {
            throw new PSArgumentOutOfRangeException(name, value, "Cell margins must be between 0 and 57.7 centimeters, or null.");
        }
    }
}
