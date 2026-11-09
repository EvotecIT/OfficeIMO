using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Word;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Updates native Word cell text, physical geometry and layout flags.</summary>
[Cmdlet(VerbsCommon.Set, "OfficeWordTableCell", DefaultParameterSetName = "Default")]
[OutputType(typeof(WordTableCell))]
public sealed class SetOfficeWordTableCellCommand : OfficeCmdlet {
    /// <summary>Cell for the native cell.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Parameter(ParameterSetName = "WidthPercent", Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public WordTableCell? Cell { get; set; }
    /// <summary>Text for the native cell.</summary>
    [Parameter(ParameterSetName = "Default")]
    [Parameter(ParameterSetName = "WidthPercent")]
    public string? Text { get; set; }
    /// <summary>MarginTop for the native cell.</summary>
    [Parameter(ParameterSetName = "Default")]
    [Parameter(ParameterSetName = "WidthPercent")]
    [AllowNull, OfficeLengthTransformation(AllowNull = true)]
    public OfficeLength? MarginTop { get; set; }
    /// <summary>MarginBottom for the native cell.</summary>
    [Parameter(ParameterSetName = "Default")]
    [Parameter(ParameterSetName = "WidthPercent")]
    [AllowNull, OfficeLengthTransformation(AllowNull = true)]
    public OfficeLength? MarginBottom { get; set; }
    /// <summary>MarginLeft for the native cell.</summary>
    [Parameter(ParameterSetName = "Default")]
    [Parameter(ParameterSetName = "WidthPercent")]
    [AllowNull, OfficeLengthTransformation(AllowNull = true)]
    public OfficeLength? MarginLeft { get; set; }
    /// <summary>MarginRight for the native cell.</summary>
    [Parameter(ParameterSetName = "Default")]
    [Parameter(ParameterSetName = "WidthPercent")]
    [AllowNull, OfficeLengthTransformation(AllowNull = true)]
    public OfficeLength? MarginRight { get; set; }
    /// <summary>BackgroundColor for the native cell.</summary>
    [Parameter(ParameterSetName = "Default")]
    [Parameter(ParameterSetName = "WidthPercent")]
    [OfficeColorTransformation, ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? BackgroundColor { get; set; }
    /// <summary>ShadingPattern for the native cell.</summary>
    [Parameter(ParameterSetName = "Default")]
    [Parameter(ParameterSetName = "WidthPercent")]
    public WordShadingPattern? ShadingPattern { get; set; }
    /// <summary>Width for the native cell.</summary>
    [Parameter(ParameterSetName = "Default")]
    [AllowNull, OfficeLengthTransformation(AllowNull = true)]
    public OfficeLength? Width { get; set; }
    /// <summary>TextDirection for the native cell.</summary>
    [Parameter(ParameterSetName = "Default")]
    [Parameter(ParameterSetName = "WidthPercent")]
    public WordTextDirection? TextDirection { get; set; }
    /// <summary>NoWrap for the native cell.</summary>
    [Parameter(ParameterSetName = "Default")]
    [Parameter(ParameterSetName = "WidthPercent")]
    public SwitchParameter NoWrap { get; set; }
    /// <summary>FitText for the native cell.</summary>
    [Parameter(ParameterSetName = "Default")]
    [Parameter(ParameterSetName = "WidthPercent")]
    public SwitchParameter FitText { get; set; }
    /// <summary>MergeRight for the native cell.</summary>
    [Parameter(ParameterSetName = "Default")]
    [Parameter(ParameterSetName = "WidthPercent")]
    public int? MergeRight { get; set; }
    /// <summary>MergeDown for the native cell.</summary>
    [Parameter(ParameterSetName = "Default")]
    [Parameter(ParameterSetName = "WidthPercent")]
    public int? MergeDown { get; set; }
    /// <summary>SplitHorizontal for the native cell.</summary>
    [Parameter(ParameterSetName = "Default")]
    [Parameter(ParameterSetName = "WidthPercent")]
    public int? SplitHorizontal { get; set; }
    /// <summary>SplitVertical for the native cell.</summary>
    [Parameter(ParameterSetName = "Default")]
    [Parameter(ParameterSetName = "WidthPercent")]
    public int? SplitVertical { get; set; }
    /// <summary>CopyParagraphs for the native cell.</summary>
    [Parameter(ParameterSetName = "Default")]
    [Parameter(ParameterSetName = "WidthPercent")]
    public SwitchParameter CopyParagraphs { get; set; }
    /// <summary>PassThru for the native cell.</summary>
    [Parameter(ParameterSetName = "Default")]
    [Parameter(ParameterSetName = "WidthPercent")]
    public SwitchParameter PassThru { get; set; }
    /// <summary>WidthPercent for the native cell.</summary>
    [Parameter(ParameterSetName = "WidthPercent", Mandatory = true)]
    public double? WidthPercent { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (Cell == null) throw new PSArgumentNullException(nameof(Cell));
        ValidateMargin(MarginTop, nameof(MarginTop)); ValidateMargin(MarginBottom, nameof(MarginBottom));
        ValidateMargin(MarginLeft, nameof(MarginLeft)); ValidateMargin(MarginRight, nameof(MarginRight));
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Text))) Cell.AddParagraph(Text ?? string.Empty, removeExistingParagraphs: true);
        if (MyInvocation.BoundParameters.ContainsKey(nameof(MarginTop))) Cell.MarginTop = MarginTop;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(MarginBottom))) Cell.MarginBottom = MarginBottom;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(MarginLeft))) Cell.MarginLeft = MarginLeft;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(MarginRight))) Cell.MarginRight = MarginRight;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(BackgroundColor))) Cell.ShadingFillColor = BackgroundColor;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(ShadingPattern))) Cell.ShadingPattern = ShadingPattern;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Width))) Cell.Width = Width.HasValue ? WordTableWidth.Length(Width.Value) : null;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(WidthPercent))) Cell.Width = WidthPercent.HasValue ? WordTableWidth.Percent(WidthPercent.Value) : null;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(TextDirection))) Cell.TextDirection = TextDirection;
        if (BoundSwitch(nameof(NoWrap)) is bool noWrap) Cell.WrapText = !noWrap;
        if (BoundSwitch(nameof(FitText)) is bool fit) Cell.FitText = fit;
        if (MergeRight.HasValue) Cell.MergeHorizontally(MergeRight.Value, CopyParagraphs);
        if (MergeDown.HasValue) Cell.MergeVertically(MergeDown.Value, CopyParagraphs);
        if (SplitHorizontal.HasValue) Cell.SplitHorizontally(SplitHorizontal.Value);
        if (SplitVertical.HasValue) Cell.SplitVertically(SplitVertical.Value);
        if (PassThru) WriteObject(Cell);
    }
    private static void ValidateMargin(OfficeLength? value, string name) {
        if (value.HasValue && (value.Value < OfficeLength.Zero || value.Value.ToTwips() > short.MaxValue))
            throw new PSArgumentOutOfRangeException(name, value, "Cell margins must fit non-negative signed 16-bit twips.");
    }
}
