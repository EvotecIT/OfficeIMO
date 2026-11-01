using System;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Updates a template layout placeholder's geometry, text margins and text style.</summary>
/// <para>Omitted properties keep their current values. Master and layout positions start at one.
/// Numbers mean points; text margins use one, two or four values in CSS order.</para>
/// <example><summary>Adjust a template title area and style.</summary>
/// <code>Set-OfficePowerPointLayoutPlaceholder -Document $deck -TemplateLayout 'Title and Content' -Type Title -X 2cm -Width 24cm -TextMargin 6pt,8pt -Size 32pt -Bold</code></example>
[Cmdlet(VerbsCommon.Set, "OfficePowerPointLayoutPlaceholder", DefaultParameterSetName = "Document")]
[OutputType(typeof(PowerPointTextBox))]
public sealed class SetOfficePowerPointLayoutPlaceholderCommand : OfficeDocumentCmdlet<PowerPointDocument> {
    /// <summary>Gets or sets a layout name, one-based position, native type or returned layout metadata.</summary>
    [Parameter(Mandatory = true), PowerPointLayoutTransformation]
    public PowerPointLayoutSelector TemplateLayout { get; set; } = null!;
    /// <summary>Gets or sets the one-based master position, defaulting to the first or to the layout metadata's master.</summary>
    [Parameter, ValidateRange(1, int.MaxValue)]
    public int Master { get; set; } = 1;
    /// <summary>Gets or sets the placeholder type.</summary>
    [Parameter(Mandatory = true)]
    public PowerPointPlaceholderType Type { get; set; }
    /// <summary>Gets or sets the optional native placeholder index.</summary>
    [Parameter]
    public uint? Index { get; set; }
    /// <summary>Gets or sets the left position; negative positions are allowed.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? X { get; set; }
    /// <summary>Gets or sets the top position; negative positions are allowed.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? Y { get; set; }
    /// <summary>Gets or sets the positive placeholder width.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? Width { get; set; }
    /// <summary>Gets or sets the positive placeholder height.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? Height { get; set; }
    /// <summary>Gets or sets margins: all; vertical,horizontal; or top,right,bottom,left.</summary>
    [Parameter, OfficeLengthArrayTransformation, ValidateNotNullOrEmpty]
    public OfficeLength[]? TextMargin { get; set; }
    /// <summary>Gets or sets a native text preset.</summary>
    [Parameter, ValidateSet("Title", "Subtitle", "Body", "Caption", "Emphasis")]
    public string? Style { get; set; }
    /// <summary>Gets or sets the font size, from one through 4000 points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? Size { get; set; }
    /// <summary>Gets or sets the Latin font family.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? Font { get; set; }
    /// <summary>Gets or sets text color, resolving theme roles against the selected master.</summary>
    [Parameter, OfficeColorTransformation]
    public OfficeColor? Color { get; set; }
    /// <summary>Gets or sets the highlight color.</summary>
    [Parameter, OfficeColorTransformation]
    public OfficeColor? Highlight { get; set; }
    /// <summary>Gets or sets bold formatting; an explicitly false switch clears it.</summary>
    [Parameter]
    public SwitchParameter Bold { get; set; }
    /// <summary>Gets or sets italic formatting; an explicitly false switch clears it.</summary>
    [Parameter]
    public SwitchParameter Italic { get; set; }
    /// <summary>Gets or sets the native underline style; None removes underlining.</summary>
    [Parameter]
    public PowerPointUnderlineStyle? Underline { get; set; }
    /// <summary>Gets or sets the zero-based paragraph indentation level.</summary>
    [Parameter, ValidateRange(0, 8)]
    public int? Level { get; set; }
    /// <summary>Gets or sets a single bullet character, exclusive with Numbering.</summary>
    [Parameter, ValidateLength(1, 1)]
    public string? BulletChar { get; set; }
    /// <summary>Gets or sets the native numbering scheme, exclusive with BulletChar.</summary>
    [Parameter]
    public PowerPointNumberingScheme? Numbering { get; set; }
    /// <summary>Gets or sets whether to create a missing placeholder.</summary>
    [Parameter]
    public SwitchParameter CreateIfMissing { get; set; }
    /// <summary>Gets or sets whether to emit the updated native text box.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        PowerPointDocument document = GetDocument();
        try {
            if (BulletChar != null && Numbering.HasValue)
                throw new ArgumentException("Specify either BulletChar or Numbering, not both.");
            if (Width.HasValue && Width.Value <= OfficeLength.Zero || Height.HasValue && Height.Value <= OfficeLength.Zero)
                throw new ArgumentException("Placeholder width and height must be positive.");
            if (TextMargin != null) {
                if (TextMargin.Length is not (1 or 2 or 4))
                    throw new ArgumentException("TextMargin takes one, two or four physical lengths.");
                foreach (var margin in TextMargin) {
                    if (margin < OfficeLength.Zero || margin.ToEmus() > int.MaxValue)
                        throw new ArgumentOutOfRangeException(nameof(TextMargin), "Text margins must fit a nonnegative DrawingML inset.");
                }
            }
            PowerPointTextDefaults.ToDrawingFontSize(Size?.ToPoints(), nameof(Size));
            if (Underline.HasValue && !Enum.IsDefined(typeof(PowerPointUnderlineStyle), Underline.Value))
                throw new ArgumentOutOfRangeException(nameof(Underline));
            if (Numbering.HasValue && !Enum.IsDefined(typeof(PowerPointNumberingScheme), Numbering.Value))
                throw new ArgumentOutOfRangeException(nameof(Numbering));
            if (!Enum.IsDefined(typeof(PowerPointPlaceholderType), Type))
                throw new ArgumentOutOfRangeException(nameof(Type));

            var layout = TemplateLayout.Resolve(document, Master, MyInvocation.BoundParameters.ContainsKey(nameof(Master)));
            var theme = new PowerPointLayoutContext(document, document.ReadLayoutTheme(layout.MasterIndex));
            var style = new PowerPointTextStyle();
            if (Style != null && !OpenXmlValueParser.TryParse(Style, out style))
                throw new ArgumentException("Unknown text style.", nameof(Style));
            if (Size.HasValue) style = style.WithFontSize(Size);
            if (Font != null) style = style.WithFontName(Font);
            if (Color.HasValue) style = style.WithColor(theme.ResolveColor(Color));
            if (Highlight.HasValue) style = style.WithHighlightColor(theme.ResolveColor(Highlight));
            if (BoundSwitch(nameof(Bold)) is bool bold) style = style.WithBold(bold);
            if (BoundSwitch(nameof(Italic)) is bool italic) style = style.WithItalic(italic);
            if (Underline.HasValue) style = style.WithUnderlineStyle(Underline);

            var box = CreateIfMissing
                ? document.EnsureLayoutPlaceholderTextBox(layout.MasterIndex, layout.LayoutIndex, Type, Index)
                : document.GetLayoutPlaceholderTextBox(layout.MasterIndex, layout.LayoutIndex, Type, Index)
                    ?? throw new ArgumentException("The selected layout placeholder does not exist; use CreateIfMissing to add it.");
            if (X.HasValue || Y.HasValue || Width.HasValue || Height.HasValue) {
                var bounds = document.GetEffectiveLayoutPlaceholderBounds(layout.MasterIndex, layout.LayoutIndex, Type, Index)
                    ?? box.Bounds;
                box.Bounds = new PowerPointLayoutBox(X ?? bounds.Left, Y ?? bounds.Top, Width ?? bounds.Width, Height ?? bounds.Height);
            }
            if (TextMargin != null) {
                OfficeLength top = TextMargin[0], right = TextMargin.Length == 1 ? top : TextMargin[1];
                OfficeLength bottom = TextMargin.Length == 4 ? TextMargin[2] : top;
                OfficeLength left = TextMargin.Length == 4 ? TextMargin[3] : right;
                box.SetTextMargins(left, top, right, bottom);
            }
            bool hasStyle = Style != null || Size.HasValue || Font != null || Color.HasValue || Highlight.HasValue
                || BoundSwitch(nameof(Bold)).HasValue || BoundSwitch(nameof(Italic)).HasValue
                || Underline.HasValue || Level.HasValue || BulletChar != null || Numbering.HasValue;
            if (hasStyle) {
                document.SetLayoutPlaceholderTextStyle(layout.MasterIndex, layout.LayoutIndex, Type, style, Index,
                    Level, BulletChar?[0], Numbering);
            }
            if (PassThru) WriteObject(box);
        } catch (ArgumentException error) { InvalidArgument(error, document); }
        catch (Exception error) when (error is not (PipelineStoppedException or ActionPreferenceStopException)) {
            ItemError(error, "powerpoint-layout-placeholder-edit-failed", document);
        }
    }
}
