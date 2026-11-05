using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Whole-shape text formatting shared by native Visio authoring commands.</summary>
public abstract class OfficeVisioTextCmdlet : OfficeMutationCmdlet {
    /// <summary>Apply bold text; explicit false clears it.</summary>
    [Parameter] public SwitchParameter Bold { get; set; }
    /// <summary>Apply italic text; explicit false clears it.</summary>
    [Parameter] public SwitchParameter Italic { get; set; }
    /// <summary>Apply one strikethrough line; explicit false clears it.</summary>
    [Parameter] public SwitchParameter Strikethrough { get; set; }
    /// <summary>Native Visio underline variant.</summary>
    [Parameter, ValidateSet("None", "Single", "Double")]
    public OfficeTextDecorationStyle? Underline { get; set; }
    /// <summary>Font family for the shape's text.</summary>
    [Parameter, ValidateNotNullOrEmpty] public string? Font { get; set; }
    /// <summary>Font size; a number means points.</summary>
    [Parameter, OfficeLengthTransformation] public OfficeLength? Size { get; set; }
    /// <summary>Text color, including CSS names and theme roles.</summary>
    [Parameter, OfficeColorTransformation, ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? Color { get; set; }
    /// <summary>Whole-shape horizontal text alignment.</summary>
    [Parameter] public OfficeAlignment? Alignment { get; set; }

    /// <inheritdoc />
    protected override void BeginProcessing() {
        if (Size?.ToPoints() <= 0) throw new PSArgumentOutOfRangeException(nameof(Size), Size, "Font size must be positive.");
    }

    /// <summary>Applies only explicitly supplied formatting, retaining native stencil defaults.</summary>
    protected void ApplyTextStyle(VisioShape shape) {
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Bold))) shape.SetBold(Bold.IsPresent);
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Italic))) shape.SetItalic(Italic.IsPresent);
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Strikethrough)))
            shape.SetStrikethrough(Strikethrough.IsPresent ? OfficeTextDecorationStyle.Single : OfficeTextDecorationStyle.None);
        if (Underline.HasValue) shape.SetUnderline(Underline.Value);
        if (Font != null) shape.SetFont(Font);
        if (Size.HasValue) shape.SetFontSize(Size.Value.ToPoints());
        if (Color.HasValue) shape.SetTextColor(Color.Value.Resolve(OfficeTheme.Default));
        if (Alignment.HasValue) shape.SetTextAlignment(Alignment.Value);
    }
}
