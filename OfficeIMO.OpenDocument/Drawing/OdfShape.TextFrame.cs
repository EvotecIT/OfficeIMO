using System.Text.RegularExpressions;
using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

public abstract partial class OdfShape {
    /// <summary>Inherited horizontal text-area placement; null removes the local override.</summary>
    /// <remarks><see cref="OfficeTextAreaAlignment.FullWidth"/> writes the native <c>justify</c> value. Text-area placement is separate from paragraph alignment.</remarks>
    /// <exception cref="NotSupportedException">The effective native value is outside the supported alignment values.</exception>
    public OfficeTextAreaAlignment? TextAreaAlignment {
        get => ReadGraphicProperty(OdfNamespaces.Draw + "textarea-horizontal-align") switch {
            null => null, "justify" => OfficeTextAreaAlignment.FullWidth,
            "left" => OfficeTextAreaAlignment.Left, "center" => OfficeTextAreaAlignment.Center, "right" => OfficeTextAreaAlignment.Right,
            _ => throw new NotSupportedException("Unsupported ODF text-area horizontal alignment.")
        };
        set {
            if (value.HasValue && !Enum.IsDefined(typeof(OfficeTextAreaAlignment), value.Value)) throw new ArgumentOutOfRangeException(nameof(value));
            string? lexical = value == OfficeTextAreaAlignment.FullWidth ? "justify" : value?.ToString().ToLowerInvariant();
            EnsureGraphicStyle().SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "textarea-horizontal-align", lexical);
        }
    }

    /// <summary>Inherited native vertical text-area placement; null removes the local override.</summary>
    /// <remarks>Native <see cref="OdfTextAreaVerticalAlignment.Justify"/> is preserved by editing and saving, but remains outside shared drawing projection.</remarks>
    /// <exception cref="NotSupportedException">The effective native value is outside the supported alignment values.</exception>
    public OdfTextAreaVerticalAlignment? TextVerticalAlignment {
        get => ReadGraphicProperty(OdfNamespaces.Draw + "textarea-vertical-align") switch {
            null => null, "top" => OdfTextAreaVerticalAlignment.Top, "middle" => OdfTextAreaVerticalAlignment.Middle,
            "bottom" => OdfTextAreaVerticalAlignment.Bottom, "justify" => OdfTextAreaVerticalAlignment.Justify,
            _ => throw new NotSupportedException("Unsupported ODF text-area vertical alignment.")
        };
        set {
            if (value.HasValue && !Enum.IsDefined(typeof(OdfTextAreaVerticalAlignment), value.Value)) throw new ArgumentOutOfRangeException(nameof(value));
            EnsureGraphicStyle().SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "textarea-vertical-align", value?.ToString().ToLowerInvariant());
        }
    }

    /// <summary>Inherited native wrapping declaration; null removes the local override.</summary>
    /// <remarks>Writes <c>wrap</c> or <c>no-wrap</c>. Native behavior and shared projection depend on the shape kind.</remarks>
    /// <exception cref="NotSupportedException">The effective native wrapping declaration is unknown.</exception>
    public bool? WrapText {
        get => ReadGraphicProperty(OdfNamespaces.Fo + "wrap-option") switch {
            null => null, "wrap" => true, "no-wrap" => false,
            _ => throw new NotSupportedException("Unsupported ODF text wrapping option.")
        };
        set => EnsureGraphicStyle().SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "wrap-option",
            value.HasValue ? value.Value ? "wrap" : "no-wrap" : null);
    }

    /// <summary>Inherited top, right, bottom, and left text insets; null removes all local padding declarations.</summary>
    /// <remarks>
    /// Each side resolves its explicit value or <c>fo:padding</c> shorthand at each style level before visiting the parent.
    /// An omitted side is zero; all sides omitted returns null. Setting writes four explicit sides and removes the local shorthand.
    /// Native lexical lengths, including pixels, are retained. Pixels remain outside point-based shared drawing projection.
    /// </remarks>
    /// <exception cref="ArgumentException">An assigned side is not a finite nonnegative absolute ODF schema length.</exception>
    public OdfInsets? TextPadding {
        get {
            string? top = ReadGraphicProperty(OdfNamespaces.Fo + "padding-top", OdfNamespaces.Fo + "padding");
            string? right = ReadGraphicProperty(OdfNamespaces.Fo + "padding-right", OdfNamespaces.Fo + "padding");
            string? bottom = ReadGraphicProperty(OdfNamespaces.Fo + "padding-bottom", OdfNamespaces.Fo + "padding");
            string? left = ReadGraphicProperty(OdfNamespaces.Fo + "padding-left", OdfNamespaces.Fo + "padding");
            return top == null && right == null && bottom == null && left == null ? null
                : new OdfInsets(OdfLength.Parse(top ?? "0cm"), OdfLength.Parse(right ?? "0cm"),
                    OdfLength.Parse(bottom ?? "0cm"), OdfLength.Parse(left ?? "0cm"));
        }
        set {
            if (value.HasValue) {
                ValidateTextPaddingLength(value.Value.Top); ValidateTextPaddingLength(value.Value.Right);
                ValidateTextPaddingLength(value.Value.Bottom); ValidateTextPaddingLength(value.Value.Left);
            }
            OdfStyle style = EnsureGraphicStyle(); // Validate every edge before changing a shared style or its owner reference.
            style.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "padding", null);
            style.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "padding-top", value?.Top.ToString());
            style.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "padding-right", value?.Right.ToString());
            style.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "padding-bottom", value?.Bottom.ToString());
            style.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "padding-left", value?.Left.ToString());
        }
    }

    private static void ValidateTextPaddingLength(OdfLength length) {
        string lexical = length.ToString();
        // ODF's nonNegativeLength grammar includes px even though OdfLength cannot convert pixels to points.
        if (!Regex.IsMatch(lexical, @"\A(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:cm|mm|in|pt|pc|px)\z", RegexOptions.CultureInvariant) ||
            !double.TryParse(lexical.Substring(0, lexical.Length - 2), NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out double magnitude) ||
            double.IsNaN(magnitude) || double.IsInfinity(magnitude) ||
            !lexical.EndsWith("px", StringComparison.Ordinal) && !length.TryToPoints(out _))
            throw new ArgumentException("Text padding must use finite nonnegative absolute ODF lengths.", "value");
    }
}
