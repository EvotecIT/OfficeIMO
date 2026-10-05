using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint;

public partial class PowerPointTextRun {
    /// <summary>
    /// Gets or sets whether the existing click hyperlink uses this run's explicit or inherited text color
    /// instead of the theme hyperlink color. Set the hyperlink before assigning this property.
    /// </summary>
    /// <remarks>
    /// Clearing <see cref="Color"/> preserves this policy and allows a link using text color to inherit its color.
    /// The native color extension is supported by Office 2019 and newer; readers may ignore it.
    /// </remarks>
    /// <exception cref="InvalidOperationException">There is no click hyperlink to format.</exception>
    public bool HyperlinkUsesTextColor {
        get => RunProperties?.GetFirstChild<A.HyperlinkOnClick>() is A.HyperlinkOnClick hyperlink
            && PowerPointTextHyperlinkColor.ReadChoice(hyperlink) == true;
        set {
            A.HyperlinkOnClick? hyperlink = RunProperties?.GetFirstChild<A.HyperlinkOnClick>();
            if (hyperlink == null) throw new InvalidOperationException("Set a click hyperlink before its color policy.");
            PowerPointTextHyperlinkColor.SetChoice(hyperlink, value);
        }
    }
}
