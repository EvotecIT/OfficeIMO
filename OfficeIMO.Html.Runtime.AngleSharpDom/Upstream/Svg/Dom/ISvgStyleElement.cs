namespace AngleSharp.Svg.Dom
{
    using AngleSharp.Attributes;
    using AngleSharp.Dom;
    using System;

    /// <summary>
    /// Represents a style SVG element.
    /// </summary>
    [DomName("SVGStyleElement")]
    public interface ISvgStyleElement : ISvgElement, ILinkStyle
    {
        /// <summary>
        /// Gets or sets if the style is enabled or disabled.
        /// </summary>
        [DomName("disabled")]
        Boolean IsDisabled { get; set; }

        /// <summary>
        /// Gets or sets the use with one or more target media.
        /// </summary>
        [DomName("media")]
        String? Media { get; set; }

        /// <summary>
        /// Gets or sets the content type of the style sheet language.
        /// </summary>
        [DomName("type")]
        String? Type { get; set; }
    }
}
