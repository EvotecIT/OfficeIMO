using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// The SDK side of a text box: constructors over SDK shapes and SDK views of it, for code that still works on SDK
    /// elements (layouts, masters, footers, the layout engine and image export).
    /// </summary>
    public partial class PowerPointTextBox {
        internal PowerPointTextBox(Shape shape, SlidePart? slidePart = null) : base(shape, slidePart) {
        }

        internal PowerPointTextBox(Shape shape, OpenXmlPartContainer ownerPart) : base(shape, ownerPart) {
        }

        private Shape Shape => (Shape)Element;

        /// <summary>The slide part the text box is on, or null for a text box of a layout or master.</summary>
        internal SlidePart? SlidePart => SlidePartRef.Current.SdkPart as SlidePart ?? (OwnerSlide is { } slide ? slide.SlidePart : null);

        internal TextBody? TextBody => Shape.TextBody;

        internal OpenXmlCompositeElement? MasterTextStyle {
            get {
                TextStyles? styles = SlidePart?.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.TextStyles;
                PlaceholderValues? placeholder = Shape.NonVisualShapeProperties?
                    .ApplicationNonVisualDrawingProperties?
                    .GetFirstChild<PlaceholderShape>()?
                    .Type?.Value;
                if (placeholder == PlaceholderValues.Title || placeholder == PlaceholderValues.CenteredTitle) return styles?.TitleStyle;
                if (placeholder == PlaceholderValues.Body || placeholder == PlaceholderValues.Object || placeholder == PlaceholderValues.SubTitle) return styles?.BodyStyle;
                return styles?.OtherStyle;
            }
        }
    }
}
