using System;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        ///     Adds a title textbox to the slide.
        /// </summary>
        public PowerPointTextBox AddTitle(string text, OfficeLength? left = null, OfficeLength? top = null,
            OfficeLength? width = null, OfficeLength? height = null) {
            if (text == null) throw new ArgumentNullException(nameof(text));
            if (!left.HasValue && !top.HasValue && !width.HasValue && !height.HasValue)
                return AddAuthoringTitle(text);
            PowerPointLayoutBox defaults = left.HasValue && top.HasValue && width.HasValue && height.HasValue
                ? new PowerPointLayoutBox(left.Value, top.Value, width.Value, height.Value) : GetAuthoringTitleBounds(text);
            long nativeLeft = (left ?? defaults.Left).ToEmus();
            long nativeTop = (top ?? defaults.Top).ToEmus();
            long nativeWidth = (width ?? defaults.Width).ToEmus();
            long nativeHeight = (height ?? defaults.Height).ToEmus();
            PreflightTitleReflow(top ?? defaults.Top, height ?? defaults.Height);

            string name = GenerateUniqueName("Title");
            Shape shape = new(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = AllocateShapeId(), Name = name },
                    new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape { Type = PlaceholderValues.Title })
                ),
                new ShapeProperties(
                    new A.Transform2D(new A.Offset { X = nativeLeft, Y = nativeTop }, new A.Extents { Cx = nativeWidth, Cy = nativeHeight }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                ),
                new TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(new A.Run(new A.Text(text)))
                )
            );

            CommonSlideData data = SlideRoot.CommonSlideData ??= new CommonSlideData(new ShapeTree());
            ShapeTree tree = data.ShapeTree ??= new ShapeTree();
            tree.AppendChild(shape);
            PowerPointTextBox textBox = TrackShape(new PowerPointTextBox(shape, _slidePart));
            ReflowAuthoringContent();
            return textBox;
        }

        /// <summary>
        ///     Adds a title textbox using a layout box.
        /// </summary>
        public PowerPointTextBox AddTitle(string text, PowerPointLayoutBox layout) {
            return AddTitle(text, layout.Left, layout.Top, layout.Width, layout.Height);
        }

        /// <summary>
        ///     Adds a textbox with the specified text.
        /// </summary>
        public PowerPointTextBox AddTextBox(string text, OfficeLength? left = null, OfficeLength? top = null, OfficeLength? width = null,
            OfficeLength? height = null) {
            long nativeLeft = (left ?? OfficeLength.Emus(838200L)).ToEmus();
            long nativeTop = (top ?? OfficeLength.Emus(2174875L)).ToEmus();
            long nativeWidth = (width ?? OfficeLength.Emus(7772400L)).ToEmus();
            long nativeHeight = (height ?? OfficeLength.Emus(3962400L)).ToEmus();
            if (text == null) {
                throw new ArgumentNullException(nameof(text));
            }

            string name = GenerateUniqueName("TextBox");
            Shape shape = new(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = AllocateShapeId(), Name = name },
                    new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }) { TextBox = true },
                    new ApplicationNonVisualDrawingProperties()
                ),
                new ShapeProperties(
                    new A.Transform2D(new A.Offset { X = nativeLeft, Y = nativeTop }, new A.Extents { Cx = nativeWidth, Cy = nativeHeight }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                ),
                new TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(new A.Run(new A.Text(text)))
                )
            );

            CommonSlideData data = SlideRoot.CommonSlideData ??= new CommonSlideData(new ShapeTree());
            ShapeTree tree = data.ShapeTree ??= new ShapeTree();
            tree.AppendChild(shape);
            PowerPointTextBox textBox = TrackShape(new PowerPointTextBox(shape, _slidePart));
            return textBox;
        }

        /// <summary>
        ///     Adds a textbox using a layout box.
        /// </summary>
        public PowerPointTextBox AddTextBox(string text, PowerPointLayoutBox layout) {
            return AddTextBox(text, layout.Left, layout.Top, layout.Width, layout.Height);
        }

        /// <summary>
        ///     Adds an empty textbox using a layout box.
        /// </summary>
        public PowerPointTextBox AddTextBox(PowerPointLayoutBox layout) {
            return AddTextBox(string.Empty, layout.Left, layout.Top, layout.Width, layout.Height);
        }

    }
}
