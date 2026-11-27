using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointTableCell {
        private static readonly string[] CellFills = { "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill" };

        /// <summary>
        /// Gets or sets the local cell fill. Setting a color replaces other local fill choices.
        /// Null removes the local solid fill and allows the remaining local or table style fill to apply.
        /// </summary>
        public OfficeColor? FillColor {
            get => PowerPointDrawingColor.Read(PropertiesElement.Child(D, "solidFill"));
            set {
                OfficeTheme? theme = ThemeFor(value);
                OfficeXmlNode properties = EnsureProperties();
                if (value == null) {
                    properties.RemoveChildren(D, "solidFill");
                } else {
                    foreach (string name in CellFills) properties.RemoveChildren(D, name);
                    properties.InsertOrdered(PowerPointDrawingColor.CreateFill(properties, value.Value, theme), ATableCellPropertiesView.Order);
                }
                MarkEdited();
            }
        }

        /// <summary>
        /// Gets or sets an explicit transparent cell fill that suppresses the table style fill.
        /// Setting true replaces other local fill choices; setting false removes the no-fill override.
        /// </summary>
        public bool NoFill {
            get => !PropertiesElement.Child(D, "noFill").IsNone;
            set {
                if (!value) {
                    PropertiesElement.RemoveChildren(D, "noFill");
                } else {
                    OfficeXmlNode properties = EnsureProperties();
                    foreach (string name in CellFills) properties.RemoveChildren(D, name);
                    properties.AddOrdered(D, "noFill", ATableCellPropertiesView.Order);
                }
                MarkEdited();
            }
        }
    }
}
