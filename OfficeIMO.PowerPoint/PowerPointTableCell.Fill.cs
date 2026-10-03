using DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointTableCell {
        /// <summary>
        /// Gets or sets the explicit RGB cell fill. Setting a color replaces other local fill choices.
        /// Null removes the local solid fill and allows the remaining local or table style fill to apply.
        /// </summary>
        public string? FillColor {
            get => Cell.TableCellProperties?.GetFirstChild<SolidFill>()?.RgbColorModelHex?.Val;
            set {
                Cell.TableCellProperties ??= new TableCellProperties();
                if (value == null) {
                    Cell.TableCellProperties.RemoveAllChildren<SolidFill>();
                    return;
                }
                RemoveCellFillChoices();
                Cell.TableCellProperties.AddChild(new SolidFill(new RgbColorModelHex { Val = value }), true);
            }
        }

        /// <summary>
        /// Gets or sets an explicit transparent cell fill that suppresses the table style fill.
        /// Setting true replaces other local fill choices; setting false removes the no-fill override.
        /// </summary>
        public bool NoFill {
            get => Cell.TableCellProperties?.GetFirstChild<NoFill>() != null;
            set {
                if (!value) {
                    Cell.TableCellProperties?.RemoveAllChildren<NoFill>();
                    return;
                }
                Cell.TableCellProperties ??= new TableCellProperties();
                RemoveCellFillChoices();
                Cell.TableCellProperties.AddChild(new NoFill(), true);
            }
        }

        private void RemoveCellFillChoices() {
            var properties = Cell.TableCellProperties!;
            properties.RemoveAllChildren<NoFill>();
            properties.RemoveAllChildren<SolidFill>();
            properties.RemoveAllChildren<GradientFill>();
            properties.RemoveAllChildren<BlipFill>();
            properties.RemoveAllChildren<PatternFill>();
            properties.RemoveAllChildren<GroupFill>();
        }
    }
}
