using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Helpers for turning Shape Data values into visible diagram adornments.
    /// </summary>
    public static class VisioDataGraphicExtensions {
        /// <summary>
        /// Adds visible data graphics for one target shape.
        /// </summary>
        /// <param name="page">Page that owns the target shape.</param>
        /// <param name="target">Shape whose Shape Data values should be visualized.</param>
        /// <param name="dataGraphic">Data graphic definition.</param>
        public static IReadOnlyList<VisioShape> AddDataGraphics(this VisioPage page, VisioShape target, VisioDataGraphic dataGraphic) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            if (target == null) {
                throw new ArgumentNullException(nameof(target));
            }

            if (dataGraphic == null) {
                throw new ArgumentNullException(nameof(dataGraphic));
            }

            if (!page.AllShapes().Contains(target)) {
                throw new InvalidOperationException("The data graphic target shape must belong to the page.");
            }

            ValidateDataGraphic(dataGraphic);
            List<VisioShape> shapes = new();
            int itemIndex = 0;
            foreach (VisioDataGraphicItem item in dataGraphic.Items) {
                string? value = target.GetShapeDataValue(item.FieldName);
                if (string.IsNullOrWhiteSpace(value) && !dataGraphic.ShowEmptyValues) {
                    itemIndex++;
                    continue;
                }

                switch (item.Kind) {
                    case VisioDataGraphicItemKind.TextBadge:
                        shapes.Add(AddBadge(page, target, item, value ?? string.Empty, dataGraphic, itemIndex));
                        break;
                    case VisioDataGraphicItemKind.DataBar:
                        shapes.AddRange(AddDataBar(page, target, item, value ?? string.Empty, dataGraphic, itemIndex));
                        break;
                    default:
                        throw new InvalidOperationException("Unsupported data graphic item kind.");
                }

                itemIndex++;
            }

            return shapes;
        }

        /// <summary>
        /// Adds visible data graphics for every selected shape.
        /// </summary>
        /// <param name="selection">Target shape selection.</param>
        /// <param name="dataGraphic">Data graphic definition.</param>
        public static IReadOnlyList<VisioShape> AddDataGraphics(this VisioShapeSelection selection, VisioDataGraphic dataGraphic) {
            if (selection == null) {
                throw new ArgumentNullException(nameof(selection));
            }

            if (selection.OwnerPage == null) {
                throw new InvalidOperationException("The shape selection must be created from a VisioPage query to add data graphics.");
            }

            List<VisioShape> shapes = new();
            foreach (VisioShape shape in selection) {
                shapes.AddRange(selection.OwnerPage.AddDataGraphics(shape, dataGraphic));
            }

            return shapes;
        }

        private static VisioShape AddBadge(VisioPage page, VisioShape target, VisioDataGraphicItem item, string value, VisioDataGraphic dataGraphic, int itemIndex) {
            GetItemCenter(target, dataGraphic, itemIndex, item.Width, item.Height, out double x, out double y);
            string id = UniqueId(page, target.Id + "-dg-" + SanitizeId(item.FieldName) + "-" + itemIndex);
            string label = item.FormatText(value);
            VisioShape badge = new VisioShape(id, x, y, item.Width, item.Height, label) {
                Name = "Data Graphic",
                NameU = "OfficeIMO Data Graphic Badge"
            };
            item.GetBadgeStyle().ApplyTo(badge);
            MarkDataGraphic(badge, target, item, value, "Badge");
            page.Shapes.Add(badge);
            if (!string.IsNullOrWhiteSpace(dataGraphic.LayerName)) {
                page.AddToLayer(dataGraphic.LayerName!, badge);
            }

            return badge;
        }

        private static IReadOnlyList<VisioShape> AddDataBar(VisioPage page, VisioShape target, VisioDataGraphicItem item, string value, VisioDataGraphic dataGraphic, int itemIndex) {
            GetItemCenter(target, dataGraphic, itemIndex, item.Width, item.Height, out double x, out double y);
            string baseId = target.Id + "-dg-" + SanitizeId(item.FieldName) + "-" + itemIndex;
            double percent = ResolvePercent(value, item);

            VisioShape background = new VisioShape(UniqueId(page, baseId + "-bar"), x, y, item.Width, item.Height, string.Empty) {
                Name = "Data Graphic Bar",
                NameU = "OfficeIMO Data Graphic Bar"
            };
            item.GetBarBackgroundStyle().ApplyTo(background);
            MarkDataGraphic(background, target, item, value, "BarBackground");
            page.Shapes.Add(background);

            double fillWidth = Math.Max(0.01D, item.Width * percent);
            double fillCenterX = x - (item.Width / 2D) + (fillWidth / 2D);
            VisioShape fill = new VisioShape(UniqueId(page, baseId + "-fill"), fillCenterX, y, fillWidth, item.Height, string.Empty) {
                Name = "Data Graphic Bar Fill",
                NameU = "OfficeIMO Data Graphic Bar Fill"
            };
            item.GetBarFillStyle().ApplyTo(fill);
            MarkDataGraphic(fill, target, item, value, "BarFill");
            fill.SetShapeData("Percent", percent.ToString("0.###", CultureInfo.InvariantCulture), "Percent", VisioShapeDataType.Number);
            page.Shapes.Add(fill);

            VisioShape label = new VisioShape(UniqueId(page, baseId + "-label"), x, y + (item.Height / 2D) + 0.12D, item.Width, 0.22D, item.FormatText(value)) {
                Name = "Data Graphic Label",
                NameU = "Text Box",
                LinePattern = 0,
                FillPattern = 0,
                LineColor = Color.Transparent,
                FillColor = Color.Transparent
            };
            label.TextStyle = item.GetLabelTextStyle().Clone();
            MarkDataGraphic(label, target, item, value, "BarLabel");
            page.Shapes.Add(label);

            if (!string.IsNullOrWhiteSpace(dataGraphic.LayerName)) {
                page.AddToLayer(dataGraphic.LayerName!, background);
                page.AddToLayer(dataGraphic.LayerName!, fill);
                page.AddToLayer(dataGraphic.LayerName!, label);
            }

            return new[] { background, fill, label };
        }

        private static void ValidateDataGraphic(VisioDataGraphic dataGraphic) {
            if (double.IsNaN(dataGraphic.Gap) || double.IsInfinity(dataGraphic.Gap) || dataGraphic.Gap < 0D) {
                throw new ArgumentOutOfRangeException(nameof(dataGraphic.Gap), "Data graphic gap must be a finite non-negative number.");
            }

            if (double.IsNaN(dataGraphic.ItemSpacing) || double.IsInfinity(dataGraphic.ItemSpacing) || dataGraphic.ItemSpacing < 0D) {
                throw new ArgumentOutOfRangeException(nameof(dataGraphic.ItemSpacing), "Data graphic item spacing must be a finite non-negative number.");
            }

            foreach (VisioDataGraphicItem item in dataGraphic.Items) {
                if (double.IsNaN(item.Width) || double.IsInfinity(item.Width) || item.Width <= 0D) {
                    throw new ArgumentOutOfRangeException(nameof(item.Width), "Data graphic item width must be a finite positive number.");
                }

                if (double.IsNaN(item.Height) || double.IsInfinity(item.Height) || item.Height <= 0D) {
                    throw new ArgumentOutOfRangeException(nameof(item.Height), "Data graphic item height must be a finite positive number.");
                }
            }
        }

        private static void MarkDataGraphic(VisioShape shape, VisioShape target, VisioDataGraphicItem item, string value, string role) {
            shape.SetUserCell(VisioSemanticUserCells.Kind, VisioSemanticUserCells.DiagramAdornmentKind, "STR", prompt: "OfficeIMO semantic kind");
            shape.SetUserCell(VisioSemanticUserCells.DataGraphicTargetId, target.Id, "STR", prompt: "OfficeIMO data graphic target shape id");
            shape.SetUserCell(VisioSemanticUserCells.DataGraphicField, item.FieldName, "STR", prompt: "OfficeIMO data graphic field");
            shape.SetUserCell(VisioSemanticUserCells.DataGraphicValue, value, "STR", prompt: "OfficeIMO data graphic value");
            shape.SetUserCell(VisioSemanticUserCells.DataGraphicRole, role, "STR", prompt: "OfficeIMO data graphic role");
            shape.SetShapeData("DataGraphicTargetId", target.Id, "Data graphic target", VisioShapeDataType.String);
            shape.SetShapeData("DataGraphicField", item.FieldName, "Data graphic field", VisioShapeDataType.String);
            shape.SetShapeData("DataGraphicValue", value, "Data graphic value", VisioShapeDataType.String);
            shape.SetShapeData("DataGraphicRole", role, "Data graphic role", VisioShapeDataType.String);
        }

        private static void GetItemCenter(VisioShape target, VisioDataGraphic dataGraphic, int itemIndex, double width, double height, out double x, out double y) {
            double offset = itemIndex * dataGraphic.ItemSpacing;
            switch (dataGraphic.Placement) {
                case VisioDataGraphicPlacement.Left:
                    x = target.PinX - (target.Width / 2D) - dataGraphic.Gap - (width / 2D);
                    y = target.PinY + offset;
                    break;
                case VisioDataGraphicPlacement.Top:
                    x = target.PinX + offset;
                    y = target.PinY + (target.Height / 2D) + dataGraphic.Gap + (height / 2D);
                    break;
                case VisioDataGraphicPlacement.Bottom:
                    x = target.PinX + offset;
                    y = target.PinY - (target.Height / 2D) - dataGraphic.Gap - (height / 2D);
                    break;
                default:
                    x = target.PinX + (target.Width / 2D) + dataGraphic.Gap + (width / 2D);
                    y = target.PinY - offset;
                    break;
            }
        }

        private static double ResolvePercent(string value, VisioDataGraphicItem item) {
            if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double number) &&
                !double.TryParse(value, out number)) {
                return 0D;
            }

            double minimum = item.MinimumValue ?? 0D;
            double maximum = item.MaximumValue ?? 100D;
            if (maximum <= minimum) {
                return 0D;
            }

            return Math.Max(0D, Math.Min(1D, (number - minimum) / (maximum - minimum)));
        }

        private static string UniqueId(VisioPage page, string requestedId) {
            string id = requestedId;
            int index = 2;
            while (page.FindShapeById(id) != null || page.Connectors.Any(connector => string.Equals(connector.Id, id, StringComparison.OrdinalIgnoreCase))) {
                id = requestedId + "-" + index.ToString(CultureInfo.InvariantCulture);
                index++;
            }

            return id;
        }

        private static string SanitizeId(string value) {
            char[] chars = value.Select(ch => char.IsLetterOrDigit(ch) ? ch : '-').ToArray();
            string sanitized = new string(chars).Trim('-');
            return string.IsNullOrEmpty(sanitized) ? "field" : sanitized;
        }
    }

    /// <summary>
    /// Reusable data graphic definition composed of one or more visible Shape Data items.
    /// </summary>
    public sealed class VisioDataGraphic {
        private readonly List<VisioDataGraphicItem> _items = new();

        /// <summary>Creates an empty data graphic definition.</summary>
        public static VisioDataGraphic Create() {
            return new VisioDataGraphic();
        }

        /// <summary>Data graphic items.</summary>
        public IReadOnlyList<VisioDataGraphicItem> Items => _items;

        /// <summary>Placement around each target shape.</summary>
        public VisioDataGraphicPlacement Placement { get; set; } = VisioDataGraphicPlacement.Right;

        /// <summary>Distance between the target and the data graphic.</summary>
        public double Gap { get; set; } = 0.12D;

        /// <summary>Distance between stacked items.</summary>
        public double ItemSpacing { get; set; } = 0.32D;

        /// <summary>Optional layer for generated data graphic shapes.</summary>
        public string? LayerName { get; set; } = "Data Graphics";

        /// <summary>Whether empty Shape Data values should still produce visible graphics.</summary>
        public bool ShowEmptyValues { get; set; }

        /// <summary>Adds a text badge item.</summary>
        public VisioDataGraphic Badge(string fieldName, string? label = null, VisioShapeStyle? style = null) {
            _items.Add(new VisioDataGraphicItem(fieldName, VisioDataGraphicItemKind.TextBadge) {
                Label = label,
                BadgeStyle = style
            });
            return this;
        }

        /// <summary>Adds a data bar item.</summary>
        public VisioDataGraphic Bar(string fieldName, double minimumValue = 0D, double maximumValue = 100D, string? label = null, VisioShapeStyle? fillStyle = null) {
            _items.Add(new VisioDataGraphicItem(fieldName, VisioDataGraphicItemKind.DataBar) {
                Label = label,
                MinimumValue = minimumValue,
                MaximumValue = maximumValue,
                BarFillStyle = fillStyle,
                Width = 1.1D,
                Height = 0.12D
            });
            return this;
        }
    }

    /// <summary>
    /// One visible item in a data graphic definition.
    /// </summary>
    public sealed class VisioDataGraphicItem {
        internal VisioDataGraphicItem(string fieldName, VisioDataGraphicItemKind kind) {
            if (string.IsNullOrWhiteSpace(fieldName)) {
                throw new ArgumentException("Shape Data field name cannot be empty.", nameof(fieldName));
            }

            FieldName = fieldName;
            Kind = kind;
        }

        /// <summary>Shape Data field name to visualize.</summary>
        public string FieldName { get; }

        /// <summary>Visual item kind.</summary>
        public VisioDataGraphicItemKind Kind { get; }

        /// <summary>Optional label prefix. When omitted, the field name is used.</summary>
        public string? Label { get; set; }

        /// <summary>Item width in page units.</summary>
        public double Width { get; set; } = 1.15D;

        /// <summary>Item height in page units.</summary>
        public double Height { get; set; } = 0.28D;

        /// <summary>Minimum value for data bars.</summary>
        public double? MinimumValue { get; set; }

        /// <summary>Maximum value for data bars.</summary>
        public double? MaximumValue { get; set; }

        /// <summary>Optional badge style.</summary>
        public VisioShapeStyle? BadgeStyle { get; set; }

        /// <summary>Optional data bar background style.</summary>
        public VisioShapeStyle? BarBackgroundStyle { get; set; }

        /// <summary>Optional data bar fill style.</summary>
        public VisioShapeStyle? BarFillStyle { get; set; }

        /// <summary>Optional label text style.</summary>
        public VisioTextStyle? LabelTextStyle { get; set; }

        internal string FormatText(string value) {
            string label = string.IsNullOrWhiteSpace(Label) ? FieldName : Label!;
            return string.IsNullOrWhiteSpace(value) ? label : label + ": " + value;
        }

        internal VisioShapeStyle GetBadgeStyle() {
            return BadgeStyle ?? new VisioShapeStyle(Color.FromRgb(232, 244, 255), Color.FromRgb(39, 103, 166), 0.01D) {
                TextStyle = GetLabelTextStyle()
            };
        }

        internal VisioShapeStyle GetBarBackgroundStyle() {
            return BarBackgroundStyle ?? new VisioShapeStyle(Color.FromRgb(232, 235, 239), Color.FromRgb(148, 163, 184), 0.006D);
        }

        internal VisioShapeStyle GetBarFillStyle() {
            return BarFillStyle ?? new VisioShapeStyle(Color.FromRgb(59, 130, 246), Color.FromRgb(37, 99, 235), 0.006D);
        }

        internal VisioTextStyle GetLabelTextStyle() {
            return LabelTextStyle?.Clone() ?? new VisioTextStyle {
                FontFamily = "Aptos",
                Size = 7.5D,
                Color = Color.FromRgb(15, 23, 42),
                Bold = true,
                HorizontalAlignment = VisioTextHorizontalAlignment.Center,
                VerticalAlignment = VisioTextVerticalAlignment.Middle,
                LeftMargin = 0.03D,
                RightMargin = 0.03D,
                TopMargin = 0.01D,
                BottomMargin = 0.01D
            };
        }
    }

    /// <summary>
    /// Visible item kind for Shape Data graphics.
    /// </summary>
    public enum VisioDataGraphicItemKind {
        /// <summary>Compact text badge.</summary>
        TextBadge,
        /// <summary>Numeric data bar.</summary>
        DataBar
    }

    /// <summary>
    /// Placement of generated data graphics around a target shape.
    /// </summary>
    public enum VisioDataGraphicPlacement {
        /// <summary>Place graphics to the right of the target.</summary>
        Right,
        /// <summary>Place graphics to the left of the target.</summary>
        Left,
        /// <summary>Place graphics above the target.</summary>
        Top,
        /// <summary>Place graphics below the target.</summary>
        Bottom
    }
}
