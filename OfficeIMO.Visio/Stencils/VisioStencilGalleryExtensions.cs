using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio.Stencils {
    /// <summary>
    /// Helpers for rendering stencil catalogs as browsable Visio pages.
    /// </summary>
    public static class VisioStencilGalleryExtensions {
        /// <summary>
        /// Adds a contact-sheet gallery for a stencil catalog to the page.
        /// </summary>
        /// <param name="page">Target page.</param>
        /// <param name="catalog">Stencil catalog to render.</param>
        /// <param name="options">Optional gallery layout and visual options.</param>
        /// <returns>The stencil instance shapes placed in the gallery.</returns>
        public static IReadOnlyList<VisioShape> AddStencilGallery(this VisioPage page, VisioStencilCatalog catalog, VisioStencilGalleryOptions? options = null) {
            if (page == null) throw new ArgumentNullException(nameof(page));
            if (catalog == null) throw new ArgumentNullException(nameof(catalog));

            VisioStencilGalleryOptions effectiveOptions = options ?? new VisioStencilGalleryOptions();
            ValidateOptions(effectiveOptions);

            List<VisioStencilShape> stencils = catalog.Shapes
                .Take(effectiveOptions.MaxShapes)
                .ToList();
            if (stencils.Count == 0) {
                return Array.Empty<VisioShape>();
            }

            HashSet<string> reservedIds = new(StringComparer.Ordinal);
            ReserveExistingShapeIds(page.Shapes, reservedIds);
            foreach (VisioConnector connector in page.Connectors) {
                reservedIds.Add(connector.Id);
            }

            int columns = Math.Min(effectiveOptions.Columns, stencils.Count);
            int rows = (int)Math.Ceiling(stencils.Count / (double)columns);
            double titleHeight = 0.46D;
            double titleGap = 0.28D;
            double requiredWidth = effectiveOptions.Left * 2D
                                   + (columns * effectiveOptions.CellWidth)
                                   + Math.Max(0, columns - 1) * effectiveOptions.ColumnGap;
            double requiredHeight = effectiveOptions.Top * 2D
                                    + titleHeight
                                    + titleGap
                                    + (rows * effectiveOptions.CellHeight)
                                    + Math.Max(0, rows - 1) * effectiveOptions.RowGap;
            if (effectiveOptions.AutoResizePage) {
                page.Width = Math.Max(page.Width, requiredWidth);
                page.Height = Math.Max(page.Height, requiredHeight);
            }

            string title = string.IsNullOrWhiteSpace(effectiveOptions.Title) ? catalog.Name : effectiveOptions.Title!;
            VisioShape titleShape = page.AddTextBox(
                ReserveId(reservedIds, effectiveOptions.IdPrefix, "title"),
                page.Width / 2D,
                page.Height - effectiveOptions.Top - (titleHeight / 2D),
                Math.Max(1D, page.Width - (effectiveOptions.Left * 2D)),
                titleHeight,
                title,
                VisioMeasurementUnit.Inches);
            titleShape.TextStyle = new VisioTextStyle {
                FontFamily = "Aptos Display",
                Size = 18D,
                Bold = true,
                Color = effectiveOptions.TitleColor,
                HorizontalAlignment = VisioTextHorizontalAlignment.Center,
                VerticalAlignment = VisioTextVerticalAlignment.Middle
            };

            double gridTop = page.Height - effectiveOptions.Top - titleHeight - titleGap;
            List<VisioShape> placed = new();
            for (int i = 0; i < stencils.Count; i++) {
                VisioStencilShape stencil = stencils[i];
                int row = i / columns;
                int column = i % columns;
                double cellLeft = effectiveOptions.Left + column * (effectiveOptions.CellWidth + effectiveOptions.ColumnGap);
                double cellTop = gridTop - row * (effectiveOptions.CellHeight + effectiveOptions.RowGap);
                double centerX = cellLeft + (effectiveOptions.CellWidth / 2D);
                double centerY = cellTop - (effectiveOptions.CellHeight / 2D);
                string itemPrefix = SafeId(effectiveOptions.IdPrefix, i.ToString(CultureInfo.InvariantCulture));

                VisioShape cell = new VisioShape(
                    ReserveId(reservedIds, itemPrefix, "cell"),
                    centerX,
                    centerY,
                    effectiveOptions.CellWidth,
                    effectiveOptions.CellHeight,
                    string.Empty) {
                    NameU = "Rectangle"
                };
                cell.FillColor = effectiveOptions.CellFillColor;
                cell.LineColor = effectiveOptions.CellBorderColor;
                cell.LineWeight = 0.01D;
                page.Shapes.Add(cell);

                FitStencil(stencil, effectiveOptions, out double iconWidth, out double iconHeight);
                double iconY = cellTop - 0.44D;
                VisioShape icon = page.AddStencilShape(
                    stencil,
                    ReserveId(reservedIds, itemPrefix, "shape"),
                    centerX,
                    iconY,
                    iconWidth,
                    iconHeight,
                    string.Empty,
                    VisioMeasurementUnit.Inches);
                if (effectiveOptions.IncludeStencilMetadataShapeData) {
                    ApplyGalleryShapeData(icon, catalog, stencil, i);
                }

                placed.Add(icon);

                AddLabel(page, ReserveId(reservedIds, itemPrefix, "name"), centerX, cellTop - 1.05D, effectiveOptions.CellWidth - 0.22D, 0.3D, stencil.Name, 9.2D, true, OfficeColor.FromRgb(28, 38, 48));
                if (effectiveOptions.ShowCategory) {
                    AddLabel(page, ReserveId(reservedIds, itemPrefix, "category"), centerX, cellTop - 1.27D, effectiveOptions.CellWidth - 0.22D, 0.22D, stencil.Category, 7.5D, false, OfficeColor.FromRgb(88, 102, 116));
                }
            }

            return placed.AsReadOnly();
        }

        private static void ApplyGalleryShapeData(VisioShape shape, VisioStencilCatalog catalog, VisioStencilShape stencil, int index) {
            shape.SetShapeData("GalleryIndex", index.ToString(CultureInfo.InvariantCulture), "Gallery index", VisioShapeDataType.Number);
            shape.SetShapeData("StencilId", stencil.Id, "Stencil id", VisioShapeDataType.String);
            shape.SetShapeData("StencilName", stencil.Name, "Stencil name", VisioShapeDataType.String);
            shape.SetShapeData("StencilCategory", stencil.Category, "Stencil category", VisioShapeDataType.String);
            shape.SetShapeData("StencilCatalog", catalog.Name, "Stencil catalog", VisioShapeDataType.String);
            shape.SetShapeData("MasterNameU", stencil.MasterNameU, "Master NameU", VisioShapeDataType.String);
            shape.SetShapeData("IconNameU", stencil.IconNameU, "Icon NameU", VisioShapeDataType.String);
            shape.SetShapeData("DefaultSize", FormatDefaultSize(stencil), "Default size", VisioShapeDataType.String);
            if (stencil.Keywords.Count > 0) {
                shape.SetShapeData("Keywords", string.Join("; ", stencil.Keywords), "Keywords", VisioShapeDataType.String);
            }

            if (stencil.Aliases.Count > 0) {
                shape.SetShapeData("Aliases", string.Join("; ", stencil.Aliases), "Aliases", VisioShapeDataType.String);
            }

            if (stencil.Tags.Count > 0) {
                shape.SetShapeData("Tags", string.Join("; ", stencil.Tags), "Tags", VisioShapeDataType.String);
            }

            if (!string.IsNullOrWhiteSpace(stencil.SourcePackagePath)) {
                shape.SetShapeData("SourcePackagePath", stencil.SourcePackagePath, "Source package path", VisioShapeDataType.String);
            }

            if (stencil.PreviewImage != null) {
                string? preview = string.IsNullOrWhiteSpace(stencil.PreviewImage.ContentType)
                    ? stencil.PreviewImage.Extension
                    : stencil.PreviewImage.ContentType;
                if (!string.IsNullOrWhiteSpace(preview)) {
                    shape.SetShapeData("PreviewImage", preview, "Preview image", VisioShapeDataType.String);
                }
            }

            if (stencil.SourceConnectionPoints.Count > 0) {
                shape.SetShapeData("SourceConnectionPoints", stencil.SourceConnectionPoints.Count.ToString(CultureInfo.InvariantCulture), "Source connection points", VisioShapeDataType.Number);
            }
        }

        private static string FormatDefaultSize(VisioStencilShape stencil) {
            string unit = stencil.DefaultUnit?.ToString() ?? "PageUnit";
            return stencil.DefaultWidth.ToString("0.###", CultureInfo.InvariantCulture) +
                   " x " +
                   stencil.DefaultHeight.ToString("0.###", CultureInfo.InvariantCulture) +
                   " " +
                   unit;
        }

        private static void AddLabel(VisioPage page, string id, double x, double y, double width, double height, string text, double size, bool bold, OfficeColor color) {
            VisioShape label = page.AddTextBox(id, x, y, width, height, text, VisioMeasurementUnit.Inches);
            label.TextStyle = new VisioTextStyle {
                FontFamily = "Aptos",
                Size = size,
                Bold = bold,
                Color = color,
                BackgroundColor = OfficeColor.White,
                BackgroundTransparency = 0,
                HorizontalAlignment = VisioTextHorizontalAlignment.Center,
                VerticalAlignment = VisioTextVerticalAlignment.Middle
            };
        }

        private static void FitStencil(VisioStencilShape stencil, VisioStencilGalleryOptions options, out double width, out double height) {
            VisioMeasurementUnit sourceUnit = stencil.DefaultUnit ?? VisioMeasurementUnit.Inches;
            width = stencil.DefaultWidth.ToInches(sourceUnit);
            height = stencil.DefaultHeight.ToInches(sourceUnit);
            double scale = Math.Min(options.IconMaxWidth / width, options.IconMaxHeight / height);
            if (scale < 1D) {
                width *= scale;
                height *= scale;
            }
        }

        private static string SafeId(string prefix, string suffix) {
            string value = prefix + "-" + suffix;
            return string.Concat(value.Select(ch => char.IsLetterOrDigit(ch) || ch == '-' || ch == '_' ? ch : '-'));
        }

        private static string ReserveId(HashSet<string> reservedIds, string prefix, string suffix) {
            string baseId = SafeId(prefix, suffix);
            string id = baseId;
            int index = 2;
            while (reservedIds.Contains(id)) {
                id = baseId + "-" + index.ToString(CultureInfo.InvariantCulture);
                index++;
            }

            reservedIds.Add(id);
            return id;
        }

        private static void ReserveExistingShapeIds(IEnumerable<VisioShape> shapes, HashSet<string> reservedIds) {
            foreach (VisioShape shape in shapes) {
                reservedIds.Add(shape.Id);
                if (shape.Children.Count > 0) {
                    ReserveExistingShapeIds(shape.Children, reservedIds);
                }
            }
        }

        private static void ValidateOptions(VisioStencilGalleryOptions options) {
            if (string.IsNullOrWhiteSpace(options.IdPrefix)) throw new ArgumentException("Gallery id prefix cannot be null or whitespace.", nameof(options));
            if (options.MaxShapes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxShapes), "Maximum gallery shape count must be positive.");
            if (options.Columns <= 0) throw new ArgumentOutOfRangeException(nameof(options.Columns), "Gallery column count must be positive.");
            ValidatePositive(options.Left, nameof(options.Left));
            ValidatePositive(options.Top, nameof(options.Top));
            ValidateNonNegative(options.ColumnGap, nameof(options.ColumnGap));
            ValidateNonNegative(options.RowGap, nameof(options.RowGap));
            ValidatePositive(options.CellWidth, nameof(options.CellWidth));
            ValidatePositive(options.CellHeight, nameof(options.CellHeight));
            ValidatePositive(options.IconMaxWidth, nameof(options.IconMaxWidth));
            ValidatePositive(options.IconMaxHeight, nameof(options.IconMaxHeight));
        }

        private static void ValidatePositive(double value, string name) {
            if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0D) {
                throw new ArgumentOutOfRangeException(name, "Value must be a finite positive number.");
            }
        }

        private static void ValidateNonNegative(double value, string name) {
            if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D) {
                throw new ArgumentOutOfRangeException(name, "Value must be a finite non-negative number.");
            }
        }
    }
}
