using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio.Stencils {
    /// <summary>
    /// Creates complete Visio stencil gallery review documents from OfficeIMO stencil catalogs.
    /// </summary>
    public static class VisioStencilGalleryDocument {
        /// <summary>
        /// Creates a stencil gallery document at the supplied path.
        /// </summary>
        public static VisioDocument Create(string path, VisioStencilCatalog catalog, VisioStencilGalleryDocumentOptions? options = null) {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Document path cannot be null or whitespace.", nameof(path));
            VisioDocument document = VisioDocument.Create(path);
            return Populate(document, catalog, options);
        }

        /// <summary>
        /// Creates a stencil gallery document in the supplied stream.
        /// </summary>
        public static VisioDocument Create(Stream stream, VisioStencilCatalog catalog, VisioStencilGalleryDocumentOptions? options = null) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            VisioDocument document = VisioDocument.Create(stream);
            return Populate(document, catalog, options);
        }

        /// <summary>
        /// Adds stencil gallery review pages to an existing document.
        /// </summary>
        public static VisioDocument AddStencilGalleryDocument(this VisioDocument document, VisioStencilCatalog catalog, VisioStencilGalleryDocumentOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            bool originalUseMastersByDefault = document.UseMastersByDefault;
            try {
                return Populate(document, catalog, options);
            } finally {
                document.UseMastersByDefault = originalUseMastersByDefault;
            }
        }

        private static VisioDocument Populate(VisioDocument document, VisioStencilCatalog catalog, VisioStencilGalleryDocumentOptions? options) {
            if (catalog == null) throw new ArgumentNullException(nameof(catalog));
            VisioStencilGalleryDocumentOptions effectiveOptions = options ?? new VisioStencilGalleryDocumentOptions();
            ValidateOptions(effectiveOptions);

            string title = string.IsNullOrWhiteSpace(effectiveOptions.Title)
                ? catalog.Name + " Stencil Gallery"
                : effectiveOptions.Title!;
            document.Title = title;
            if (!string.IsNullOrWhiteSpace(effectiveOptions.Author)) {
                document.Author = effectiveOptions.Author;
            }

            document.UseMastersByDefault = effectiveOptions.UseMastersByDefault;

            if (effectiveOptions.IncludeOverviewPage) {
                AddOverviewPage(document, catalog, effectiveOptions, title);
            }

            IReadOnlyList<GalleryPageGroup> groups = CreateGroups(catalog, effectiveOptions);
            int pageIndex = 1;
            foreach (GalleryPageGroup group in groups) {
                string pageName = MakePageName(group.Title, pageIndex);
                VisioPage page = document.AddPage(pageName, effectiveOptions.PageWidth, effectiveOptions.PageHeight, effectiveOptions.PageUnit);
                VisioStencilCatalog pageCatalog = new(group.CatalogName, group.Shapes);
                page.AddStencilGallery(pageCatalog, new VisioStencilGalleryOptions {
                    Title = group.Title,
                    IdPrefix = MakeIdPrefix(effectiveOptions.IdPrefix, pageIndex),
                    MaxShapes = group.Shapes.Count,
                    Columns = effectiveOptions.Columns,
                    AutoResizePage = effectiveOptions.AutoResizePages,
                    ShowCategory = effectiveOptions.ShowCategory,
                    IncludeStencilMetadataShapeData = effectiveOptions.IncludeStencilMetadataShapeData,
                    CellFillColor = effectiveOptions.CellFillColor,
                    CellBorderColor = effectiveOptions.CellBorderColor,
                    TitleColor = OfficeColor.FromRgb(31, 48, 63)
                });
                AddCatalogFooter(page, catalog, group, effectiveOptions, pageIndex, groups.Count);
                pageIndex++;
            }

            return document;
        }

        private static void AddOverviewPage(VisioDocument document, VisioStencilCatalog catalog, VisioStencilGalleryDocumentOptions options, string title) {
            VisioPage page = document.AddPage("Stencil Gallery Overview", options.PageWidth, options.PageHeight, options.PageUnit);
            HashSet<string> usedIds = new(StringComparer.OrdinalIgnoreCase);
            double contentWidth = Math.Max(1D, page.Width - 1.2D);
            double centerX = page.Width / 2D;
            double headingY = Math.Max(0.8D, page.Height - 0.75D);
            VisioShape heading = page.AddTextBox(ReserveId(usedIds, MakeId(options.IdPrefix, "overview-title")), centerX, headingY, contentWidth, 0.55, title, VisioMeasurementUnit.Inches);
            VisioSemanticUserCells.MarkGeneratedAdornment(heading);
            heading.TextStyle = new VisioTextStyle {
                FontFamily = "Aptos Display",
                Size = 22D,
                Bold = true,
                Color = OfficeColor.FromRgb(27, 43, 57),
                HorizontalAlignment = VisioTextHorizontalAlignment.Center,
                VerticalAlignment = VisioTextVerticalAlignment.Middle
            };

            string summary = catalog.Shapes.Count.ToString(CultureInfo.InvariantCulture) +
                             " stencils across " +
                             catalog.Categories.Count.ToString(CultureInfo.InvariantCulture) +
                             " categories";
            VisioShape subtitle = page.AddTextBox(ReserveId(usedIds, MakeId(options.IdPrefix, "overview-summary")), centerX, headingY - 0.55D, contentWidth, 0.35, summary, VisioMeasurementUnit.Inches);
            VisioSemanticUserCells.MarkGeneratedAdornment(subtitle);
            subtitle.TextStyle = new VisioTextStyle {
                FontFamily = "Aptos",
                Size = 10.5D,
                Color = OfficeColor.FromRgb(78, 93, 108),
                HorizontalAlignment = VisioTextHorizontalAlignment.Center,
                VerticalAlignment = VisioTextVerticalAlignment.Middle
            };

            double metricGap = 0.25D;
            double metricWidth = Math.Min(2.5D, Math.Max(1D, (contentWidth - (metricGap * 2D)) / 3D));
            double metricTotalWidth = (metricWidth * 3D) + (metricGap * 2D);
            double metricStartX = centerX - (metricTotalWidth / 2D) + (metricWidth / 2D);
            double metricY = Math.Max(1.8D, headingY - 1.45D);
            AddMetric(page, ReserveId(usedIds, MakeId(options.IdPrefix, "overview-count")), metricStartX, metricY, metricWidth, "Shapes", catalog.Shapes.Count.ToString(CultureInfo.InvariantCulture), options.AccentColor);
            AddMetric(page, ReserveId(usedIds, MakeId(options.IdPrefix, "overview-categories")), metricStartX + metricWidth + metricGap, metricY, metricWidth, "Categories", catalog.Categories.Count.ToString(CultureInfo.InvariantCulture), OfficeColor.FromRgb(46, 160, 67));
            AddMetric(page, ReserveId(usedIds, MakeId(options.IdPrefix, "overview-packages")), metricStartX + ((metricWidth + metricGap) * 2D), metricY, metricWidth, "Source packs", CountSourcePackages(catalog).ToString(CultureInfo.InvariantCulture), OfficeColor.FromRgb(137, 87, 229));

            double y = metricY - 1.05D;
            foreach (IGrouping<string, VisioStencilShape> group in catalog.Shapes
                .GroupBy(shape => shape.Category, StringComparer.OrdinalIgnoreCase)
                .OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase)
                .Take(10)) {
                VisioShape row = page.AddTextBox(
                    ReserveId(usedIds, MakeId(options.IdPrefix, "overview-category-" + Slug(group.Key))),
                    centerX,
                    y,
                    contentWidth,
                    0.32,
                    group.Key + " - " + group.Count().ToString(CultureInfo.InvariantCulture) + " stencils",
                    VisioMeasurementUnit.Inches);
                VisioSemanticUserCells.MarkGeneratedAdornment(row);
                row.TextStyle = new VisioTextStyle {
                    FontFamily = "Aptos",
                    Size = 9.4D,
                    Color = OfficeColor.FromRgb(41, 54, 68),
                    HorizontalAlignment = VisioTextHorizontalAlignment.Left,
                    VerticalAlignment = VisioTextVerticalAlignment.Middle
                };
                row.SetShapeData("Category", group.Key, "Category", VisioShapeDataType.String);
                row.SetShapeData("StencilCount", group.Count().ToString(CultureInfo.InvariantCulture), "Stencil count", VisioShapeDataType.Number);
                y -= 0.38D;
            }
        }

        private static void AddMetric(VisioPage page, string id, double x, double y, double width, string label, string value, OfficeColor accent) {
            VisioShape shape = new(id, x, y, width, 0.82D, value + Environment.NewLine + label) {
                NameU = "Rectangle",
                FillColor = OfficeColor.FromRgb(248, 251, 254),
                LineColor = accent,
                LineWeight = 0.018D
            };
            shape.TextStyle = new VisioTextStyle {
                FontFamily = "Aptos",
                Size = 10D,
                Bold = true,
                Color = OfficeColor.FromRgb(27, 43, 57),
                HorizontalAlignment = VisioTextHorizontalAlignment.Center,
                VerticalAlignment = VisioTextVerticalAlignment.Middle
            };
            page.Shapes.Add(shape);
        }

        private static void AddCatalogFooter(VisioPage page, VisioStencilCatalog catalog, GalleryPageGroup group, VisioStencilGalleryDocumentOptions options, int pageIndex, int pageCount) {
            string footerText = catalog.Name +
                                " | " +
                                group.Shapes.Count.ToString(CultureInfo.InvariantCulture) +
                                " shown | Page " +
                                pageIndex.ToString(CultureInfo.InvariantCulture) +
                                " of " +
                                pageCount.ToString(CultureInfo.InvariantCulture);
            VisioShape footer = page.AddTextBox(MakeId(options.IdPrefix, "footer-" + pageIndex.ToString(CultureInfo.InvariantCulture)), page.Width / 2D, 0.28D, page.Width - 1D, 0.22D, footerText, VisioMeasurementUnit.Inches);
            VisioSemanticUserCells.MarkGeneratedAdornment(footer);
            footer.TextStyle = new VisioTextStyle {
                FontFamily = "Aptos",
                Size = 7.5D,
                Color = OfficeColor.FromRgb(93, 107, 122),
                HorizontalAlignment = VisioTextHorizontalAlignment.Center,
                VerticalAlignment = VisioTextVerticalAlignment.Middle
            };
        }

        private static IReadOnlyList<GalleryPageGroup> CreateGroups(VisioStencilCatalog catalog, VisioStencilGalleryDocumentOptions options) {
            IEnumerable<IGrouping<string, VisioStencilShape>> categoryGroups = options.GroupByCategory
                ? catalog.Shapes.GroupBy(shape => shape.Category, StringComparer.OrdinalIgnoreCase)
                : new[] { new GalleryGrouping(catalog.Name, catalog.Shapes) };

            List<GalleryPageGroup> groups = new();
            foreach (IGrouping<string, VisioStencilShape> category in categoryGroups.OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase)) {
                VisioStencilShape[] ordered = category
                    .OrderBy(shape => shape.Name, StringComparer.OrdinalIgnoreCase)
                    .ToArray();
                for (int offset = 0; offset < ordered.Length; offset += options.ShapesPerPage) {
                    VisioStencilShape[] chunk = ordered.Skip(offset).Take(options.ShapesPerPage).ToArray();
                    int part = (offset / options.ShapesPerPage) + 1;
                    int totalParts = (int)Math.Ceiling(ordered.Length / (double)options.ShapesPerPage);
                    string title = category.Key;
                    if (totalParts > 1) {
                        title += " (" + part.ToString(CultureInfo.InvariantCulture) + " of " + totalParts.ToString(CultureInfo.InvariantCulture) + ")";
                    }

                    groups.Add(new GalleryPageGroup(title, catalog.Name + " - " + category.Key, chunk));
                }
            }

            return groups.AsReadOnly();
        }

        private static int CountSourcePackages(VisioStencilCatalog catalog) {
            return catalog.Shapes
                .Select(shape => shape.SourcePackagePath)
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Count();
        }

        private static string MakePageName(string value, int pageIndex) {
            string name = string.IsNullOrWhiteSpace(value) ? "Stencil Gallery" : value.Trim();
            if (name.Length > 28) {
                name = name.Substring(0, 28).Trim();
            }

            return pageIndex.ToString("00", CultureInfo.InvariantCulture) + " " + name;
        }

        private static string MakeIdPrefix(string prefix, int pageIndex) {
            return MakeId(prefix, pageIndex.ToString("00", CultureInfo.InvariantCulture));
        }

        private static string MakeId(string prefix, string suffix) {
            return Slug(prefix) + "-" + Slug(suffix);
        }

        private static string ReserveId(HashSet<string> usedIds, string baseId) {
            string id = baseId;
            int suffix = 2;
            while (!usedIds.Add(id)) {
                id = baseId + "-" + suffix.ToString(CultureInfo.InvariantCulture);
                suffix++;
            }

            return id;
        }

        private static string Slug(string value) {
            string normalized = string.IsNullOrWhiteSpace(value) ? "gallery" : value.Trim();
            string slug = string.Concat(normalized.Select(ch => char.IsLetterOrDigit(ch) || ch == '-' || ch == '_' ? char.ToLowerInvariant(ch) : '-'));
            while (slug.Contains("--")) {
                slug = slug.Replace("--", "-");
            }

            return slug.Trim('-');
        }

        private static void ValidateOptions(VisioStencilGalleryDocumentOptions options) {
            if (string.IsNullOrWhiteSpace(options.IdPrefix)) throw new ArgumentException("Gallery document id prefix cannot be null or whitespace.", nameof(options));
            if (options.ShapesPerPage <= 0) throw new ArgumentOutOfRangeException(nameof(options.ShapesPerPage), "Shapes per page must be positive.");
            if (options.Columns <= 0) throw new ArgumentOutOfRangeException(nameof(options.Columns), "Column count must be positive.");
            ValidatePositive(options.PageWidth, nameof(options.PageWidth));
            ValidatePositive(options.PageHeight, nameof(options.PageHeight));
        }

        private static void ValidatePositive(double value, string name) {
            if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0D) {
                throw new ArgumentOutOfRangeException(name, "Value must be a finite positive number.");
            }
        }

        private sealed class GalleryPageGroup {
            public GalleryPageGroup(string title, string catalogName, IReadOnlyList<VisioStencilShape> shapes) {
                Title = title;
                CatalogName = catalogName;
                Shapes = shapes;
            }

            public string Title { get; }

            public string CatalogName { get; }

            public IReadOnlyList<VisioStencilShape> Shapes { get; }
        }

        private sealed class GalleryGrouping : IGrouping<string, VisioStencilShape> {
            private readonly IReadOnlyList<VisioStencilShape> _shapes;

            public GalleryGrouping(string key, IReadOnlyList<VisioStencilShape> shapes) {
                Key = key;
                _shapes = shapes;
            }

            public string Key { get; }

            public IEnumerator<VisioStencilShape> GetEnumerator() {
                return _shapes.GetEnumerator();
            }

            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() {
                return GetEnumerator();
            }
        }
    }
}
