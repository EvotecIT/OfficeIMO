using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace OfficeIMO.Visio.Stencils {
    /// <summary>
    /// Creates OfficeIMO-native stencil catalogs from Visio package master metadata.
    /// </summary>
    public static class VisioStencilPackageCatalog {
        /// <summary>
        /// Loads supported master metadata from a `.vsdx`, `.vssx`, or `.vstx` package into a generated OfficeIMO stencil catalog.
        /// </summary>
        /// <param name="packagePath">Path to a Visio package.</param>
        /// <param name="options">Load options.</param>
        public static VisioStencilCatalog Load(string packagePath, VisioStencilPackageLoadOptions? options = null) {
            if (string.IsNullOrWhiteSpace(packagePath)) throw new ArgumentException("Package path cannot be null or whitespace.", nameof(packagePath));
            if (!File.Exists(packagePath)) throw new FileNotFoundException("Visio package was not found.", packagePath);

            options ??= new VisioStencilPackageLoadOptions();
            if (options.DefaultWidth <= 0) throw new ArgumentOutOfRangeException(nameof(options.DefaultWidth), "Default width must be positive.");
            if (options.DefaultHeight <= 0) throw new ArgumentOutOfRangeException(nameof(options.DefaultHeight), "Default height must be positive.");

            string fileName = Path.GetFileNameWithoutExtension(packagePath);
            string extension = Path.GetExtension(packagePath).TrimStart('.').ToLowerInvariant();
            string catalogName = string.IsNullOrWhiteSpace(options.CatalogName) ? fileName : options.CatalogName!;
            string category = string.IsNullOrWhiteSpace(options.Category) ? catalogName : options.Category!;
            string idPrefix = string.IsNullOrWhiteSpace(options.IdPrefix) ? Slug(fileName) : Slug(options.IdPrefix!);
            HashSet<string>? filter = options.MasterNames != null
                ? new HashSet<string>(options.MasterNames.Where(name => !string.IsNullOrWhiteSpace(name)), StringComparer.OrdinalIgnoreCase)
                : null;
            Dictionary<string, VisioAssets.MasterContent> masterContents = options.LearnMasterDimensions
                ? VisioAssets.LoadMasterContents(packagePath)
                    .GroupBy(master => master.Id, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase)
                : new Dictionary<string, VisioAssets.MasterContent>(StringComparer.OrdinalIgnoreCase);

            VisioStencilCatalogBuilder builder = new(catalogName);
            HashSet<string> usedIds = new(StringComparer.OrdinalIgnoreCase);
            foreach (VisioAssets.MasterInfo master in VisioAssets.ListMasters(packagePath)) {
                if (!VisioMasterIdentity.MatchesAny(master, filter)) {
                    continue;
                }

                bool supported = VisioDocument.IsBuiltinMasterSupported(master.NameU);
                if (!supported && !options.IncludeUnsupportedMasters) {
                    continue;
                }

                string displayName = string.IsNullOrWhiteSpace(master.Name) ? master.NameU : master.Name!;
                string localId = VisioMasterIdentity.ToSlug(master.NameU, "package");
                string id = UniqueId(idPrefix + "." + localId, master.Id, usedIds);
                string[] keywords = new[] { master.NameU, displayName, extension }
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToArray();
                string[] aliases = new[] { master.Id, master.RelationshipId, localId, displayName.Replace(" ", "-") }
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToArray();
                string[] tags = new[] { "package", extension, supported ? "supported" : "generic", category }
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToArray();
                double defaultWidth = options.DefaultWidth;
                double defaultHeight = options.DefaultHeight;
                VisioMeasurementUnit? defaultUnit = null;
                if (masterContents.TryGetValue(master.Id, out VisioAssets.MasterContent? content) &&
                    TryReadMasterDimensions(content, out double masterWidth, out double masterHeight)) {
                    defaultWidth = masterWidth;
                    defaultHeight = masterHeight;
                    defaultUnit = VisioMeasurementUnit.Inches;
                }

                builder.AddWithMetadata(
                    id,
                    displayName,
                    master.NameU,
                    category,
                    defaultWidth,
                    defaultHeight,
                    keywords,
                    aliases,
                    tags,
                    master.NameU,
                    defaultUnit);
            }

            return builder.Build();
        }

        private static bool TryReadMasterDimensions(VisioAssets.MasterContent content, out double width, out double height) {
            XNamespace v = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement? shape = content.MasterXml.Root?
                .Element(v + "Shapes")?
                .Elements(v + "Shape")
                .FirstOrDefault();
            if (shape == null) {
                shape = content.MasterXml.Root?
                    .Elements()
                    .FirstOrDefault(element => string.Equals(element.Name.LocalName, "Shapes", StringComparison.OrdinalIgnoreCase))?
                    .Elements()
                    .FirstOrDefault(element => string.Equals(element.Name.LocalName, "Shape", StringComparison.OrdinalIgnoreCase));
            }

            if (shape != null &&
                TryReadPositiveCell(shape, "Width", out width) &&
                TryReadPositiveCell(shape, "Height", out height)) {
                return true;
            }

            width = 0;
            height = 0;
            return false;
        }

        private static bool TryReadPositiveCell(XElement shape, string name, out double value) {
            XElement? cell = shape.Elements()
                .FirstOrDefault(element =>
                    string.Equals(element.Name.LocalName, "Cell", StringComparison.OrdinalIgnoreCase) &&
                    string.Equals((string?)element.Attribute("N"), name, StringComparison.OrdinalIgnoreCase));
            string? rawValue = (string?)cell?.Attribute("V");
            if (string.IsNullOrWhiteSpace(rawValue) ||
                !double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out value) ||
                value <= 0 ||
                double.IsNaN(value) ||
                double.IsInfinity(value)) {
                value = 0;
                return false;
            }

            return true;
        }

        private static string UniqueId(string baseId, string fallback, HashSet<string> usedIds) {
            string id = baseId;
            if (usedIds.Add(id)) {
                return id;
            }

            string suffix = VisioMasterIdentity.ToSlug(fallback);
            id = string.IsNullOrWhiteSpace(suffix) ? baseId + "-2" : baseId + "-" + suffix;
            int counter = 2;
            while (!usedIds.Add(id)) {
                id = baseId + "-" + counter.ToString(System.Globalization.CultureInfo.InvariantCulture);
                counter++;
            }

            return id;
        }

        private static string Slug(string value) => VisioMasterIdentity.ToSlug(value, "package");
    }
}
