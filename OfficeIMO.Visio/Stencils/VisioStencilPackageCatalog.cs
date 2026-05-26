using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

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

            VisioStencilCatalogBuilder builder = new(catalogName);
            HashSet<string> usedIds = new(StringComparer.OrdinalIgnoreCase);
            foreach (VisioAssets.MasterInfo master in VisioAssets.ListMasters(packagePath)) {
                if (filter != null && !filter.Contains(master.NameU)) {
                    continue;
                }

                bool supported = VisioDocument.IsBuiltinMasterSupported(master.NameU);
                if (!supported && !options.IncludeUnsupportedMasters) {
                    continue;
                }

                string displayName = string.IsNullOrWhiteSpace(master.Name) ? master.NameU : master.Name!;
                string localId = Slug(master.NameU);
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

                builder.AddWithMetadata(
                    id,
                    displayName,
                    master.NameU,
                    category,
                    options.DefaultWidth,
                    options.DefaultHeight,
                    keywords,
                    aliases,
                    tags,
                    master.NameU);
            }

            return builder.Build();
        }

        private static string UniqueId(string baseId, string fallback, HashSet<string> usedIds) {
            string id = baseId;
            if (usedIds.Add(id)) {
                return id;
            }

            string suffix = Slug(fallback);
            id = string.IsNullOrWhiteSpace(suffix) ? baseId + "-2" : baseId + "-" + suffix;
            int counter = 2;
            while (!usedIds.Add(id)) {
                id = baseId + "-" + counter.ToString(System.Globalization.CultureInfo.InvariantCulture);
                counter++;
            }

            return id;
        }

        private static string Slug(string value) {
            StringBuilder builder = new(value.Length);
            bool previousDash = false;
            foreach (char character in value.Trim()) {
                if (char.IsLetterOrDigit(character)) {
                    builder.Append(char.ToLowerInvariant(character));
                    previousDash = false;
                } else if (!previousDash) {
                    builder.Append('-');
                    previousDash = true;
                }
            }

            string slug = builder.ToString().Trim('-');
            return string.IsNullOrWhiteSpace(slug) ? "package" : slug;
        }
    }
}
