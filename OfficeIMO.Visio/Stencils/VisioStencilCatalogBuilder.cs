using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio.Stencils {
    /// <summary>
    /// Fluent builder for OfficeIMO-native stencil catalogs.
    /// </summary>
    public sealed class VisioStencilCatalogBuilder {
        private readonly List<VisioStencilShape> _shapes = new();
        private readonly HashSet<string> _ids = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Initializes a new stencil catalog builder.
        /// </summary>
        /// <param name="name">Catalog name.</param>
        public VisioStencilCatalogBuilder(string name) {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Catalog name cannot be null or whitespace.", nameof(name));
            Name = name;
        }

        /// <summary>
        /// Gets the catalog name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Gets shapes added to the builder.
        /// </summary>
        public IReadOnlyList<VisioStencilShape> Shapes => _shapes.AsReadOnly();

        /// <summary>
        /// Adds a stencil shape and derives common aliases and tags from the id, category, master, name, and keywords.
        /// </summary>
        public VisioStencilCatalogBuilder Add(
            string id,
            string name,
            string masterNameU,
            string category,
            double defaultWidth,
            double defaultHeight,
            params string[] keywords) {
            return AddWithMetadata(id, name, masterNameU, category, defaultWidth, defaultHeight, keywords);
        }

        /// <summary>
        /// Adds a stencil shape with explicit search metadata.
        /// </summary>
        public VisioStencilCatalogBuilder AddWithMetadata(
            string id,
            string name,
            string masterNameU,
            string category,
            double defaultWidth,
            double defaultHeight,
            IEnumerable<string>? keywords = null,
            IEnumerable<string>? aliases = null,
            IEnumerable<string>? tags = null,
            string? iconNameU = null) {
            return Add(CreateShape(id, name, masterNameU, category, defaultWidth, defaultHeight, keywords, aliases, tags, iconNameU));
        }

        /// <summary>
        /// Adds an existing stencil shape definition.
        /// </summary>
        public VisioStencilCatalogBuilder Add(VisioStencilShape shape) {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (!_ids.Add(shape.Id)) {
                throw new ArgumentException($"Stencil shape id '{shape.Id}' already exists in catalog '{Name}'.", nameof(shape));
            }

            _shapes.Add(shape);
            return this;
        }

        /// <summary>
        /// Builds the catalog.
        /// </summary>
        public VisioStencilCatalog Build() {
            return new VisioStencilCatalog(Name, _shapes);
        }

        private static VisioStencilShape CreateShape(
            string id,
            string name,
            string masterNameU,
            string category,
            double defaultWidth,
            double defaultHeight,
            IEnumerable<string>? keywords,
            IEnumerable<string>? aliases,
            IEnumerable<string>? tags,
            string? iconNameU) {
            string prefix = id.Contains(".") ? id.Substring(0, id.IndexOf('.')) : id;
            string localId = id.Contains(".") ? id.Substring(id.IndexOf('.') + 1) : id;
            IEnumerable<string> effectiveKeywords = keywords ?? Enumerable.Empty<string>();
            string[] effectiveAliases = effectiveKeywords
                .Concat(new[] { localId, name.Replace(" ", "-") })
                .Concat(aliases ?? Enumerable.Empty<string>())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
            string[] effectiveTags = new[] { prefix, category, masterNameU }
                .Concat(tags ?? Enumerable.Empty<string>())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();

            return new VisioStencilShape(
                id,
                name,
                masterNameU,
                category,
                defaultWidth,
                defaultHeight,
                effectiveKeywords,
                effectiveAliases,
                effectiveTags,
                iconNameU ?? masterNameU);
        }
    }
}
