using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio.Stencils {
    /// <summary>
    /// Describes an OfficeIMO-native stencil shape that can be generated without depending on a VSDX/VSSX template.
    /// </summary>
    public sealed class VisioStencilShape {
        /// <summary>
        /// Initializes a new stencil shape definition.
        /// </summary>
        public VisioStencilShape(string id, string name, string masterNameU, string category, double defaultWidth, double defaultHeight, IEnumerable<string>? keywords = null, IEnumerable<string>? aliases = null, IEnumerable<string>? tags = null, string? iconNameU = null)
            : this(id, name, masterNameU, category, defaultWidth, defaultHeight, keywords, aliases, tags, iconNameU, null) {
        }

        /// <summary>
        /// Initializes a new stencil shape definition with an explicit default-size unit.
        /// </summary>
        public VisioStencilShape(string id, string name, string masterNameU, string category, double defaultWidth, double defaultHeight, IEnumerable<string>? keywords, IEnumerable<string>? aliases, IEnumerable<string>? tags, string? iconNameU, VisioMeasurementUnit? defaultUnit) {
            if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Stencil shape id cannot be null or whitespace.", nameof(id));
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Stencil shape name cannot be null or whitespace.", nameof(name));
            if (string.IsNullOrWhiteSpace(masterNameU)) throw new ArgumentException("Master NameU cannot be null or whitespace.", nameof(masterNameU));
            if (string.IsNullOrWhiteSpace(category)) throw new ArgumentException("Stencil category cannot be null or whitespace.", nameof(category));
            if (defaultWidth <= 0) throw new ArgumentOutOfRangeException(nameof(defaultWidth), "Default width must be positive.");
            if (defaultHeight <= 0) throw new ArgumentOutOfRangeException(nameof(defaultHeight), "Default height must be positive.");

            Id = id;
            Name = name;
            MasterNameU = masterNameU;
            Category = category;
            DefaultWidth = defaultWidth;
            DefaultHeight = defaultHeight;
            Keywords = (keywords ?? Enumerable.Empty<string>())
                .Where(keyword => !string.IsNullOrWhiteSpace(keyword))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList()
                .AsReadOnly();
            Aliases = (aliases ?? Enumerable.Empty<string>())
                .Where(alias => !string.IsNullOrWhiteSpace(alias))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList()
                .AsReadOnly();
            Tags = (tags ?? Enumerable.Empty<string>())
                .Where(tag => !string.IsNullOrWhiteSpace(tag))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList()
                .AsReadOnly();
            IconNameU = string.IsNullOrWhiteSpace(iconNameU) ? masterNameU : iconNameU!;
            DefaultUnit = defaultUnit;
        }

        /// <summary>
        /// Gets a stable OfficeIMO stencil identifier.
        /// </summary>
        public string Id { get; }

        /// <summary>
        /// Gets the display name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Gets the generated master universal name used by OfficeIMO.
        /// </summary>
        public string MasterNameU { get; }

        /// <summary>
        /// Gets the category name.
        /// </summary>
        public string Category { get; }

        /// <summary>
        /// Gets the default shape width in the caller's placement unit.
        /// </summary>
        public double DefaultWidth { get; }

        /// <summary>
        /// Gets the default shape height in the caller's placement unit.
        /// </summary>
        public double DefaultHeight { get; }

        /// <summary>
        /// Gets the unit used by the default size, when it is fixed by the source catalog.
        /// When null, default sizes are interpreted in the caller's placement unit.
        /// </summary>
        public VisioMeasurementUnit? DefaultUnit { get; }

        /// <summary>
        /// Gets searchable keywords.
        /// </summary>
        public IReadOnlyList<string> Keywords { get; }

        /// <summary>
        /// Gets alternate lookup names.
        /// </summary>
        public IReadOnlyList<string> Aliases { get; }

        /// <summary>
        /// Gets semantic tags used by stencil catalog search.
        /// </summary>
        public IReadOnlyList<string> Tags { get; }

        /// <summary>
        /// Gets the generated master universal name that can be used as this stencil shape's preview icon.
        /// </summary>
        public string IconNameU { get; }
    }
}
