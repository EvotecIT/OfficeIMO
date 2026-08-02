using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;

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
        public VisioStencilShape(string id, string name, string masterNameU, string category, double defaultWidth, double defaultHeight, IEnumerable<string>? keywords, IEnumerable<string>? aliases, IEnumerable<string>? tags, string? iconNameU, VisioMeasurementUnit? defaultUnit)
            : this(id, name, masterNameU, category, defaultWidth, defaultHeight, keywords, aliases, tags, iconNameU, defaultUnit, null) {
        }

        /// <summary>
        /// Initializes a new stencil shape definition with source package metadata.
        /// </summary>
        public VisioStencilShape(string id, string name, string masterNameU, string category, double defaultWidth, double defaultHeight, IEnumerable<string>? keywords, IEnumerable<string>? aliases, IEnumerable<string>? tags, string? iconNameU, VisioMeasurementUnit? defaultUnit, string? sourcePackagePath)
            : this(id, name, masterNameU, category, defaultWidth, defaultHeight, keywords, aliases, tags, iconNameU, defaultUnit, sourcePackagePath, null) {
        }

        /// <summary>
        /// Initializes a new stencil shape definition with source package and preview image metadata.
        /// </summary>
        public VisioStencilShape(string id, string name, string masterNameU, string category, double defaultWidth, double defaultHeight, IEnumerable<string>? keywords, IEnumerable<string>? aliases, IEnumerable<string>? tags, string? iconNameU, VisioMeasurementUnit? defaultUnit, string? sourcePackagePath, VisioStencilPreviewImage? previewImage)
            : this(id, name, masterNameU, category, defaultWidth, defaultHeight, keywords, aliases, tags, iconNameU, defaultUnit, sourcePackagePath, previewImage, null) {
        }

        /// <summary>
        /// Initializes a new stencil shape definition with source package, preview image, and connection point metadata.
        /// </summary>
        public VisioStencilShape(string id, string name, string masterNameU, string category, double defaultWidth, double defaultHeight, IEnumerable<string>? keywords, IEnumerable<string>? aliases, IEnumerable<string>? tags, string? iconNameU, VisioMeasurementUnit? defaultUnit, string? sourcePackagePath, VisioStencilPreviewImage? previewImage, IEnumerable<VisioStencilConnectionPoint>? sourceConnectionPoints)
            : this(id, name, masterNameU, category, defaultWidth, defaultHeight, keywords, aliases, tags, iconNameU, defaultUnit, sourcePackagePath, previewImage, sourceConnectionPoints, true, null, null) {
        }

        /// <summary>
        /// Initializes a new stencil shape definition with source package, preview image, connection point, support, and licensing metadata.
        /// </summary>
        public VisioStencilShape(string id, string name, string masterNameU, string category, double defaultWidth, double defaultHeight, IEnumerable<string>? keywords, IEnumerable<string>? aliases, IEnumerable<string>? tags, string? iconNameU, VisioMeasurementUnit? defaultUnit, string? sourcePackagePath, VisioStencilPreviewImage? previewImage, IEnumerable<VisioStencilConnectionPoint>? sourceConnectionPoints, bool isSupported = true, string? sourceLicense = null, string? sourceAttribution = null) {
            if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Stencil shape id cannot be null or whitespace.", nameof(id));
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Stencil shape name cannot be null or whitespace.", nameof(name));
            if (string.IsNullOrWhiteSpace(masterNameU)) throw new ArgumentException("Master NameU cannot be null or whitespace.", nameof(masterNameU));
            if (string.IsNullOrWhiteSpace(category)) throw new ArgumentException("Stencil category cannot be null or whitespace.", nameof(category));
            if (defaultWidth <= 0) throw new ArgumentOutOfRangeException(nameof(defaultWidth), "Default width must be positive.");
            if (defaultHeight <= 0) throw new ArgumentOutOfRangeException(nameof(defaultHeight), "Default height must be positive.");

            Id = ValidateXmlValue(id, nameof(id));
            Name = ValidateXmlValue(name, nameof(name));
            MasterNameU = ValidateXmlValue(masterNameU, nameof(masterNameU));
            Category = ValidateXmlValue(category, nameof(category));
            DefaultWidth = defaultWidth;
            DefaultHeight = defaultHeight;
            Keywords = ValidateXmlValues(keywords, nameof(keywords))
                .Where(keyword => !string.IsNullOrWhiteSpace(keyword))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList()
                .AsReadOnly();
            Aliases = ValidateXmlValues(aliases, nameof(aliases))
                .Where(alias => !string.IsNullOrWhiteSpace(alias))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList()
                .AsReadOnly();
            Tags = ValidateXmlValues(tags, nameof(tags))
                .Where(tag => !string.IsNullOrWhiteSpace(tag))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList()
                .AsReadOnly();
            IconNameU = string.IsNullOrWhiteSpace(iconNameU)
                ? MasterNameU
                : ValidateXmlValue(iconNameU!, nameof(iconNameU));
            DefaultUnit = defaultUnit;
            SourcePackagePath = NormalizeOptional(sourcePackagePath,
                nameof(sourcePackagePath));
            PreviewImage = previewImage;
            SourceConnectionPoints = (sourceConnectionPoints ?? Enumerable.Empty<VisioStencilConnectionPoint>())
                .Where(point => point != null)
                .ToList()
                .AsReadOnly();
            IsSupported = isSupported;
            SourceLicense = NormalizeOptional(sourceLicense,
                nameof(sourceLicense));
            SourceAttribution = NormalizeOptional(sourceAttribution,
                nameof(sourceAttribution));
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
        /// Gets the source `.vssx`, `.vstx`, or `.vsdx` package path when this shape was cataloged from a package.
        /// </summary>
        public string? SourcePackagePath { get; }

        /// <summary>
        /// Gets preview/icon image metadata discovered from a source package master, when available.
        /// </summary>
        public VisioStencilPreviewImage? PreviewImage { get; }

        /// <summary>
        /// Gets native connection points discovered from a source package master, when available.
        /// </summary>
        public IReadOnlyList<VisioStencilConnectionPoint> SourceConnectionPoints { get; }

        /// <summary>
        /// Gets whether OfficeIMO has a typed master implementation for this shape.
        /// Unsupported package masters may still be cataloged explicitly as generic placeholders.
        /// </summary>
        public bool IsSupported { get; }

        /// <summary>
        /// Gets the caller-supplied source package license identifier or notice.
        /// OfficeIMO does not infer or grant rights to third-party stencil content.
        /// </summary>
        public string? SourceLicense { get; }

        /// <summary>Gets the caller-supplied source attribution.</summary>
        public string? SourceAttribution { get; }

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

        private static string? NormalizeOptional(string? value,
            string parameterName) => string.IsNullOrWhiteSpace(value)
                ? null
                : ValidateXmlValue(value!.Trim(), parameterName);

        private static IEnumerable<string> ValidateXmlValues(
            IEnumerable<string>? values, string parameterName) {
            foreach (string value in values ?? Enumerable.Empty<string>()) {
                if (value != null) yield return ValidateXmlValue(value,
                    parameterName);
            }
        }

        private static string ValidateXmlValue(string value,
            string parameterName) {
            try {
                XmlConvert.VerifyXmlChars(value);
                return value;
            } catch (XmlException exception) {
                throw new ArgumentException(
                    "Stencil metadata contains characters that cannot be represented in XML.",
                    parameterName, exception);
            }
        }
    }
}
