using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Visio.Stencils {
    /// <summary>
    /// Reads and writes dependency-free OfficeIMO-native stencil catalog manifests.
    /// </summary>
    public static class VisioStencilCatalogManifest {
        private const string FormatVersion = "1";
        private static readonly XNamespace Ns = "urn:officeimo:visio:stencils";

        /// <summary>
        /// Saves a stencil catalog manifest to a file.
        /// </summary>
        public static void Save(VisioStencilCatalog catalog, string path) {
            if (catalog == null) throw new ArgumentNullException(nameof(catalog));
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Manifest path cannot be null or whitespace.", nameof(path));

            string? directory = Path.GetDirectoryName(Path.GetFullPath(path));
            if (!string.IsNullOrEmpty(directory)) {
                Directory.CreateDirectory(directory);
            }

            using FileStream stream = File.Create(path);
            Save(catalog, stream);
        }

        /// <summary>
        /// Saves a stencil catalog manifest to a stream.
        /// </summary>
        public static void Save(VisioStencilCatalog catalog, Stream stream) {
            if (catalog == null) throw new ArgumentNullException(nameof(catalog));
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanWrite) throw new ArgumentException("Stream must be writable.", nameof(stream));

            XDocument document = ToXml(catalog);
            document.Save(stream, SaveOptions.DisableFormatting);
        }

        /// <summary>
        /// Loads a stencil catalog manifest from a file.
        /// </summary>
        public static VisioStencilCatalog Load(string path) {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Manifest path cannot be null or whitespace.", nameof(path));
            using FileStream stream = File.OpenRead(path);
            return Load(stream);
        }

        /// <summary>
        /// Loads a stencil catalog manifest from a stream.
        /// </summary>
        public static VisioStencilCatalog Load(Stream stream) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

            XDocument document = XDocument.Load(stream);
            return FromXml(document);
        }

        /// <summary>
        /// Converts a stencil catalog to an XML manifest document.
        /// </summary>
        public static XDocument ToXml(VisioStencilCatalog catalog) {
            if (catalog == null) throw new ArgumentNullException(nameof(catalog));

            XElement root = new(Ns + "StencilCatalog",
                new XAttribute("Version", FormatVersion),
                new XAttribute("Name", catalog.Name),
                catalog.Shapes.Select(shape =>
                    new XElement(Ns + "Shape",
                        new XAttribute("Id", shape.Id),
                        new XAttribute("Name", shape.Name),
                        new XAttribute("MasterNameU", shape.MasterNameU),
                        new XAttribute("Category", shape.Category),
                        new XAttribute("DefaultWidth", XmlConvert.ToString(shape.DefaultWidth)),
                        new XAttribute("DefaultHeight", XmlConvert.ToString(shape.DefaultHeight)),
                        new XAttribute("IconNameU", shape.IconNameU),
                        shape.DefaultUnit.HasValue ? new XAttribute("DefaultUnit", shape.DefaultUnit.Value.ToString()) : null,
                        Values("Keywords", shape.Keywords),
                        Values("Aliases", shape.Aliases),
                        Values("Tags", shape.Tags))));

            return new XDocument(new XDeclaration("1.0", "utf-8", null), root);
        }

        /// <summary>
        /// Converts an XML manifest document to a stencil catalog.
        /// </summary>
        public static VisioStencilCatalog FromXml(XDocument document) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            XElement root = document.Root ?? throw new InvalidDataException("Stencil manifest does not contain a root element.");
            if (root.Name != Ns + "StencilCatalog") {
                throw new InvalidDataException("Stencil manifest root element is not recognized.");
            }

            string version = (string?)root.Attribute("Version") ?? string.Empty;
            if (!string.Equals(version, FormatVersion, StringComparison.Ordinal)) {
                throw new InvalidDataException($"Stencil manifest version '{version}' is not supported.");
            }

            string name = Required(root, "Name");
            VisioStencilCatalogBuilder builder = new(name);
            foreach (XElement shape in root.Elements(Ns + "Shape")) {
                builder.AddWithMetadata(
                    Required(shape, "Id"),
                    Required(shape, "Name"),
                    Required(shape, "MasterNameU"),
                    Required(shape, "Category"),
                    ReadPositiveDouble(shape, "DefaultWidth"),
                    ReadPositiveDouble(shape, "DefaultHeight"),
                    ReadValues(shape, "Keywords"),
                    ReadValues(shape, "Aliases"),
                    ReadValues(shape, "Tags"),
                    (string?)shape.Attribute("IconNameU"),
                    ReadUnit(shape, "DefaultUnit"));
            }

            return builder.Build();
        }

        private static XElement Values(string name, IEnumerable<string> values) {
            return new XElement(Ns + name,
                values
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Select(value => new XElement(Ns + "Value", value)));
        }

        private static IReadOnlyList<string> ReadValues(XElement shape, string name) {
            IReadOnlyList<string>? values = shape.Element(Ns + name)?
                .Elements(Ns + "Value")
                .Select(value => value.Value)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList()
                .AsReadOnly();

            return values ?? Array.Empty<string>();
        }

        private static string Required(XElement element, string attributeName) {
            string? value = (string?)element.Attribute(attributeName);
            if (string.IsNullOrWhiteSpace(value)) {
                throw new InvalidDataException($"Stencil manifest element '{element.Name.LocalName}' is missing required attribute '{attributeName}'.");
            }

            return value!;
        }

        private static double ReadPositiveDouble(XElement element, string attributeName) {
            string value = Required(element, attributeName);
            double parsed = double.Parse(value, CultureInfo.InvariantCulture);
            if (parsed <= 0) {
                throw new InvalidDataException($"Stencil manifest attribute '{attributeName}' must be positive.");
            }

            return parsed;
        }

        private static VisioMeasurementUnit? ReadUnit(XElement element, string attributeName) {
            string? value = (string?)element.Attribute(attributeName);
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            if (Enum.TryParse(value, ignoreCase: true, out VisioMeasurementUnit unit)) {
                return unit;
            }

            throw new InvalidDataException($"Stencil manifest attribute '{attributeName}' is not a supported measurement unit.");
        }
    }
}
