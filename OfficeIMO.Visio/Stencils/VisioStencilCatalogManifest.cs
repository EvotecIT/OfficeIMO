using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Visio.Stencils {
    /// <summary>
    /// Reads and writes dependency-free OfficeIMO-native stencil catalog manifests.
    /// </summary>
    public static class VisioStencilCatalogManifest {
        private const string FormatVersion = "1";
        private static readonly XNamespace Ns = "urn:officeimo:visio:stencils";
        private const long MaxManifestXmlCharacters = 10_000_000;

        private static readonly XmlReaderSettings ManifestXmlReaderSettings = new() {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = MaxManifestXmlCharacters,
            MaxCharactersFromEntities = 0,
        };

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
            return Load(path, null);
        }

        /// <summary>
        /// Loads a stencil catalog manifest from a file with explicit source-package trust options.
        /// </summary>
        public static VisioStencilCatalog Load(string path, VisioStencilCatalogManifestLoadOptions? options) {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Manifest path cannot be null or whitespace.", nameof(path));
            using FileStream stream = File.OpenRead(path);
            return Load(stream, CreateOptionsForPath(path, options));
        }

        /// <summary>
        /// Loads a stencil catalog manifest from a stream.
        /// </summary>
        public static VisioStencilCatalog Load(Stream stream) {
            return Load(stream, null);
        }

        /// <summary>
        /// Loads a stencil catalog manifest from a stream with explicit source-package trust options.
        /// </summary>
        public static VisioStencilCatalog Load(Stream stream, VisioStencilCatalogManifestLoadOptions? options) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

            XDocument document = LoadXml(stream);
            return FromXml(document, options);
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
                        string.IsNullOrWhiteSpace(shape.SourcePackagePath) ? null : new XAttribute("SourcePackagePath", shape.SourcePackagePath),
                        PreviewImage(shape.PreviewImage),
                        ConnectionPoints(shape.SourceConnectionPoints),
                        Values("Keywords", shape.Keywords),
                        Values("Aliases", shape.Aliases),
                        Values("Tags", shape.Tags))));

            return new XDocument(new XDeclaration("1.0", "utf-8", null), root);
        }

        /// <summary>
        /// Converts an XML manifest document to a stencil catalog.
        /// </summary>
        public static VisioStencilCatalog FromXml(XDocument document) {
            return FromXml(document, null);
        }

        /// <summary>
        /// Converts an XML manifest document to a stencil catalog with explicit source-package trust options.
        /// </summary>
        public static VisioStencilCatalog FromXml(XDocument document, VisioStencilCatalogManifestLoadOptions? options) {
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
                    ReadUnit(shape, "DefaultUnit"),
                    ResolveSourcePackagePath((string?)shape.Attribute("SourcePackagePath"), options),
                    ReadPreviewImage(shape),
                    ReadConnectionPoints(shape));
            }

            return builder.Build();
        }

        private static VisioStencilCatalogManifestLoadOptions? CreateOptionsForPath(string path, VisioStencilCatalogManifestLoadOptions? options) {
            string? directory = Path.GetDirectoryName(Path.GetFullPath(path));
            if (string.IsNullOrWhiteSpace(directory)) {
                return options;
            }

            if (options == null) {
                return new VisioStencilCatalogManifestLoadOptions {
                    BaseDirectory = directory
                };
            }

            return new VisioStencilCatalogManifestLoadOptions {
                AllowSourcePackagePaths = options.AllowSourcePackagePaths,
                AllowExternalSourcePackagePaths = options.AllowExternalSourcePackagePaths,
                BaseDirectory = string.IsNullOrWhiteSpace(options.BaseDirectory) ? directory : options.BaseDirectory
            };
        }

        private static string? ResolveSourcePackagePath(string? sourcePackagePath, VisioStencilCatalogManifestLoadOptions? options) {
            if (string.IsNullOrWhiteSpace(sourcePackagePath) ||
                options?.AllowSourcePackagePaths != true) {
                return null;
            }

            string candidate = sourcePackagePath!.Trim();
            if (Uri.TryCreate(candidate, UriKind.Absolute, out Uri? uri)) {
                if (!uri.IsFile) {
                    return null;
                }

                candidate = uri.LocalPath;
            }

            try {
                bool hasBaseDirectory = !string.IsNullOrWhiteSpace(options.BaseDirectory);
                if (!Path.IsPathRooted(candidate) && !hasBaseDirectory) {
                    return null;
                }

                if (!Path.IsPathRooted(candidate)) {
                    candidate = Path.Combine(options.BaseDirectory!, candidate);
                }

                string fullPath = Path.GetFullPath(candidate);
                if (hasBaseDirectory &&
                    !options.AllowExternalSourcePackagePaths &&
                    !IsPathWithinDirectory(fullPath, Path.GetFullPath(options.BaseDirectory!))) {
                    return null;
                }

                if (!hasBaseDirectory &&
                    Path.IsPathRooted(candidate) &&
                    !options.AllowExternalSourcePackagePaths) {
                    return null;
                }

                return fullPath;
            } catch (ArgumentException) {
                return null;
            } catch (NotSupportedException) {
                return null;
            } catch (PathTooLongException) {
                return null;
            }
        }

        private static bool IsPathWithinDirectory(string path, string directory) {
            string normalizedDirectory = directory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                + Path.DirectorySeparatorChar;
            string normalizedPath = path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            StringComparison comparison = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
                ? StringComparison.OrdinalIgnoreCase
                : StringComparison.Ordinal;
            return normalizedPath.StartsWith(normalizedDirectory, comparison);
        }

        private static XElement Values(string name, IEnumerable<string> values) {
            return new XElement(Ns + name,
                values
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Select(value => new XElement(Ns + "Value", value)));
        }

        private static XElement? PreviewImage(VisioStencilPreviewImage? previewImage) {
            if (previewImage == null) {
                return null;
            }

            return new XElement(Ns + "PreviewImage",
                new XAttribute("RelationshipId", previewImage.RelationshipId),
                new XAttribute("Target", previewImage.Target),
                string.IsNullOrWhiteSpace(previewImage.ContentType) ? null : new XAttribute("ContentType", previewImage.ContentType),
                string.IsNullOrWhiteSpace(previewImage.Extension) ? null : new XAttribute("Extension", previewImage.Extension),
                previewImage.ByteLength.HasValue ? new XAttribute("ByteLength", previewImage.ByteLength.Value) : null,
                previewImage.IsExternal ? new XAttribute("External", true) : null);
        }

        private static XElement? ConnectionPoints(IReadOnlyList<VisioStencilConnectionPoint> points) {
            if (points.Count == 0) {
                return null;
            }

            return new XElement(Ns + "ConnectionPoints",
                points.Select(point =>
                    new XElement(Ns + "ConnectionPoint",
                        point.SectionIndex.HasValue ? new XAttribute("IX", point.SectionIndex.Value) : null,
                        new XAttribute("X", XmlConvert.ToString(point.X)),
                        new XAttribute("Y", XmlConvert.ToString(point.Y)),
                        new XAttribute("DirX", XmlConvert.ToString(point.DirX)),
                        new XAttribute("DirY", XmlConvert.ToString(point.DirY)),
                        point.SourceWidth.HasValue ? new XAttribute("SourceWidth", XmlConvert.ToString(point.SourceWidth.Value)) : null,
                        point.SourceHeight.HasValue ? new XAttribute("SourceHeight", XmlConvert.ToString(point.SourceHeight.Value)) : null)));
        }

        private static VisioStencilPreviewImage? ReadPreviewImage(XElement shape) {
            XElement? preview = shape.Element(Ns + "PreviewImage");
            if (preview == null) {
                return null;
            }

            return new VisioStencilPreviewImage(
                Required(preview, "RelationshipId"),
                Required(preview, "Target"),
                (string?)preview.Attribute("ContentType"),
                (string?)preview.Attribute("Extension"),
                ReadNullableLong(preview, "ByteLength"),
                ReadBoolean(preview, "External"));
        }

        private static IReadOnlyList<VisioStencilConnectionPoint> ReadConnectionPoints(XElement shape) {
            XElement? section = shape.Element(Ns + "ConnectionPoints");
            if (section == null) {
                return Array.Empty<VisioStencilConnectionPoint>();
            }

            return section.Elements(Ns + "ConnectionPoint")
                .Select(point => new VisioStencilConnectionPoint(
                    ReadDouble(point, "X"),
                    ReadDouble(point, "Y"),
                    ReadDouble(point, "DirX"),
                    ReadDouble(point, "DirY"),
                    ReadNullableInt(point, "IX"),
                    ReadNullableDouble(point, "SourceWidth"),
                    ReadNullableDouble(point, "SourceHeight")))
                .ToList()
                .AsReadOnly();
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

        private static long? ReadNullableLong(XElement element, string attributeName) {
            string? value = (string?)element.Attribute(attributeName);
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            return XmlConvert.ToInt64(value);
        }

        private static int? ReadNullableInt(XElement element, string attributeName) {
            string? value = (string?)element.Attribute(attributeName);
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            int parsed = XmlConvert.ToInt32(value);
            if (parsed < 0) {
                throw new InvalidDataException($"Stencil manifest attribute '{attributeName}' must be zero or greater.");
            }

            return parsed;
        }

        private static double ReadDouble(XElement element, string attributeName) {
            string value = Required(element, attributeName);
            double parsed = XmlConvert.ToDouble(value);
            if (double.IsNaN(parsed) || double.IsInfinity(parsed)) {
                throw new InvalidDataException($"Stencil manifest attribute '{attributeName}' must be finite.");
            }

            return parsed;
        }

        private static double? ReadNullableDouble(XElement element, string attributeName) {
            string? value = (string?)element.Attribute(attributeName);
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            double parsed = XmlConvert.ToDouble(value);
            if (double.IsNaN(parsed) || double.IsInfinity(parsed) || parsed <= 0) {
                throw new InvalidDataException($"Stencil manifest attribute '{attributeName}' must be positive and finite.");
            }

            return parsed;
        }

        private static bool ReadBoolean(XElement element, string attributeName) {
            string? value = (string?)element.Attribute(attributeName);
            return !string.IsNullOrWhiteSpace(value) && XmlConvert.ToBoolean(value);
        }

        private static XDocument LoadXml(Stream stream) {
            using XmlReader reader = XmlReader.Create(stream, ManifestXmlReaderSettings);
            return XDocument.Load(reader, LoadOptions.None);
        }
    }
}
