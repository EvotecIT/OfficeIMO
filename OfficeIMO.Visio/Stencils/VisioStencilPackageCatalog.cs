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
        private static readonly string[] SupportedPackageExtensions = {
            ".vsdx",
            ".vssx",
            ".vstx",
            ".vsdm",
            ".vssm",
            ".vstm"
        };

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
            Dictionary<string, VisioAssets.MasterContent> masterContents = options.LearnMasterDimensions ||
                                                                            options.ExtractPreviewImageMetadata ||
                                                                            options.ExtractConnectionPointMetadata
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
                VisioStencilPreviewImage? previewImage = null;
                IReadOnlyList<VisioStencilConnectionPoint> sourceConnectionPoints = Array.Empty<VisioStencilConnectionPoint>();
                if (masterContents.TryGetValue(master.Id, out VisioAssets.MasterContent? content)) {
                    if (options.LearnMasterDimensions &&
                        TryReadMasterDimensions(content, out double masterWidth, out double masterHeight)) {
                        defaultWidth = masterWidth;
                        defaultHeight = masterHeight;
                        defaultUnit = VisioMeasurementUnit.Inches;
                    }

                    if (options.ExtractPreviewImageMetadata) {
                        previewImage = ReadPreviewImage(content);
                    }

                    if (options.ExtractConnectionPointMetadata) {
                        sourceConnectionPoints = ReadConnectionPoints(content);
                    }
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
                    defaultUnit,
                    Path.GetFullPath(packagePath),
                    previewImage,
                    sourceConnectionPoints);
            }

            return builder.Build();
        }

        /// <summary>
        /// Loads multiple Visio packages into one catalog. Each shape retains its source package path so
        /// <see cref="VisioStencilPageExtensions.AddStencilShape(VisioPage, VisioStencilShape, string, double, double, string?)"/>
        /// can import the real master automatically.
        /// </summary>
        /// <param name="packagePaths">Package paths to load.</param>
        /// <param name="options">Load options applied to every package.</param>
        public static VisioStencilCatalog LoadMany(IEnumerable<string> packagePaths, VisioStencilPackageLoadOptions? options = null) {
            if (packagePaths == null) throw new ArgumentNullException(nameof(packagePaths));

            options ??= new VisioStencilPackageLoadOptions();
            VisioStencilCatalogBuilder builder = new(string.IsNullOrWhiteSpace(options.CatalogName) ? "Visio Packages" : options.CatalogName!);
            foreach (string packagePath in packagePaths.Where(path => !string.IsNullOrWhiteSpace(path)).Distinct(StringComparer.OrdinalIgnoreCase)) {
                VisioStencilPackageLoadOptions packageOptions = CloneOptionsForPackage(options, packagePath);
                VisioStencilCatalog catalog = Load(packagePath, packageOptions);
                foreach (VisioStencilShape shape in catalog.Shapes) {
                    builder.Add(shape);
                }
            }

            return builder.Build();
        }

        /// <summary>
        /// Loads all supported Visio package files from a directory into one catalog.
        /// </summary>
        /// <param name="directoryPath">Directory containing `.vssx`, `.vstx`, or `.vsdx` packages.</param>
        /// <param name="options">Load options applied to every package.</param>
        /// <param name="recursive">Whether to search subdirectories.</param>
        public static VisioStencilCatalog LoadDirectory(string directoryPath, VisioStencilPackageLoadOptions? options = null, bool recursive = false) {
            return LoadMany(EnumeratePackageFiles(directoryPath, recursive), options);
        }

        /// <summary>
        /// Extracts embedded preview/icon image payloads from package-backed masters.
        /// </summary>
        /// <param name="packagePath">Path to a Visio package.</param>
        /// <param name="options">Load options used for master filtering and unsupported-master inclusion.</param>
        public static IReadOnlyList<VisioStencilPreviewImageData> ExtractPreviewImages(string packagePath, VisioStencilPackageLoadOptions? options = null) {
            if (string.IsNullOrWhiteSpace(packagePath)) throw new ArgumentException("Package path cannot be null or whitespace.", nameof(packagePath));
            if (!File.Exists(packagePath)) throw new FileNotFoundException("Visio package was not found.", packagePath);

            options ??= new VisioStencilPackageLoadOptions();
            HashSet<string>? filter = options.MasterNames != null
                ? new HashSet<string>(options.MasterNames.Where(name => !string.IsNullOrWhiteSpace(name)), StringComparer.OrdinalIgnoreCase)
                : null;
            List<VisioAssets.MasterInfo> masters = VisioAssets.ListMasters(packagePath)
                .Where(master => VisioMasterIdentity.MatchesAny(master, filter))
                .Where(master => options.IncludeUnsupportedMasters || VisioDocument.IsBuiltinMasterSupported(master.NameU))
                .ToList();
            if (masters.Count == 0) {
                return Array.Empty<VisioStencilPreviewImageData>();
            }

            Dictionary<string, VisioAssets.MasterContent> masterContents = VisioAssets.LoadMasterContents(packagePath, masters.Select(master => master.NameU))
                .GroupBy(master => master.Id, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);
            List<VisioStencilPreviewImageData> images = new();
            foreach (VisioAssets.MasterInfo master in masters) {
                if (!masterContents.TryGetValue(master.Id, out VisioAssets.MasterContent? content)) {
                    continue;
                }

                VisioAssets.MasterRelationshipContent? relationship = FindPreviewImageRelationship(content);
                if (relationship?.Data == null || relationship.Data.Length == 0) {
                    continue;
                }

                images.Add(new VisioStencilPreviewImageData(
                    master.Id,
                    master.NameU,
                    master.Name,
                    CreatePreviewImage(relationship),
                    relationship.Data));
            }

            return images.AsReadOnly();
        }

        /// <summary>
        /// Extracts embedded preview/icon image payloads from package-backed masters and saves them to a directory.
        /// </summary>
        /// <param name="packagePath">Path to a Visio package.</param>
        /// <param name="outputDirectory">Directory that receives extracted preview/icon files.</param>
        /// <param name="options">Load options used for master filtering and unsupported-master inclusion.</param>
        public static IReadOnlyList<string> ExtractPreviewImagesToDirectory(string packagePath, string outputDirectory, VisioStencilPackageLoadOptions? options = null) {
            if (string.IsNullOrWhiteSpace(outputDirectory)) throw new ArgumentException("Output directory cannot be null or whitespace.", nameof(outputDirectory));

            return ExtractPreviewImages(packagePath, options)
                .Select(image => image.SaveToDirectory(outputDirectory))
                .ToList()
                .AsReadOnly();
        }

        /// <summary>
        /// Extracts embedded preview/icon image payloads and writes a browsable HTML gallery index for review.
        /// </summary>
        /// <param name="packagePath">Path to a Visio package.</param>
        /// <param name="outputDirectory">Directory that receives the gallery index and extracted preview/icon files.</param>
        /// <param name="options">Load options used for master filtering and unsupported-master inclusion.</param>
        /// <param name="galleryOptions">Gallery output options.</param>
        public static VisioStencilPreviewGallery CreatePreviewGallery(
            string packagePath,
            string outputDirectory,
            VisioStencilPackageLoadOptions? options = null,
            VisioStencilPreviewGalleryOptions? galleryOptions = null) {
            if (string.IsNullOrWhiteSpace(outputDirectory)) throw new ArgumentException("Output directory cannot be null or whitespace.", nameof(outputDirectory));

            galleryOptions ??= new VisioStencilPreviewGalleryOptions();
            VisioStencilPreviewGalleryWriter.ValidateOptions(galleryOptions);
            IReadOnlyList<VisioStencilPreviewImageData> images = ExtractPreviewImages(packagePath, options);
            return VisioStencilPreviewGalleryWriter.Create(packagePath, outputDirectory, images, galleryOptions);
        }

        /// <summary>
        /// Enumerates supported Visio package files from a directory.
        /// </summary>
        /// <param name="directoryPath">Directory containing Visio packages.</param>
        /// <param name="recursive">Whether to search subdirectories.</param>
        public static IReadOnlyList<string> EnumeratePackageFiles(string directoryPath, bool recursive = false) {
            if (string.IsNullOrWhiteSpace(directoryPath)) throw new ArgumentException("Directory path cannot be null or whitespace.", nameof(directoryPath));
            if (!Directory.Exists(directoryPath)) throw new DirectoryNotFoundException("Visio package directory was not found: " + directoryPath);

            SearchOption searchOption = recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            return Directory
                .EnumerateFiles(directoryPath, "*.*", searchOption)
                .Where(IsSupportedPackagePath)
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToList()
                .AsReadOnly();
        }

        /// <summary>
        /// Discovers installed Microsoft Visio package stencils and templates without automating Visio.
        /// </summary>
        public static IReadOnlyList<string> DiscoverInstalledVisioPackages() {
            return GetInstalledVisioContentDirectories()
                .SelectMany(directory => Directory.Exists(directory) ? EnumeratePackageFiles(directory, recursive: true) : Enumerable.Empty<string>())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToList()
                .AsReadOnly();
        }

        /// <summary>
        /// Gets likely local Visio content directories.
        /// </summary>
        public static IReadOnlyList<string> GetInstalledVisioContentDirectories() {
            List<string> directories = new();
            AddOfficeVisioContentDirectory(directories, Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles));
            AddOfficeVisioContentDirectory(directories, Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86));

            string documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (!string.IsNullOrWhiteSpace(documents)) {
                directories.Add(Path.Combine(documents, "My Shapes"));
            }

            return directories
                .Where(directory => !string.IsNullOrWhiteSpace(directory))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList()
                .AsReadOnly();
        }

        private static bool TryReadMasterDimensions(VisioAssets.MasterContent content, out double width, out double height) {
            XElement? shape = GetMasterRootShape(content);

            if (shape != null &&
                TryReadPositiveCell(shape, "Width", out width) &&
                TryReadPositiveCell(shape, "Height", out height)) {
                return true;
            }

            width = 0;
            height = 0;
            return false;
        }

        private static XElement? GetMasterRootShape(VisioAssets.MasterContent content) {
            XNamespace v = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement? shape = content.MasterXml.Root?
                .Element(v + "Shapes")?
                .Elements(v + "Shape")
                .FirstOrDefault();
            if (shape != null) {
                return shape;
            }

            return content.MasterXml.Root?
                .Elements()
                .FirstOrDefault(element => string.Equals(element.Name.LocalName, "Shapes", StringComparison.OrdinalIgnoreCase))?
                .Elements()
                .FirstOrDefault(element => string.Equals(element.Name.LocalName, "Shape", StringComparison.OrdinalIgnoreCase));
        }

        private static VisioStencilPreviewImage? ReadPreviewImage(VisioAssets.MasterContent content) {
            VisioAssets.MasterRelationshipContent? relationship = FindPreviewImageRelationship(content);
            return relationship == null ? null : CreatePreviewImage(relationship);
        }

        private static VisioStencilPreviewImage CreatePreviewImage(VisioAssets.MasterRelationshipContent relationship) {
            return new VisioStencilPreviewImage(
                relationship.Id,
                relationship.Target,
                relationship.ContentType,
                relationship.Extension,
                relationship.Data?.LongLength,
                relationship.IsExternal);
        }

        private static VisioAssets.MasterRelationshipContent? FindPreviewImageRelationship(VisioAssets.MasterContent content) {
            HashSet<string> preferredRelationshipIds = GetPreferredPreviewRelationshipIds(content);
            return content.Relationships
                .Where(IsImageRelationship)
                .OrderByDescending(item => preferredRelationshipIds.Contains(item.Id))
                .ThenBy(item => item.Id, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault();
        }

        private static HashSet<string> GetPreferredPreviewRelationshipIds(VisioAssets.MasterContent content) {
            XNamespace rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            return new HashSet<string>(content.MasterXml
                .Descendants()
                .Where(element => string.Equals(element.Name.LocalName, "ForeignData", StringComparison.OrdinalIgnoreCase))
                .Select(element => (string?)element.Attribute(rel + "id"))
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => value!),
                StringComparer.OrdinalIgnoreCase);
        }

        private static bool IsImageRelationship(VisioAssets.MasterRelationshipContent relationship) {
            return relationship.Type.EndsWith("/image", StringComparison.OrdinalIgnoreCase) ||
                   relationship.ContentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase) ||
                   IsImageExtension(relationship.Extension) ||
                   IsImageExtension(Path.GetExtension(relationship.Target));
        }

        private static bool IsImageExtension(string? extension) {
            if (string.IsNullOrWhiteSpace(extension)) {
                return false;
            }

            string normalized = extension!.TrimStart('.').ToLowerInvariant();
            return normalized switch {
                "emf" or "wmf" or "png" or "jpg" or "jpeg" or "gif" or "svg" or "tif" or "tiff" or "bmp" => true,
                _ => false
            };
        }

        private static VisioStencilPackageLoadOptions CloneOptionsForPackage(VisioStencilPackageLoadOptions options, string packagePath) {
            string fileName = Path.GetFileNameWithoutExtension(packagePath);
            return new VisioStencilPackageLoadOptions {
                CatalogName = fileName,
                Category = string.IsNullOrWhiteSpace(options.Category) ? fileName : options.Category,
                IdPrefix = string.IsNullOrWhiteSpace(options.IdPrefix) ? fileName : options.IdPrefix + "." + fileName,
                MasterNames = options.MasterNames,
                IncludeUnsupportedMasters = options.IncludeUnsupportedMasters,
                LearnMasterDimensions = options.LearnMasterDimensions,
                ExtractPreviewImageMetadata = options.ExtractPreviewImageMetadata,
                ExtractConnectionPointMetadata = options.ExtractConnectionPointMetadata,
                DefaultWidth = options.DefaultWidth,
                DefaultHeight = options.DefaultHeight
            };
        }

        private static IReadOnlyList<VisioStencilConnectionPoint> ReadConnectionPoints(VisioAssets.MasterContent content) {
            XElement? shape = GetMasterRootShape(content);
            XElement? connectionSection = shape?
                .Elements()
                .FirstOrDefault(element =>
                    string.Equals(element.Name.LocalName, "Section", StringComparison.OrdinalIgnoreCase) &&
                    string.Equals((string?)element.Attribute("N"), "Connection", StringComparison.OrdinalIgnoreCase));
            if (connectionSection == null) {
                return Array.Empty<VisioStencilConnectionPoint>();
            }

            List<VisioStencilConnectionPoint> points = new();
            foreach (XElement row in connectionSection.Elements().Where(element => string.Equals(element.Name.LocalName, "Row", StringComparison.OrdinalIgnoreCase))) {
                if (!TryReadCell(row, "X", out double x) ||
                    !TryReadCell(row, "Y", out double y)) {
                    continue;
                }

                TryReadCell(row, "DirX", out double dirX);
                TryReadCell(row, "DirY", out double dirY);
                int? sectionIndex = null;
                if (int.TryParse((string?)row.Attribute("IX"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedSectionIndex) &&
                    parsedSectionIndex >= 0) {
                    sectionIndex = parsedSectionIndex;
                }

                points.Add(new VisioStencilConnectionPoint(x, y, dirX, dirY, sectionIndex));
            }

            return points.AsReadOnly();
        }

        private static bool IsSupportedPackagePath(string path) {
            return SupportedPackageExtensions.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase);
        }

        private static void AddOfficeVisioContentDirectory(ICollection<string> directories, string root) {
            if (string.IsNullOrWhiteSpace(root)) {
                return;
            }

            string contentRoot = Path.Combine(root, "Microsoft Office", "root", "Office16", "Visio Content");
            directories.Add(Path.Combine(contentRoot, CultureInfo.CurrentUICulture.LCID.ToString(CultureInfo.InvariantCulture)));
            directories.Add(Path.Combine(contentRoot, "1033"));
            directories.Add(contentRoot);
        }

        private static bool TryReadPositiveCell(XElement shape, string name, out double value) {
            if (!TryReadCell(shape, name, out value) ||
                value <= 0) {
                value = 0;
                return false;
            }

            return true;
        }

        private static bool TryReadCell(XElement element, string name, out double value) {
            XElement? cell = element.Elements()
                .FirstOrDefault(child =>
                    string.Equals(child.Name.LocalName, "Cell", StringComparison.OrdinalIgnoreCase) &&
                    string.Equals((string?)child.Attribute("N"), name, StringComparison.OrdinalIgnoreCase));
            string? rawValue = (string?)cell?.Attribute("V");
            if (string.IsNullOrWhiteSpace(rawValue) ||
                !double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out value) ||
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
