using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Core.Internal;
using OfficeIMO.Visio.Stencils;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents a Visio document containing pages.
    /// </summary>
    public partial class VisioDocument {
        private enum BuiltinGeometryKind {
            Rectangle,
            Ellipse,
            Diamond,
            Triangle,
            Pentagon,
            Parallelogram,
            Hexagon,
            Trapezoid,
            OffPageReference,
            DynamicConnector
        }

        private sealed class BuiltinMasterDefinition {
            public BuiltinMasterDefinition(
                string id,
                BuiltinGeometryKind geometryKind,
                bool lockAspect,
                string prompt,
                bool matchByName,
                bool iconUpdate,
                int masterType,
                string shapeKeywords,
                bool addConnectorLayer = false) {
                Id = id;
                GeometryKind = geometryKind;
                LockAspect = lockAspect;
                Prompt = prompt;
                MatchByName = matchByName;
                IconUpdate = iconUpdate;
                MasterType = masterType;
                ShapeKeywords = shapeKeywords;
                AddConnectorLayer = addConnectorLayer;
            }

            public string Id { get; }

            public BuiltinGeometryKind GeometryKind { get; }

            public bool LockAspect { get; }

            public string Prompt { get; }

            public bool MatchByName { get; }

            public bool IconUpdate { get; }

            public int MasterType { get; }

            public string ShapeKeywords { get; }

            public bool AddConnectorLayer { get; }

            public string BaseId { get; set; } = string.Empty;

            public string UniqueId { get; set; } = string.Empty;
        }

        internal sealed class PreservedStyleSheetData {
            public IList<XAttribute> Attributes { get; } = new List<XAttribute>();

            public IList<XElement> ChildElements { get; } = new List<XElement>();
        }

        private readonly List<VisioPage> _pages = new();
        private bool _requestRecalcOnOpen;
        private string? _filePath;
        private Stream? _sourceStream;
        private readonly Dictionary<string, VisioMaster> _builtinMasters = new(StringComparer.OrdinalIgnoreCase);
        private readonly List<VisioMaster> _registeredMasters = new();
        internal IList<XAttribute> PreservedDocumentAttributes { get; } = new List<XAttribute>();
        internal IList<XElement> PreservedDocumentElements { get; } = new List<XElement>();
        internal IList<XAttribute> PreservedDocumentSettingsAttributes { get; } = new List<XAttribute>();
        internal IList<XElement> PreservedDocumentSettingsElements { get; } = new List<XElement>();
        internal IList<XAttribute> PreservedColorsAttributes { get; } = new List<XAttribute>();
        internal IList<XElement> PreservedColorsElements { get; } = new List<XElement>();
        internal IList<XAttribute> PreservedFaceNamesAttributes { get; } = new List<XAttribute>();
        internal IList<XElement> PreservedFaceNamesElements { get; } = new List<XElement>();
        internal IList<XAttribute> PreservedStyleSheetsAttributes { get; } = new List<XAttribute>();
        internal IList<XElement> PreservedStyleSheetsElements { get; } = new List<XElement>();
        internal IDictionary<string, PreservedStyleSheetData> PreservedGeneratedStyleSheets { get; } = new Dictionary<string, PreservedStyleSheetData>(StringComparer.Ordinal);
        internal IList<XElement> PreservedAdditionalStyleSheets { get; } = new List<XElement>();
        private static readonly IReadOnlyDictionary<string, BuiltinMasterDefinition> BuiltinMasterDefinitions = CreateBuiltinMasterDefinitions();

        private const string DocumentRelationshipType = "http://schemas.microsoft.com/visio/2010/relationships/document";
        private const string DocumentContentType = "application/vnd.ms-visio.drawing.main+xml";
        private const string VisioNamespace = "http://schemas.microsoft.com/office/visio/2012/main";
        private const string ThemeRelationshipType = "http://schemas.microsoft.com/visio/2010/relationships/theme";
        private const string ThemeContentType = "application/vnd.ms-visio.theme+xml";
        private const string CommentsRelationshipType = "http://schemas.microsoft.com/visio/2010/relationships/comments";
        private const string CommentsContentType = "application/vnd.ms-visio.comments+xml";
        private const string WindowsRelationshipType = "http://schemas.microsoft.com/visio/2010/relationships/windows";
        private const string WindowsContentType = "application/vnd.ms-visio.windows+xml";
        private const string PagesRelationshipType = "http://schemas.microsoft.com/visio/2010/relationships/pages";
        private const string PagesContentType = "application/vnd.ms-visio.pages+xml";
        private const string PageRelationshipType = "http://schemas.microsoft.com/visio/2010/relationships/page";
        private const string PageContentType = "application/vnd.ms-visio.page+xml";
        private const string MastersRelationshipType = "http://schemas.microsoft.com/visio/2010/relationships/masters";
        private const string MasterRelationshipType = "http://schemas.microsoft.com/visio/2010/relationships/master";
        private const string OriginalIdPropName = "OfficeIMOOriginalId";

        /// <summary>
        /// Gets the collection of pages in the document.
        /// </summary>
        public IReadOnlyList<VisioPage> Pages => _pages;

        /// <summary>
        /// Gets the masters currently registered on the document.
        /// </summary>
        public IReadOnlyCollection<VisioMaster> Masters => _registeredMasters.ToList().AsReadOnly();

        /// <summary>
        /// Gets or sets the theme applied to the document.
        /// </summary>
        public VisioTheme? Theme { get; set; }

        /// <summary>
        /// Gets or sets the title of the document.
        /// </summary>
        public string? Title { get; set; }

        /// <summary>
        /// Gets or sets the author of the document.
        /// </summary>
        public string? Author { get; set; }

        /// <summary>
        /// When true, shapes with a known <see cref="VisioShape.NameU"/> (e.g. "Rectangle")
        /// are automatically backed by a document master if none is assigned explicitly.
        /// Defaults to false to preserve current library behavior. Set to true to align
        /// with native Visio authoring semantics and asset templates.
        /// </summary>
        public bool UseMastersByDefault { get; set; } = false;

        /// <summary>
        /// When writing page shapes that reference a master, emit only delta cells required
        /// for the page instance (e.g., PinX/PinY and minimal style hints) instead of a full
        /// XForm block. This better matches Visio-authored files and the provided assets.
        /// Defaults to true.
        /// </summary>
        public bool WriteMasterDeltasOnly { get; set; } = true;

        /// <summary>
        /// Lists masters available in a given VSDX file (helper that proxies <see cref="VisioAssets.ListMasters"/>).
        /// </summary>
        public static IReadOnlyList<VisioAssets.MasterInfo> ListMastersIn(string vsdxPath) => VisioAssets.ListMasters(vsdxPath);

        /// <summary>
        /// Learns master names from a VSDX file and registers only the library-supported
        /// generated equivalents so shapes can reference them by NameU. This does not ingest,
        /// clone, or depend on the source VSDX as a runtime template.
        /// </summary>
        /// <param name="vsdxPath">Path to a Visio package that contains master metadata.</param>
        /// <param name="names">Optional filters matching master NameU, display name, relationship id, numeric id, or normalized slug.</param>
        public IReadOnlyList<VisioMaster> LearnMastersFromVsdx(string vsdxPath, IEnumerable<string>? names = null) {
            HashSet<string>? filter = names != null ? new HashSet<string>(names, StringComparer.OrdinalIgnoreCase) : null;
            IReadOnlyList<VisioAssets.MasterInfo> discovered = VisioAssets.ListMasters(vsdxPath);
            List<VisioMaster> imported = new();
            foreach (VisioAssets.MasterInfo masterInfo in discovered) {
                if (!VisioMasterIdentity.MatchesAny(masterInfo, filter)) {
                    continue;
                }

                if (TryEnsureBuiltinMaster(masterInfo.NameU, out VisioMaster? importedMaster) && importedMaster != null) {
                    imported.Add(importedMaster);
                }
            }

            return imported.AsReadOnly();
        }

        /// <summary>
        /// Imports master names from a VSDX file and registers only the library-supported
        /// generated equivalents so shapes can reference them by NameU. If <paramref name="names"/>
        /// is null or empty, all discoverable supported masters are registered.
        /// </summary>
        public void ImportMasters(string vsdxPath, IEnumerable<string>? names = null) {
            LearnMastersFromVsdx(vsdxPath, names);
        }

        /// <summary>
        /// Imports master names from a VSDX file and returns the registered generated masters
        /// that are natively implemented by the library.
        /// </summary>
        public IReadOnlyList<VisioMaster> ImportMastersAndGet(string vsdxPath, IEnumerable<string>? names = null) {
            return LearnMastersFromVsdx(vsdxPath, names);
        }

        /// <summary>
        /// Imports actual master definitions from a user-supplied Visio stencil, drawing, or template package.
        /// Unlike <see cref="LearnMastersFromVsdx"/>, this preserves the package master XML so shapes can use
        /// real external stencil artwork. Use this only for stencil packs the caller is allowed to embed.
        /// </summary>
        /// <param name="packagePath">Path to a `.vssx`, `.vsdx`, or `.vstx` package.</param>
        /// <param name="names">Optional filters matching master NameU, display name, relationship id, numeric id, or normalized slug.</param>
        public void ImportStencilMasters(string packagePath, IEnumerable<string>? names = null) {
            ImportStencilMastersAndGet(packagePath, names);
        }

        /// <summary>
        /// Imports actual master definitions from a user-supplied Visio stencil, drawing, or template package and returns the registered masters.
        /// </summary>
        /// <param name="packagePath">Path to a `.vssx`, `.vsdx`, or `.vstx` package.</param>
        /// <param name="names">Optional filters matching master NameU, display name, relationship id, numeric id, or normalized slug.</param>
        public IReadOnlyList<VisioMaster> ImportStencilMastersAndGet(string packagePath, IEnumerable<string>? names = null) {
            if (string.IsNullOrWhiteSpace(packagePath)) throw new ArgumentException("Package path cannot be null or whitespace.", nameof(packagePath));
            if (!File.Exists(packagePath)) throw new FileNotFoundException("Visio package was not found.", packagePath);

            ImportStencilPackageVisualContext(packagePath);
            IEnumerable<string>? resolvedNames = ResolvePackageMasterNameFilters(packagePath, names);
            IReadOnlyList<VisioAssets.MasterContent> contents = VisioAssets.LoadMasterContents(packagePath, resolvedNames);
            List<VisioMaster> imported = new();
            foreach (VisioAssets.MasterContent content in contents) {
                VisioShape shape = CreateImportedMasterShape(content);
                XDocument rawMasterXml = new(content.MasterXml);
                NormalizeImportedMasterRoot(rawMasterXml);
                VisioMaster master = new(content.Id, content.NameU, shape) {
                    RawMasterContentXml = rawMasterXml,
                    IsPackageBacked = true,
                    StencilSourcePackagePath = VisioStencilMetadata.NormalizePath(packagePath)
                };
                foreach (VisioAssets.MasterRelationshipContent relationship in content.Relationships) {
                    master.RawMasterRelationships.Add(relationship);
                }

                RegisterMaster(master);
                imported.Add(master);
            }

            return imported.AsReadOnly();
        }

        /// <summary>
        /// Imports the package-backed masters required by the provided stencil shapes.
        /// Shapes loaded from multiple `.vssx`, `.vstx`, or `.vsdx` packages are grouped by source package.
        /// </summary>
        /// <param name="shapes">Stencil shapes with <see cref="VisioStencilShape.SourcePackagePath"/> metadata.</param>
        public void ImportStencilMasters(IEnumerable<VisioStencilShape> shapes) {
            ImportStencilMastersAndGet(shapes);
        }

        /// <summary>
        /// Imports the package-backed masters required by the provided stencil shapes and returns the imported masters.
        /// </summary>
        /// <param name="shapes">Stencil shapes with <see cref="VisioStencilShape.SourcePackagePath"/> metadata.</param>
        public IReadOnlyList<VisioMaster> ImportStencilMastersAndGet(IEnumerable<VisioStencilShape> shapes) {
            if (shapes == null) throw new ArgumentNullException(nameof(shapes));

            List<VisioMaster> imported = new();
            foreach (IGrouping<string, VisioStencilShape> group in shapes
                .Where(shape => shape != null && !string.IsNullOrWhiteSpace(shape.SourcePackagePath))
                .GroupBy(shape => shape.SourcePackagePath!, StringComparer.OrdinalIgnoreCase)) {
                IReadOnlyList<VisioMaster> importedFromPackage = ImportStencilMastersAndGet(group.Key, group.Select(shape => shape.MasterNameU));
                imported.AddRange(importedFromPackage);
                foreach (VisioStencilShape shape in group) {
                    if (TryGetMaster(shape.MasterNameU, out VisioMaster? master) && master != null) {
                        VisioStencilMetadata.Apply(master, shape, catalogName: null);
                    }
                }
            }

            return imported.AsReadOnly();
        }

        /// <summary>
        /// Adds a new page to the document.
        /// </summary>
        /// <param name="name">Name of the page.</param>
        /// <param name="width">Page width.</param>
        /// <param name="height">Page height.</param>
        /// <param name="unit">Measurement unit for width and height.</param>
        /// <param name="id">Optional page identifier. If not specified, uses zero-based index.</param>
        public VisioPage AddPage(string name, double width = 8.26771653543307, double height = 11.69291338582677, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches, int? id = null) {
            if (id.HasValue && id.Value < 0) {
                throw new ArgumentOutOfRangeException(nameof(id), "Page id must be zero or greater.");
            }

            double widthInches = width.ToInches(unit);
            double heightInches = height.ToInches(unit);
            VisioPage page = new(name, widthInches, heightInches) { Id = id ?? GetNextPageId() };
            page.OwnerDocument = this;
            page.DefaultUnit = unit; // remember authoring unit for this page
            page.ScaleMeasurementUnit = unit;
            _pages.Add(page);
            return page;
        }

        /// <summary>
        /// Adds a Visio background page that can be reused by foreground pages.
        /// </summary>
        /// <param name="name">Name of the background page.</param>
        /// <param name="width">Page width.</param>
        /// <param name="height">Page height.</param>
        /// <param name="unit">Measurement unit for width and height.</param>
        /// <param name="id">Optional page identifier.</param>
        public VisioPage AddBackgroundPage(string name, double width = 8.26771653543307, double height = 11.69291338582677, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches, int? id = null) {
            VisioPage page = AddPage(name, width, height, unit, id);
            page.IsBackground = true;
            return page;
        }

        private int GetNextPageId() {
            HashSet<int> usedIds = new();
            foreach (VisioPage page in _pages) {
                if (page.Id >= 0) {
                    usedIds.Add(page.Id);
                }
            }

            int nextId = 0;
            while (usedIds.Contains(nextId)) {
                nextId++;
            }

            return nextId;
        }

        /// <summary>
        /// Requests Visio to relayout and reroute connectors when the document is opened.
        /// </summary>
        public void RequestRecalcOnOpen() {
            _requestRecalcOnOpen = true;
        }

        /// <summary>
        /// Creates a new <see cref="VisioDocument"/> with the given save path.
        /// </summary>
        /// <param name="path">Path where the document will be saved.</param>
        public static VisioDocument Create(string path) {
            return new VisioDocument { _filePath = path };
        }

        /// <summary>
        /// Creates a new <see cref="VisioDocument"/> that will be saved to the provided stream.
        /// </summary>
        /// <param name="stream">Destination stream for the VSDX package.</param>
        public static VisioDocument Create(Stream stream) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanWrite) throw new ArgumentException("Stream must be writable.", nameof(stream));
            if (!OfficeStreamWriter.CanReplaceContents(stream)) {
                throw new ArgumentException("Stream must support seeking when used as an associated destination.", nameof(stream));
            }
            return new VisioDocument { _sourceStream = stream };
        }

        private static IEnumerable<string>? ResolvePackageMasterNameFilters(string packagePath, IEnumerable<string>? names) {
            if (names == null) {
                return null;
            }

            HashSet<string> filter = new(names, StringComparer.OrdinalIgnoreCase);
            if (filter.Count == 0) {
                return null;
            }

            return VisioAssets.ListMasters(packagePath)
                .Where(master => VisioMasterIdentity.MatchesAny(master, filter))
                .Select(master => master.NameU)
                .ToArray();
        }

        private void ImportStencilPackageVisualContext(string packagePath) {
            VisioAssets.PackageVisualContext context = VisioAssets.LoadVisualContext(packagePath);
            XNamespace ns = VisioNamespace;

            XElement? documentRoot = context.DocumentXml?.Root;
            if (documentRoot != null) {
                ImportColors(documentRoot.Element(ns + "Colors"), ns);
                ImportFaceNames(documentRoot.Element(ns + "FaceNames"), ns);
                ImportStyleSheets(documentRoot.Element(ns + "StyleSheets"), ns);
            }

            if (Theme == null && context.ThemeXml?.Root != null) {
                Theme = new VisioTheme {
                    Name = context.ThemeXml.Root.Attribute("name")?.Value,
                    TemplateXml = new XDocument(context.ThemeXml)
                };
            }
        }

        private void ImportColors(XElement? colors, XNamespace ns) {
            if (colors == null) {
                return;
            }

            foreach (XAttribute attribute in colors.Attributes().Where(ShouldPreserveColorsAttribute)) {
                AddMissingAttribute(PreservedColorsAttributes, attribute);
            }

            HashSet<string> existingColorIndexes = new HashSet<string>(PreservedColorsElements
                .Where(element => string.Equals(element.Name.LocalName, "ColorEntry", StringComparison.OrdinalIgnoreCase))
                .Select(element => element.Attribute("IX")?.Value ?? string.Empty)
                .Where(value => !string.IsNullOrWhiteSpace(value)), StringComparer.OrdinalIgnoreCase);
            foreach (XElement element in colors.Elements().Where(ShouldPreserveColorsElement)) {
                string? colorIndex = element.Attribute("IX")?.Value;
                if (!string.IsNullOrWhiteSpace(colorIndex) && !existingColorIndexes.Add(colorIndex!)) {
                    continue;
                }

                PreservedColorsElements.Add(new XElement(element));
            }
        }

        private void ImportFaceNames(XElement? faceNames, XNamespace ns) {
            if (faceNames == null) {
                return;
            }

            foreach (XAttribute attribute in faceNames.Attributes().Where(ShouldPreserveFaceNamesAttribute)) {
                AddMissingAttribute(PreservedFaceNamesAttributes, attribute);
            }

            HashSet<string> existingFaces = new HashSet<string>(PreservedFaceNamesElements
                .Where(element => string.Equals(element.Name.LocalName, "FaceName", StringComparison.OrdinalIgnoreCase))
                .Select(element => element.Attribute("NameU")?.Value ?? element.Attribute("Name")?.Value ?? element.Attribute("ID")?.Value ?? string.Empty)
                .Where(value => !string.IsNullOrWhiteSpace(value)), StringComparer.OrdinalIgnoreCase);
            foreach (XElement element in faceNames.Elements().Where(ShouldPreserveFaceNamesElement)) {
                string faceKey = element.Attribute("NameU")?.Value ?? element.Attribute("Name")?.Value ?? element.Attribute("ID")?.Value ?? Guid.NewGuid().ToString("N");
                if (!existingFaces.Add(faceKey)) {
                    continue;
                }

                PreservedFaceNamesElements.Add(new XElement(element));
            }
        }

        private void ImportStyleSheets(XElement? styleSheets, XNamespace ns) {
            if (styleSheets == null) {
                return;
            }

            foreach (XAttribute attribute in styleSheets.Attributes().Where(ShouldPreserveStyleSheetsAttribute)) {
                AddMissingAttribute(PreservedStyleSheetsAttributes, attribute);
            }

            foreach (XElement element in styleSheets.Elements().Where(ShouldPreserveStyleSheetsElement)) {
                if (!PreservedStyleSheetsElements.Any(existing => XNode.DeepEquals(existing, element))) {
                    PreservedStyleSheetsElements.Add(new XElement(element));
                }
            }

            HashSet<string> existingAdditionalStyleIds = new HashSet<string>(PreservedAdditionalStyleSheets
                .Select(styleSheet => styleSheet.Attribute("ID")?.Value ?? string.Empty)
                .Where(value => !string.IsNullOrWhiteSpace(value)), StringComparer.Ordinal);
            foreach (XElement styleSheet in styleSheets.Elements(ns + "StyleSheet")) {
                string id = styleSheet.Attribute("ID")?.Value ?? string.Empty;
                if (string.IsNullOrWhiteSpace(id)) {
                    continue;
                }

                if (!IsGeneratedStyleSheet(id)) {
                    if (existingAdditionalStyleIds.Add(id)) {
                        PreservedAdditionalStyleSheets.Add(new XElement(styleSheet));
                    }

                    continue;
                }

                PreservedStyleSheetData preserved = GetOrCreatePreservedStyleSheet(this, id);
                foreach (XAttribute attribute in styleSheet.Attributes().Where(attribute => ShouldPreserveStyleSheetAttribute(attribute, id))) {
                    AddMissingAttribute(preserved.Attributes, attribute);
                }

                foreach (XElement element in styleSheet.Elements().Where(element => ShouldPreserveStyleSheetElement(element, id))) {
                    if (!preserved.ChildElements.Any(existing => XNode.DeepEquals(existing, element))) {
                        preserved.ChildElements.Add(new XElement(element));
                    }
                }
            }
        }

        private static void AddMissingAttribute(IList<XAttribute> attributes, XAttribute attribute) {
            if (!attributes.Any(existing => existing.Name == attribute.Name)) {
                attributes.Add(new XAttribute(attribute));
            }
        }

        private static void NormalizeImportedMasterRoot(XDocument masterXml) {
            XElement? rootShape = FindFirstMasterShape(masterXml);
            if (rootShape == null) {
                return;
            }

            if (TryReadImportedMasterCell(rootShape, "LocPinX", out double locPinX)) {
                SetImportedMasterCell(rootShape, "PinX", locPinX);
            } else if (TryReadImportedMasterCell(rootShape, "Width", out double width) && width > 0) {
                SetImportedMasterCell(rootShape, "PinX", width / 2D);
            }

            if (TryReadImportedMasterCell(rootShape, "LocPinY", out double locPinY)) {
                SetImportedMasterCell(rootShape, "PinY", locPinY);
            } else if (TryReadImportedMasterCell(rootShape, "Height", out double height) && height > 0) {
                SetImportedMasterCell(rootShape, "PinY", height / 2D);
            }
        }

        private static XElement? FindFirstMasterShape(XDocument masterXml) {
            XNamespace ns = VisioNamespace;
            XElement? shapeElement = masterXml.Root?
                .Element(ns + "Shapes")?
                .Elements(ns + "Shape")
                .FirstOrDefault();
            if (shapeElement != null) {
                return shapeElement;
            }

            return masterXml.Root?
                .Elements()
                .FirstOrDefault(element => string.Equals(element.Name.LocalName, "Shapes", StringComparison.OrdinalIgnoreCase))?
                .Elements()
                .FirstOrDefault(element => string.Equals(element.Name.LocalName, "Shape", StringComparison.OrdinalIgnoreCase));
        }

        private static void SetImportedMasterCell(XElement shapeElement, string name, double value) {
            XElement? cell = shapeElement
                .Elements()
                .FirstOrDefault(element =>
                    string.Equals(element.Name.LocalName, "Cell", StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(element.Attribute("N")?.Value, name, StringComparison.OrdinalIgnoreCase));
            XNamespace ns = shapeElement.Name.Namespace;
            if (cell == null) {
                shapeElement.AddFirst(new XElement(ns + "Cell", new XAttribute("N", name)));
                cell = shapeElement
                    .Elements()
                    .FirstOrDefault(element =>
                        string.Equals(element.Name.LocalName, "Cell", StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(element.Attribute("N")?.Value, name, StringComparison.OrdinalIgnoreCase));
            }

            if (cell != null) {
                cell.SetAttributeValue("V", value.ToString(CultureInfo.InvariantCulture));
            }
        }

        private static VisioShape CreateImportedMasterShape(VisioAssets.MasterContent content) {
            XElement? shapeElement = FindFirstMasterShape(content.MasterXml);

            string shapeId = shapeElement?.Attribute("ID")?.Value ?? "1";
            double width = TryReadImportedMasterCell(shapeElement, "Width", out double parsedWidth) && parsedWidth > 0 ? parsedWidth : 1D;
            double height = TryReadImportedMasterCell(shapeElement, "Height", out double parsedHeight) && parsedHeight > 0 ? parsedHeight : 1D;
            double pinX = TryReadImportedMasterCell(shapeElement, "PinX", out double parsedPinX) ? parsedPinX : width / 2D;
            double pinY = TryReadImportedMasterCell(shapeElement, "PinY", out double parsedPinY) ? parsedPinY : height / 2D;
            double locPinX = TryReadImportedMasterCell(shapeElement, "LocPinX", out double parsedLocPinX) ? parsedLocPinX : width / 2D;
            double locPinY = TryReadImportedMasterCell(shapeElement, "LocPinY", out double parsedLocPinY) ? parsedLocPinY : height / 2D;

            return new VisioShape(shapeId, pinX, pinY, width, height, string.Empty) {
                Name = shapeElement?.Attribute("Name")?.Value ?? content.NameU,
                NameU = shapeElement?.Attribute("NameU")?.Value ?? content.NameU,
                Type = shapeElement?.Attribute("Type")?.Value,
                LocPinX = locPinX,
                LocPinY = locPinY
            };
        }

        private static bool TryReadImportedMasterCell(XElement? shapeElement, string name, out double value) {
            string? rawValue = shapeElement?
                .Elements()
                .FirstOrDefault(element =>
                    string.Equals(element.Name.LocalName, "Cell", StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(element.Attribute("N")?.Value, name, StringComparison.OrdinalIgnoreCase))?
                .Attribute("V")?
                .Value;

            return double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out value) &&
                   !double.IsNaN(value) &&
                   !double.IsInfinity(value);
        }

        /// <summary>
        /// Returns the built-in master names that this library can generate natively.
        /// </summary>
        public static IReadOnlyCollection<string> SupportedBuiltinMasters =>
            BuiltinMasterDefinitions.Keys.ToList().AsReadOnly();

        /// <summary>
        /// Returns whether the library can generate a built-in master with the provided universal name.
        /// </summary>
        public static bool IsBuiltinMasterSupported(string? nameU) {
            return TryGetBuiltinMasterDefinition(nameU, out _);
        }

        /// <summary>
        /// Registers a master on the document so it can be reused by name.
        /// </summary>
        public VisioMaster RegisterMaster(VisioMaster master) {
            if (master == null) throw new ArgumentNullException(nameof(master));
            if (string.IsNullOrWhiteSpace(master.NameU)) throw new ArgumentException("Master NameU cannot be null or whitespace.", nameof(master));

            int existingIndex = _registeredMasters.FindIndex(existing =>
                string.Equals(existing.Id, master.Id, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(existing.NameU, master.NameU, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(existing.StencilSourcePackagePath, master.StencilSourcePackagePath, StringComparison.OrdinalIgnoreCase));
            if (existingIndex >= 0) {
                _registeredMasters[existingIndex] = master;
            } else {
                _registeredMasters.Add(master);
            }

            _builtinMasters[master.NameU] = master;
            return master;
        }

        /// <summary>
        /// Registers a master blueprint under a NameU and returns the registered master.
        /// </summary>
        public VisioMaster RegisterMaster(string nameU, VisioShape shape, string? id = null) {
            if (string.IsNullOrWhiteSpace(nameU)) throw new ArgumentException("Master NameU cannot be null or whitespace.", nameof(nameU));
            if (shape == null) throw new ArgumentNullException(nameof(shape));

            VisioMaster master = new(id ?? Guid.NewGuid().ToString("N"), nameU, shape);
            return RegisterMaster(master);
        }

        /// <summary>
        /// Attempts to find a registered master by NameU.
        /// </summary>
        public bool TryGetMaster(string nameU, out VisioMaster? master) {
            if (string.IsNullOrWhiteSpace(nameU)) {
                master = null;
                return false;
            }

            if (_builtinMasters.TryGetValue(nameU, out VisioMaster? existing)) {
                master = existing;
                return true;
            }

            master = null;
            return false;
        }

        /// <summary>
        /// Gets a registered master by NameU.
        /// </summary>
        public VisioMaster GetMaster(string nameU) {
            if (TryGetMaster(nameU, out VisioMaster? master) && master != null) {
                return master;
            }

            throw new KeyNotFoundException($"Master '{nameU}' is not registered on this document.");
        }

        /// <summary>
        /// Ensures a built-in master exists for a given NameU (e.g. Rectangle) and returns it.
        /// IDs are stable so that generated XML is deterministic and closer to assets.
        /// </summary>
        internal VisioMaster EnsureBuiltinMaster(string nameU) {
            if (_builtinMasters.TryGetValue(nameU, out var existing)) return existing;

            if (!TryGetBuiltinMasterDefinition(nameU, out BuiltinMasterDefinition? definition)) {
                VisioMaster fallbackMaster = new("10", nameU, CreateMasterBlueprint(nameU, null));
                return RegisterMaster(fallbackMaster);
            }

            VisioMaster builtInMaster = new(definition!.Id, nameU, CreateMasterBlueprint(nameU, definition));
            return RegisterMaster(builtInMaster);
        }

        internal bool TryEnsureBuiltinMaster(string nameU, out VisioMaster? master) {
            if (_builtinMasters.TryGetValue(nameU, out VisioMaster? existing)) {
                master = existing;
                return true;
            }

            if (!TryGetBuiltinMasterDefinition(nameU, out BuiltinMasterDefinition? _)) {
                master = null;
                return false;
            }

            master = EnsureBuiltinMaster(nameU);
            return true;
        }

        private static bool TryGetBuiltinMasterDefinition(string? nameU, out BuiltinMasterDefinition? definition) {
            if (nameU != null && BuiltinMasterDefinitions.TryGetValue(nameU, out BuiltinMasterDefinition? builtIn)) {
                definition = builtIn;
                return true;
            }

            definition = null;
            return false;
        }

        private static VisioShape CreateMasterBlueprint(string nameU, BuiltinMasterDefinition? definition) {
            VisioShape blueprint = new("1", 0.5, 0.5, 1, 1, string.Empty) {
                NameU = nameU
            };

            if (definition?.GeometryKind != BuiltinGeometryKind.DynamicConnector) {
                blueprint.ConnectionPoints.Add(new VisioConnectionPoint(0.0, 0.5, 1, 0));
                blueprint.ConnectionPoints.Add(new VisioConnectionPoint(1.0, 0.5, -1, 0));
                blueprint.ConnectionPoints.Add(new VisioConnectionPoint(0.5, 0.0, 0, 1));
                blueprint.ConnectionPoints.Add(new VisioConnectionPoint(0.5, 1.0, 0, -1));
            }

            return blueprint;
        }

        private static IReadOnlyDictionary<string, BuiltinMasterDefinition> CreateBuiltinMasterDefinitions() {
            Dictionary<string, BuiltinMasterDefinition> definitions = new(StringComparer.OrdinalIgnoreCase) {
                ["Rectangle"] = new("2", BuiltinGeometryKind.Rectangle, false, "Drag onto the page.", false, true, 2, "basic,shape,geometry,polygon,rectangle,right,angle,four-sided"),
                ["Ellipse"] = new("3", BuiltinGeometryKind.Ellipse, false, "Drag onto the page.", false, true, 2, "basic,shape,geometry,round,ellipse,oval"),
                ["Square"] = new("4", BuiltinGeometryKind.Rectangle, true, "Drag onto the page.", false, true, 2, "basic,shape,geometry,polygon,rectangle,right,angle,four-sided"),
                ["Circle"] = new("5", BuiltinGeometryKind.Ellipse, true, "Drag onto the page.", false, true, 2, "basic,shape,geometry,circular,round"),
                ["Dynamic connector"] = new("6", BuiltinGeometryKind.DynamicConnector, false, "This connector automatically routes between the shapes it connects.", true, false, 0, string.Empty, addConnectorLayer: true),
                ["Diamond"] = new("7", BuiltinGeometryKind.Diamond, false, "Drag onto the page.", false, true, 2, "basic,shape,geometry,decision,diamond,rhombus"),
                ["Triangle"] = new("8", BuiltinGeometryKind.Triangle, false, "Drag onto the page.", false, true, 2, "basic,shape,geometry,polygon,triangle"),
                ["Process"] = new("9", BuiltinGeometryKind.Rectangle, false, "Drag onto the page.", false, true, 2, "flowchart,process,step,task,operation"),
                ["Decision"] = new("11", BuiltinGeometryKind.Diamond, false, "Drag onto the page.", false, true, 2, "flowchart,decision,branch,diamond"),
                ["Data"] = new("12", BuiltinGeometryKind.Parallelogram, false, "Drag onto the page.", false, true, 2, "flowchart,data,input,output,parallelogram"),
                ["Preparation"] = new("13", BuiltinGeometryKind.Hexagon, false, "Drag onto the page.", false, true, 2, "flowchart,preparation,setup,hexagon"),
                ["Manual operation"] = new("14", BuiltinGeometryKind.Trapezoid, false, "Drag onto the page.", false, true, 2, "flowchart,manual,operation,trapezoid"),
                ["Off-page reference"] = new("15", BuiltinGeometryKind.OffPageReference, false, "Drag onto the page.", false, true, 2, "flowchart,reference,off-page,pentagon"),
                ["Parallelogram"] = new("16", BuiltinGeometryKind.Parallelogram, false, "Drag onto the page.", false, true, 2, "basic,shape,geometry,parallelogram,quadrilateral"),
                ["Hexagon"] = new("17", BuiltinGeometryKind.Hexagon, false, "Drag onto the page.", false, true, 2, "basic,shape,geometry,hexagon,polygon"),
                ["Trapezoid"] = new("18", BuiltinGeometryKind.Trapezoid, false, "Drag onto the page.", false, true, 2, "basic,shape,geometry,trapezoid,polygon"),
                ["Pentagon"] = new("19", BuiltinGeometryKind.Pentagon, false, "Drag onto the page.", false, true, 2, "basic,shape,geometry,pentagon,polygon"),
            };

            foreach (KeyValuePair<string, BuiltinMasterDefinition> entry in definitions) {
                entry.Value.BaseId = CreateDeterministicGuid("Base", entry.Key);
                entry.Value.UniqueId = CreateDeterministicGuid("Unique", entry.Key, entry.Value.Id);
            }

            return definitions;
        }

        private static string CreateDeterministicGuid(params string[] segments) {
            string payload = string.Join("|", segments);
            using MD5 md5 = MD5.Create();
            byte[] hash = md5.ComputeHash(Encoding.UTF8.GetBytes(payload));
            return new Guid(hash).ToString("B").ToUpperInvariant();
        }
    }
}
