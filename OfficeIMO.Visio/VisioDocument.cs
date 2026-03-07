using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;

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

        private readonly List<VisioPage> _pages = new();
        private bool _requestRecalcOnOpen;
        private string? _filePath;
        private Stream? _sourceStream;
        private readonly Dictionary<string, VisioMaster> _builtinMasters = new(StringComparer.OrdinalIgnoreCase);
        private static readonly IReadOnlyDictionary<string, BuiltinMasterDefinition> BuiltinMasterDefinitions = CreateBuiltinMasterDefinitions();

        private const string DocumentRelationshipType = "http://schemas.microsoft.com/visio/2010/relationships/document";
        private const string DocumentContentType = "application/vnd.ms-visio.drawing.main+xml";
        private const string VisioNamespace = "http://schemas.microsoft.com/office/visio/2012/main";
        private const string ThemeRelationshipType = "http://schemas.microsoft.com/visio/2010/relationships/theme";
        private const string ThemeContentType = "application/vnd.ms-visio.theme+xml";
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
        /// Gets the masters currently registered on the document, keyed by <see cref="VisioMaster.NameU"/>.
        /// </summary>
        public IReadOnlyCollection<VisioMaster> Masters => _builtinMasters.Values.ToList().AsReadOnly();

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
        /// Imports master names from a VSDX file and registers only the library-supported
        /// generated equivalents so shapes can reference them by NameU. If <paramref name="names"/>
        /// is null or empty, all discoverable supported masters are registered.
        /// </summary>
        public void ImportMasters(string vsdxPath, IEnumerable<string>? names = null) {
            ImportMastersAndGet(vsdxPath, names);
        }

        /// <summary>
        /// Imports master names from a VSDX file and returns the registered generated masters
        /// that are natively implemented by the library.
        /// </summary>
        public IReadOnlyList<VisioMaster> ImportMastersAndGet(string vsdxPath, IEnumerable<string>? names = null) {
            HashSet<string>? filter = names != null ? new HashSet<string>(names, StringComparer.OrdinalIgnoreCase) : null;
            IReadOnlyList<VisioAssets.MasterInfo> discovered = VisioAssets.ListMasters(vsdxPath);
            List<VisioMaster> imported = new();
            foreach (VisioAssets.MasterInfo masterInfo in discovered) {
                if (filter != null && !filter.Contains(masterInfo.NameU)) {
                    continue;
                }

                if (TryEnsureBuiltinMaster(masterInfo.NameU, out VisioMaster? importedMaster) && importedMaster != null) {
                    imported.Add(importedMaster);
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
            double widthInches = width.ToInches(unit);
            double heightInches = height.ToInches(unit);
            VisioPage page = new(name, widthInches, heightInches) { Id = id ?? _pages.Count };
            page.OwnerDocument = this;
            page.DefaultUnit = unit; // remember authoring unit for this page
            page.ScaleMeasurementUnit = unit;
            _pages.Add(page);
            return page;
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
            return new VisioDocument { _sourceStream = stream };
        }

        /// <summary>
        /// Learns the available master names from a template VSDX and registers generated
        /// library equivalents for those names.
        /// </summary>
        /// <param name="vsdxPath">Path to a VSDX file that contains canonical masters.</param>
        public void UseMastersFromTemplate(string vsdxPath) {
            ImportMasters(vsdxPath);
        }

        /// <summary>
        /// Returns the built-in master names that this library can generate natively.
        /// </summary>
        public static IReadOnlyCollection<string> SupportedBuiltinMasters =>
            BuiltinMasterDefinitions.Keys.ToList().AsReadOnly();

        /// <summary>
        /// Registers a master on the document so it can be reused by name.
        /// </summary>
        public VisioMaster RegisterMaster(VisioMaster master) {
            if (master == null) throw new ArgumentNullException(nameof(master));
            if (string.IsNullOrWhiteSpace(master.NameU)) throw new ArgumentException("Master NameU cannot be null or whitespace.", nameof(master));

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
                _builtinMasters[nameU] = fallbackMaster;
                return fallbackMaster;
            }

            VisioMaster builtInMaster = new(definition!.Id, nameU, CreateMasterBlueprint(nameU, definition));
            _builtinMasters[nameU] = builtInMaster;
            return builtInMaster;
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
            VisioShape blueprint = new("1") {
                NameU = nameU,
                Width = 1,
                Height = 1
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

