using System;
using System.Collections.Generic;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents a Visio document containing pages.
    /// </summary>
    public partial class VisioDocument {
        private readonly List<VisioPage> _pages = new();
        private bool _requestRecalcOnOpen;
        private string? _filePath;
        private readonly Dictionary<string, VisioMaster> _builtinMasters = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, (string Id, System.Xml.Linq.XDocument Xml, System.Xml.Linq.XElement MasterElement)> _templateMasters = new(StringComparer.OrdinalIgnoreCase);

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

        /// <summary>
        /// Gets the collection of pages in the document.
        /// </summary>
        public IReadOnlyList<VisioPage> Pages => _pages;

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
        /// Imports masters from a VSDX file into this document's template catalog so that
        /// shapes can reference them by NameU when <see cref="UseMastersByDefault"/> is enabled.
        /// If <paramref name="names"/> is null or empty, all masters are imported.
        /// </summary>
        public void ImportMasters(string vsdxPath, IEnumerable<string>? names = null) {
            var packs = VisioAssets.LoadMasterContents(vsdxPath, names);
            foreach (var m in packs) {
                var bpTemplate = new VisioShape("1") { NameU = m.NameU, Width = 1, Height = 1 };
                var templated = new VisioMaster(m.Id, m.NameU, bpTemplate) {
                    TemplateXml = m.MasterXml,
                    TemplateMasterElement = new System.Xml.Linq.XElement(m.MasterElement)
                };
                _templateMasters[m.NameU] = (templated.Id, templated.TemplateXml!, templated.TemplateMasterElement!);
                _builtinMasters[m.NameU] = templated;
            }
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
            page.DefaultUnit = unit; // remember authoring unit for this page
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
        /// Loads masters from a template VSDX and makes them available when <see cref="EnsureBuiltinMaster"/>
        /// is called (mapped by NameU).
        /// </summary>
        /// <param name="vsdxPath">Path to a VSDX file that contains canonical masters.</param>
        public void UseMastersFromTemplate(string vsdxPath) {
            using var zip = System.IO.Compression.ZipFile.OpenRead(vsdxPath);
            var mastersList = zip.GetEntry("visio/masters/masters.xml");
            if (mastersList == null) return;
            using var mastersStream = mastersList.Open();
            var mastersDoc = System.Xml.Linq.XDocument.Load(mastersStream);
            var ns = System.Xml.Linq.XNamespace.Get(VisioNamespace);
            var relNs = System.Xml.Linq.XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            foreach (var m in mastersDoc.Root!.Elements(ns + "Master")) {
                string id = (string?)m.Attribute("ID") ?? string.Empty;
                string nameU = (string?)m.Attribute("NameU") ?? string.Empty;
                string relId = (string?)m.Element(ns + "Rel")?.Attribute(relNs + "id") ?? string.Empty;
                if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(nameU) || string.IsNullOrEmpty(relId)) continue;
                // Resolve relationship to master*.xml
                var mastersRels = zip.GetEntry("visio/masters/_rels/masters.xml.rels");
                if (mastersRels == null) continue;
                using var relsStream = mastersRels.Open();
                var relsDoc = System.Xml.Linq.XDocument.Load(relsStream);
                var rNs = System.Xml.Linq.XNamespace.Get("http://schemas.openxmlformats.org/package/2006/relationships");
                var rel = relsDoc.Root!.Elements(rNs + "Relationship").FirstOrDefault(e => (string?)e.Attribute("Id") == relId);
                if (rel == null) continue;
                string target = (string?)rel.Attribute("Target") ?? string.Empty;
                if (string.IsNullOrEmpty(target)) continue;
                string partPath = "visio/masters/" + target;
                var part = zip.GetEntry(partPath);
                if (part == null) continue;
                using var partStream = part.Open();
                var partDoc = System.Xml.Linq.XDocument.Load(partStream);
                // Capture the canonical <Master> element to mirror attributes/PageSheet/Icon
                var masterElem = m; // already an XElement from masters.xml
                _templateMasters[nameU] = (id, partDoc, masterElem);
            }
        }

        /// <summary>
        /// Ensures a built-in master exists for a given NameU (e.g. Rectangle) and returns it.
        /// IDs are stable so that generated XML is deterministic and closer to assets.
        /// </summary>
        internal VisioMaster EnsureBuiltinMaster(string nameU) {
            if (_builtinMasters.TryGetValue(nameU, out var existing)) return existing;

            // If template masters are available, prefer them to ensure exact fidelity
            if (_templateMasters.TryGetValue(nameU, out var t)) {
                var bpTemplate = new VisioShape("1") { NameU = nameU, Width = 1, Height = 1 };
                var templated = new VisioMaster(t.Id, nameU, bpTemplate) { TemplateXml = t.Xml, TemplateMasterElement = new System.Xml.Linq.XElement(t.MasterElement) };
                _builtinMasters[nameU] = templated;
                return templated;
            }

            // Stable IDs inspired by common basic shapes; Rectangle uses 2 to match asset sample.
            string id = nameU.Equals("Rectangle", StringComparison.OrdinalIgnoreCase) ? "2" : nameU switch {
                "Ellipse" => "3",
                "Square" => "4",
                "Circle" => "5",
                "Diamond" => "7",
                "Triangle" => "8",
                "Dynamic connector" => "6", // matches DrawingWithShapes.vsdx
                _ => "10" // fallback stable id bucket
            };

            // Minimal blueprint; MasterContents will be written from this shape (rectangle geometry by default).
            var blueprint = new VisioShape("1") { NameU = nameU, Width = 1, Height = 1 };
            // Provide default connection points for 2D shapes so connectors snap to sides.
            if (!nameU.Equals("Dynamic connector", StringComparison.OrdinalIgnoreCase)) {
                // Left, Right, Bottom, Top (Dir vectors point outward)
                blueprint.ConnectionPoints.Add(new VisioConnectionPoint(0.0, 0.5, 1, 0));
                blueprint.ConnectionPoints.Add(new VisioConnectionPoint(1.0, 0.5, -1, 0));
                blueprint.ConnectionPoints.Add(new VisioConnectionPoint(0.5, 0.0, 0, 1));
                blueprint.ConnectionPoints.Add(new VisioConnectionPoint(0.5, 1.0, 0, -1));
            }
            var master = new VisioMaster(id, nameU, blueprint);
            _builtinMasters[nameU] = master;
            return master;
        }
    }
}

