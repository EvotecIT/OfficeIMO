using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Visio;

// The container changes here; all model interpretation stays in VisioDocument's shared loader/writer.
internal static partial class VisioLegacyXmlCodec {
    internal static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";
    private static readonly XNamespace Modern = "http://schemas.microsoft.com/office/visio/2012/main";
    private static readonly XNamespace Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string RelationshipBase = "http://schemas.microsoft.com/visio/2010/relationships/";
    private static readonly string[] RootOrder = "DocumentProperties DocumentSettings Colors PrintSetup Fonts FaceNames StyleSheets DocumentSheet Masters Pages Windows EventList HeaderFooter SolutionXML".Split(' ');

    internal static void ValidateFamily(VisioPackageType type) {
        if (type != VisioPackageType.Drawing && type != VisioPackageType.Template && type != VisioPackageType.Stencil)
            throw new NotSupportedException("Legacy XML supports macro-free drawing, template and stencil families.");
    }

    internal static MemoryStream ToPackage(XDocument source, VisioPackageType type, VisioXmlConversionReport report) {
        XElement root = NormalizeLegacyRoot(source, report);
        if (root.Descendants(Legacy + "VBProjectData").Any()) throw new NotSupportedException("Legacy XML VBA projects are not supported.");
        XElement converted = ToModern(root, report);
        ResolveLegacyFontTable(converted, report);
        ResolveLegacyPalette(converted, report);
        CaptureLegacyCellMetadata(converted, report);
        ReportUnmodeledConnectors(converted, report);
        var result = new MemoryStream();
        int foreignResourceCount = 0;
        long foreignBytes = 0;
        try {
            using (Package package = Package.Open(result, FileMode.Create, FileAccess.ReadWrite)) {
                PackagePart document = AddPart(package, "/visio/document.xml", VisioPackageFormat.GetContentType(type));
                package.CreateRelationship(document.Uri, TargetMode.Internal, RelationshipBase + "document", "rId1");
                foreach (string family in new[] { "Pages", "Masters" }) {
                    XElement? container = converted.Element(Modern + family);
                    if (container == null && family == "Masters") continue;
                    container ??= new XElement(Modern + family);
                    if (container.Parent != null) container.Remove();
                    string lower = family.ToLowerInvariant();
                    string item = family == "Pages" ? "Page" : "Master";
                    var indexXml = new XElement(Modern + family, CopyAttributes(container));
                    PackagePart index = AddPart(package, "/visio/" + lower + "/" + lower + ".xml", "application/vnd.ms-visio." + lower + "+xml");
                    document.CreateRelationship(index.Uri, TargetMode.Internal, RelationshipBase + lower, "rId" + family);
                    int ordinal = 0;
                    foreach (XElement entry in container.Elements(Modern + item)) {
                        string relId = "rId" + (++ordinal).ToString(System.Globalization.CultureInfo.InvariantCulture);
                        // Text whitespace inherited from VisioDocument must remain significant
                        // after moving Shapes into a separate page or master package part.
                        var content = new XElement(Modern + (item + "Contents"), new XAttribute(XNamespace.Xml + "space", "preserve"));
                        foreach (XElement child in entry.Elements().Where(e => e.Name == Modern + "Shapes" || e.Name == Modern + "Connects").ToList()) {
                            child.Remove(); content.Add(child);
                        }
                        var reference = new XElement(entry);
                        reference.Add(new XElement(Modern + "Rel", new XAttribute(Relationships + "id", relId)));
                        indexXml.Add(reference);
                        PackagePart part = AddPart(package, "/visio/" + lower + "/" + item.ToLowerInvariant() + ordinal + ".xml", "application/vnd.ms-visio." + item.ToLowerInvariant() + "+xml");
                        index.CreateRelationship(part.Uri, TargetMode.Internal, RelationshipBase + item.ToLowerInvariant(), relId);
                        ExtractForeignData(content, part, report, ref foreignResourceCount, ref foreignBytes);
                        Write(part, content);
                    }
                    Write(index, indexXml);
                }
                XElement? properties = converted.Element(Modern + "DocumentProperties");
                package.PackageProperties.Title = (string?)properties?.Element(Modern + "Title");
                package.PackageProperties.Creator = (string?)properties?.Element(Modern + "Creator");
                Write(document, converted);
            }
            result.Position = 0; return result;
        } catch { result.Dispose(); throw; }
    }

    internal static XDocument FromPackage(Package package, VisioXmlConversionReport report) {
        PackageRelationship relationship = package.GetRelationshipsByType(RelationshipBase + "document").Single();
        PackagePart document = package.GetPart(PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), relationship.TargetUri));
        XElement modernRoot = Read(document);
        var handled = new HashSet<Uri> { document.Uri };
        foreach (string family in new[] { "Pages", "Masters" }) {
            string lower = family.ToLowerInvariant(), item = family == "Pages" ? "Page" : "Master";
            PackageRelationship? indexRel = document.GetRelationshipsByType(RelationshipBase + lower).SingleOrDefault();
            if (indexRel == null) continue;
            PackagePart index = package.GetPart(PackUriHelper.ResolvePartUri(document.Uri, indexRel.TargetUri)); handled.Add(index.Uri);
            XElement indexXml = Read(index);
            var container = new XElement(Modern + family, CopyAttributes(indexXml).Where(attribute => attribute.Name != XNamespace.Xml + "space"));
            foreach (XElement reference in indexXml.Elements(Modern + item)) {
                XElement entry = new XElement(reference);
                XElement? relElement = entry.Element(Modern + "Rel");
                string relId = (string?)relElement?.Attribute(Relationships + "id") ?? throw new InvalidDataException("Visio page/master reference has no relationship.");
                relElement!.Remove();
                PackagePart part = package.GetPart(PackUriHelper.ResolvePartUri(index.Uri, index.GetRelationship(relId).TargetUri)); handled.Add(part.Uri);
                entry.Add(InlineForeignData(part, handled, report).Elements());
                container.Add(entry);
            }
            modernRoot.Add(container);
        }
        VisioFontWireCodec fontWire = VisioFontWireCodec.ReadDocument(modernRoot);
        fontWire.DecodeCells(modernRoot, VisioNativeCellMetadata.Read(modernRoot.Elements()));
        RestoreLegacyCellMetadata(modernRoot);
        XElement root = FromModern(modernRoot, report);
        root.SetAttributeValue(XNamespace.Xml + "space", "preserve");
        XElement properties = root.Element(Legacy + "DocumentProperties") ?? new XElement(Legacy + "DocumentProperties");
        if (properties.Parent == null) root.Add(properties);
        properties.SetElementValue(Legacy + "Title", package.PackageProperties.Title);
        properties.SetElementValue(Legacy + "Creator", package.PackageProperties.Creator);
        // These optional tables need no representation when empty. Some independent VDX readers
        // consume subsequent pages while searching for an end tag on self-closing tables.
        foreach (XElement empty in root.Elements().Where(e =>
                     (e.Name == Legacy + "Colors" || e.Name == Legacy + "FaceNames" || e.Name == Legacy + "StyleSheets") &&
                     !e.HasElements && !e.HasAttributes).ToList()) empty.Remove();
        foreach (PackagePart part in package.GetParts()) {
            string path = part.Uri.OriginalString;
            if (handled.Contains(part.Uri) || path.EndsWith(".rels", StringComparison.OrdinalIgnoreCase) || path.StartsWith("/docProps/", StringComparison.OrdinalIgnoreCase)) continue;
            if (part.ContentType == "application/vnd.ms-visio.windows+xml") {
                // The shared writer generates a window when none was loaded; it is not source document content.
                // Legacy window state, when present on import, is already preserved on the root.
                continue;
            }
            report.Add("VDX_PACKAGE_PART", "Package content has no legacy XML mapping: " + path, location: path);
        }
        foreach (XElement child in root.Elements().ToList()) {
            if (child.Name.Namespace == Legacy && !RootOrder.Contains(child.Name.LocalName)) {
                report.Add("VDX_DOCUMENT_ELEMENT", "Document element has no legacy mapping: " + child.Name.LocalName); child.Remove();
            }
        }
        // The root uses a sequence in DatadiagramML, unlike most ShapeSheet containers.
        var ordered = root.Elements().OrderBy(e => { int i = Array.IndexOf(RootOrder, e.Name.LocalName); return i < 0 ? int.MaxValue : i; }).ToList();
        root.ReplaceNodes(ordered);
        root.SetAttributeValue("version", "11.0");
        return new XDocument(new XDeclaration("1.0", "utf-8", null), root);
    }

    private static void ReportUnmodeledConnectors(XElement root, VisioXmlConversionReport report) {
        foreach (XElement page in root.Descendants().Where(element => element.Name == Modern + "Page" || element.Name == Modern + "Master")) {
            var connects = page.Element(Modern + "Connects")?.Elements(Modern + "Connect").ToList() ?? new List<XElement>();
            foreach (XElement shape in page.Element(Modern + "Shapes")?.Elements(Modern + "Shape") ?? Enumerable.Empty<XElement>()) {
                if (!shape.Elements(Modern + "Cell").Any(cell => (string?)cell.Attribute("N") == "OneD" && (string?)cell.Attribute("V") == "1")) continue;
                string? id = (string?)shape.Attribute("ID");
                var endpoints = connects.Where(connect => (string?)connect.Attribute("FromSheet") == id).ToList();
                bool ReadableEndpoint(string prefix) {
                    XElement? connect = endpoints.FirstOrDefault(c => (string?)c.Attribute("FromCell") == prefix + "X");
                    if (connect != null) return page.Descendants(Modern + "Shape").Any(target => (string?)target.Attribute("ID") == (string?)connect.Attribute("ToSheet"));
                    return new[] { prefix + "X", prefix + "Y" }.All(name =>
                        double.TryParse((string?)shape.Elements(Modern + "Cell").FirstOrDefault(c => (string?)c.Attribute("N") == name)?.Attribute("V"),
                            System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double number) &&
                        !double.IsNaN(number) && !double.IsInfinity(number));
                }
                if (!ReadableEndpoint("Begin") || !ReadableEndpoint("End"))
                    report.Add("VDX_CONNECTOR_PROFILE", "A connector with missing endpoint coordinates or unresolved shape attachments is preserved as native XML outside the editable connector model.",
                        OfficeConversionLossKind.Approximation, id);
            }
        }
    }

    private static IEnumerable<XAttribute> CopyAttributes(XElement element) => element.Attributes().Where(a => !a.IsNamespaceDeclaration).Select(a => new XAttribute(a));
    private static void ResolveLegacyPalette(XElement root, VisioXmlConversionReport report) {
        var colors = root.Element(Modern + "Colors")?.Elements(Modern + "ColorEntry")
            .Where(e => e.Attribute("IX") != null && e.Attribute("RGB") != null)
            .GroupBy(e => (string)e.Attribute("IX")!, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => (string)group.First().Attribute("RGB")!, StringComparer.Ordinal)
            ?? new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (XElement cell in root.Descendants(Modern + "Cell")) {
            string? name = (string?)cell.Attribute("N"), value = (string?)cell.Attribute("V");
            if (name == null || value == null) continue;
            string? section = (string?)cell.Parent?.Parent?.Attribute("N");
            // Layer.Color is an index in the typed model; retain it and the root Colors table.
            if (name == "Color" && section == "Layer") {
                if (!int.TryParse(value, out _)) report.Add("VDX_LAYER_COLOR", "The shared layer model cannot interpret this non-indexed color.", OfficeConversionLossKind.Approximation);
                continue;
            }
            // Indexed text backgrounds have a different encoding from ordinary color cells.
            if (name == "TextBkgnd" && int.TryParse(value, out int backgroundIndex)) {
                if (backgroundIndex == 0 || backgroundIndex == 255 || backgroundIndex == 1 || backgroundIndex == 2 ||
                    backgroundIndex > 0 && backgroundIndex <= 24 && colors.ContainsKey((backgroundIndex - 1).ToString(System.Globalization.CultureInfo.InvariantCulture))) continue;
                report.Add("VDX_TEXT_BACKGROUND", "The indexed text background has no resolvable palette color; its native cell is preserved while rendering uses the shared fallback color.", OfficeConversionLossKind.Approximation);
                continue;
            }
            bool rgbCell = name == "LineColor" || name == "FillForegnd" || name == "FillBkgnd" ||
                name == "ShdwForegnd" || name == "ShdwBkgnd" || (name == "Color" && section == "Character");
            if (!rgbCell) continue;
            if (colors.TryGetValue(value, out string? rgb)) cell.SetAttributeValue("V", rgb);
            else if (int.TryParse(value, out int index) && index > 1)
                report.Add("VDX_COLOR_PALETTE", "Palette index " + value + " has no explicit color entry; the model uses its fallback color.", OfficeConversionLossKind.Approximation);
        }
    }
    private static PackagePart AddPart(Package package, string path, string contentType) => package.CreatePart(new Uri(path, UriKind.Relative), contentType);
    private static void Write(PackagePart part, XElement root) {
        using Stream stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using XmlWriter writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = new System.Text.UTF8Encoding(false) });
        root.Save(writer);
    }
    private static XElement Read(PackagePart part) {
        using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
        return XElement.Load(reader, LoadOptions.PreserveWhitespace);
    }
}
