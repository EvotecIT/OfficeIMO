using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Threading;

namespace OfficeIMO.Visio;

internal static partial class VisioLegacyXmlCodec {
    private static bool IsSheet(string name) => name == "Shape" || name == "PageSheet" || name == "StyleSheet" || name == "DocumentSheet";

    private static XElement ToModern(XElement source, VisioXmlConversionReport report, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (source.Name.Namespace != Legacy) return CloneImportedElement(source, cancellationToken);
        var target = new XElement(Modern + source.Name.LocalName, CopyAttributes(source));
        if (source.Name.LocalName == "Connect" && source.Attribute("FromCell") == null) {
            string? part = (string?)source.Attribute("FromPart");
            if (part == "9" || part == "12") target.SetAttributeValue("FromCell", part == "9" ? "BeginX" : "EndX");
            else report.Add("VDX_CONNECT_ENDPOINT", "Connection endpoint has no supported FromCell or FromPart mapping.",
                OfficeConversionLossKind.Approximation, (string?)source.Attribute("FromSheet"));
        }
        if (source.Name.LocalName == "Shape" && source.Element(Legacy + "XForm1D") != null)
            target.Add(new XElement(Modern + "Cell", new XAttribute("N", "OneD"), new XAttribute("V", "1")));
        if (!IsSheet(source.Name.LocalName)) {
            foreach (XNode node in source.Nodes()) {
                cancellationToken.ThrowIfCancellationRequested();
                target.Add(node is XElement element ? ToModern(element, report, cancellationToken) : CloneNode(node));
            }
            return target;
        }
        foreach (XElement child in source.Elements()) {
            cancellationToken.ThrowIfCancellationRequested();
            string name = child.Name.LocalName;
            if (child.Name.Namespace != Legacy) { target.Add(CloneImportedElement(child, cancellationToken)); continue; }
            if (SingletonRows.ContainsKey(name)) {
                foreach (XElement cell in child.Elements()) target.Add(ToCell(cell, cancellationToken));
                if (child.HasAttributes) report.Add("VDX_SINGLETON_ATTRIBUTES", "Singleton row attributes are not mapped.", location: name);
            } else if (IndexedRows.TryGetValue(name, out string? sectionName)) {
                XElement? section = target.Elements(Modern + "Section").FirstOrDefault(e => (string?)e.Attribute("N") == sectionName);
                if (section == null) { section = new XElement(Modern + "Section", new XAttribute("N", sectionName)); target.Add(section); }
                var row = new XElement(Modern + "Row", CopyRowAttributes(child, true, NamedRows.Contains(name)));
                row.Add(child.Elements().Select(cell => ToCell(cell, cancellationToken))); section.Add(row);
            } else if (name == "Geom") {
                var section = new XElement(Modern + "Section", new XAttribute("N", "Geometry"), CopyAttributes(child));
                foreach (XElement entry in child.Elements()) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (entry.Name.LocalName.StartsWith("No", StringComparison.Ordinal) && !entry.HasElements) section.Add(ToCell(entry, cancellationToken));
                    else section.Add(new XElement(Modern + "Row", new XAttribute("T", entry.Name.LocalName), CopyRowAttributes(entry, true), entry.Elements().Select(cell => ToCell(cell, cancellationToken))));
                }
                target.Add(section);
            } else if (name == "ForeignData") {
                target.Add(new XElement(Modern + "ForeignData", CopyAttributes(child), child.Value));
            } else if (name == "Tabs" || name == "ConnectionABCD") {
                // Preserve the native fragment rather than translating it into an incorrect modern section.
                target.Add(CloneImportedElement(child, cancellationToken));
                report.Add("VDX_PRESERVED_ROW", "Native row is preserved without model interpretation.", OfficeConversionLossKind.Approximation, name);
            } else target.Add(ToModern(child, report, cancellationToken));
        }
        return target;
    }

    private static XElement FromModern(XElement source, VisioXmlConversionReport report) {
        if (source.Name.Namespace != Modern) return new XElement(source);
        var target = new XElement(Legacy + source.Name.LocalName, CopyAttributes(source));
        if (source.Name.LocalName == "StyleSheet" && target.Attribute("BasedOn") is XAttribute basedOn) {
            foreach (string style in new[] { "LineStyle", "FillStyle", "TextStyle" }) {
                if (target.Attribute(style) == null) target.SetAttributeValue(style, basedOn.Value);
            }
            basedOn.Remove();
        }
        if (source.Name.LocalName == "Master") {
            foreach (string name in new[] { "IsCustomNameU", "IsCustomName", "MasterType" }) {
                if (target.Attribute(name) is not XAttribute attribute) continue;
                report.Add("VDX_MASTER_ATTRIBUTE", "Master attribute has no legacy mapping: " + name,
                    OfficeConversionLossKind.Approximation, (string?)source.Attribute("ID"));
                attribute.Remove();
            }
        }
        if (!IsSheet(source.Name.LocalName)) {
            foreach (XNode node in source.Nodes()) {
                if (node is XElement element) {
                    if (source.Name.LocalName == "DocumentSettings" && element.Name.LocalName == "RelayoutAndRerouteUponOpen") {
                        if (element.Value != "0") report.Add("VDX_RECALC", "Recalculate-on-open setting is not carried to legacy XML.");
                        continue;
                    }
                    target.Add(FromModern(element, report));
                } else target.Add(CloneNode(node));
            }
            return target;
        }
        foreach (XElement child in source.Elements()) {
            if (child.Name == Modern + "Cell") {
                string name = (string?)child.Attribute("N") ?? string.Empty;
                if (name == "OneD" && ((string?)child.Attribute("V") == "0" || (string?)child.Attribute("V") == "1")) continue; // Legacy identity comes from XForm1D.
                if (!CellRows.TryGetValue(name, out string? rowName)) {
                    if ((name == "PageLockReplace" || name == "PageLockDuplicate" || name == "DrawingResizeType") &&
                        (string?)child.Attribute("V") == "0" && child.Attribute("F") == null) continue;
                    report.Add("VDX_CELL", "Cell has no legacy mapping: " + name, location: (string?)source.Attribute("ID")); continue;
                }
                XElement? row = target.Element(Legacy + rowName);
                if (row == null) { row = new XElement(Legacy + rowName); target.Add(row); }
                row.Add(FromCell(child, name));
            } else if (child.Name == Modern + "Section") {
                string name = (string?)child.Attribute("N") ?? string.Empty;
                if (name == "Geometry") {
                    var geometry = new XElement(Legacy + "Geom", CopyAttributes(child).Where(a => a.Name != "N"));
                    foreach (XElement entry in child.Elements()) {
                        if (entry.Name == Modern + "Cell") AddGeometryCell(geometry, entry, report);
                        else if (entry.Name == Modern + "Row") {
                            string type = (string?)entry.Attribute("T") ?? "LineTo";
                            if (type == "Geometry") {
                                foreach (XElement cell in entry.Elements(Modern + "Cell")) AddGeometryCell(geometry, cell, report);
                                continue;
                            }
                            if (!GeometryRows.Contains(type)) { report.Add("VDX_GEOMETRY", "Geometry row has no legacy mapping: " + type); continue; }
                            geometry.Add(new XElement(Legacy + type, CopyRowAttributes(entry, false), entry.Elements(Modern + "Cell").Select(cell => FromCell(cell, (string?)cell.Attribute("N") ?? "Unknown"))));
                        }
                    }
                    var indices = new HashSet<string>(geometry.Elements().Where(e => GeometryRows.Contains(e.Name.LocalName))
                        .Select(e => (string?)e.Attribute("IX")).Where(value => value != null)!);
                    int nextIndex = 1;
                    foreach (XElement row in geometry.Elements().Where(e => GeometryRows.Contains(e.Name.LocalName) && e.Attribute("IX") == null)) {
                        while (indices.Contains(nextIndex.ToString(System.Globalization.CultureInfo.InvariantCulture))) nextIndex++;
                        string index = nextIndex++.ToString(System.Globalization.CultureInfo.InvariantCulture);
                        row.SetAttributeValue("IX", index); indices.Add(index);
                    }
                    target.Add(geometry);
                } else {
                    if ((string?)child.Attribute("Del") is "1" or "true")
                        report.Add("VDX_SECTION_DELETION", "Collection-level deletion has no legacy XML section equivalent: " + name);
                    string? legacyName = IndexedRows.FirstOrDefault(pair => pair.Value == name || pair.Key == name).Key;
                    if (legacyName == null) { report.Add("VDX_SECTION", "Section has no legacy mapping: " + name); continue; }
                    foreach (XElement row in child.Elements(Modern + "Row"))
                        target.Add(new XElement(Legacy + legacyName, CopyRowAttributes(row, false, NamedRows.Contains(legacyName)),
                            row.Elements(Modern + "Cell").Select(cell => FromCell(cell, (string?)cell.Attribute("N") ?? "Unknown"))));
                }
            } else target.Add(FromModern(child, report));
        }
        // ShapeSheet content precedes the nested Shapes extension in DatadiagramML.
        foreach (XElement nested in target.Elements(Legacy + "Shapes").ToList()) { nested.Remove(); target.Add(nested); }
        return target;
    }

    private static readonly HashSet<string> NamedRows = new("User Prop Hyperlink SmartTagDef".Split(' '), StringComparer.Ordinal);
    private static readonly HashSet<string> GeometryRows = new("MoveTo LineTo ArcTo InfiniteLine Ellipse EllipticalArcTo SplineStart SplineKnot PolylineTo NURBSTo".Split(' '), StringComparer.Ordinal);
    private static void AddGeometryCell(XElement geometry, XElement cell, VisioXmlConversionReport report) {
        string name = (string?)cell.Attribute("N") ?? string.Empty;
        if (name == "NoQuickDrag" && (string?)cell.Attribute("V") == "0") return;
        if (name != "NoFill" && name != "NoLine" && name != "NoShow" && name != "NoSnap") {
            report.Add("VDX_GEOMETRY_SETTING", "Geometry setting has no legacy mapping: " + name); return;
        }
        geometry.Add(FromCell(cell, name));
    }
    private static XElement ToCell(XElement source, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        var cell = new XElement(Modern + "Cell", new XAttribute("N", source.Name.LocalName), new XAttribute("V", source.Value));
        string? error = (string?)source.Attribute("Err");
        string? preservedError = error != null && !ModernErrors.Contains(error) ? error : null;
        foreach (XAttribute attribute in source.Attributes().Where(a => !a.IsNamespaceDeclaration && a.Name != "V" && !(a.Name == "Err" && preservedError != null)))
            cell.Add(new XAttribute(attribute.Name == "Unit" ? "U" : attribute.Name == "Err" ? "E" : attribute.Name, attribute.Value));
        string? condition = (string?)source.Attribute("V");
        if (condition != null || preservedError != null) cell.AddAnnotation(new LegacyCellState(condition, preservedError));
        return cell;
    }
    private static XElement FromCell(XElement source, string name) {
        var cell = new XElement(Legacy + name, (string?)source.Attribute("V") ?? string.Empty);
        foreach (XAttribute attribute in source.Attributes().Where(a => !a.IsNamespaceDeclaration && a.Name != "N" && a.Name != "V"))
            cell.Add(new XAttribute(attribute.Name == "U" ? "Unit" : attribute.Name == "E" ? "Err" : attribute.Name, attribute.Value));
        if (source.Annotation<LegacyCellState>() is LegacyCellState state) {
            cell.SetAttributeValue("V", state.Condition);
            if (state.Error != null) cell.SetAttributeValue("Err", state.Error);
        }
        return cell;
    }
    private static IEnumerable<XAttribute> CopyRowAttributes(XElement source, bool toModern, bool namedRow = false) {
        foreach (XAttribute attribute in CopyAttributes(source)) {
            if (!toModern && attribute.Name == "T") continue;
            XName name = toModern && attribute.Name == "NameU" ? "N" : !toModern && attribute.Name == "N" ? "NameU" : attribute.Name;
            if (namedRow && toModern && name == "ID") name = "IX";
            if (namedRow && !toModern && name == "IX") name = "ID";
            yield return new XAttribute(name, attribute.Value);
        }
        if (toModern && source.Attribute("NameU") == null && source.Attribute("Name") is XAttribute localizedName)
            yield return new XAttribute("N", localizedName.Value);
    }
    private static XNode CloneNode(XNode node) => node is XCData cdata ? new XCData(cdata.Value)
        : node is XText text ? new XText(text.Value)
        : node is XComment comment ? new XComment(comment.Value)
        : node is XProcessingInstruction instruction ? new XProcessingInstruction(instruction.Target, instruction.Data)
        : throw new NotSupportedException("Unsupported XML node: " + node.NodeType);
}
