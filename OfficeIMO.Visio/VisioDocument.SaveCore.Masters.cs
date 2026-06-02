using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio {
    public partial class VisioDocument {

        private Dictionary<string, VisioMaster> BuildEffectiveShapeMasterMap(VisioPage page) {
            Dictionary<string, VisioMaster> map = new(StringComparer.Ordinal);

            void Visit(VisioShape shape) {
                VisioMaster? effectiveMaster = ResolveEffectiveMaster(shape);
                if (effectiveMaster != null) {
                    map[shape.Id] = effectiveMaster;
                }

                foreach (VisioShape child in shape.Children) {
                    Visit(child);
                }
            }

            foreach (VisioShape shape in page.Shapes) {
                Visit(shape);
            }

            return map;
        }

        private static void AddMastersInShapeOrder(IEnumerable<VisioShape> shapes, IReadOnlyDictionary<string, VisioMaster> pageMasters, IList<VisioMaster> masterCandidates) {
            foreach (VisioShape shape in shapes) {
                if (pageMasters.TryGetValue(shape.Id, out VisioMaster? master)) {
                    masterCandidates.Add(master);
                }

                if (shape.Children.Count > 0) {
                    AddMastersInShapeOrder(shape.Children, pageMasters, masterCandidates);
                }
            }
        }

        private VisioMaster? ResolveEffectiveMaster(VisioShape shape) {
            if (shape.Master != null) {
                return shape.Master;
            }

            if (!UseMastersByDefault) {
                return null;
            }

            string shapeNameU = shape.NameU?.Trim() ?? string.Empty;
            if (shapeNameU.Length == 0) {
                return null;
            }

            return TryEnsureBuiltinMaster(shapeNameU, out VisioMaster? master) ? master : null;
        }

        private static VisioMaster GetMastersRootMetadataSource(IReadOnlyList<PackageMasterEntry> masters) {
            foreach (PackageMasterEntry entry in masters) {
                if (entry.Master.PreservedMastersRootAttributes.Count > 0 ||
                    entry.Master.PreservedMastersRootElements.Count > 0) {
                    return entry.Master;
                }
            }

            return masters[0].Master;
        }

        private static List<PackageMasterEntry> CreatePackageMasterEntries(IEnumerable<VisioMaster> masters) {
            List<PackageMasterEntry> entries = new();
            HashSet<int> usedIds = new();

            foreach (VisioMaster master in masters) {
                if (entries.Any(entry => ReferenceEquals(entry.Master, master))) {
                    continue;
                }

                int packageId;
                if (int.TryParse(master.Id, NumberStyles.Integer, CultureInfo.InvariantCulture, out int suggestedId) &&
                    suggestedId > 0 &&
                    usedIds.Add(suggestedId)) {
                    packageId = suggestedId;
                } else {
                    packageId = usedIds.Count == 0 ? 1 : usedIds.Max() + 1;
                    while (!usedIds.Add(packageId)) {
                        packageId++;
                    }
                }

                entries.Add(new PackageMasterEntry {
                    Master = master,
                    PackageId = packageId.ToString(CultureInfo.InvariantCulture),
                    PartNumber = entries.Count + 1
                });
            }

            return entries;
        }

        private static string GetPackageMasterId(IReadOnlyList<PackageMasterEntry> packageMasters, VisioMaster master) {
            foreach (PackageMasterEntry entry in packageMasters) {
                if (ReferenceEquals(entry.Master, master)) {
                    return entry.PackageId;
                }
            }

            throw new InvalidOperationException($"Master '{master.NameU}' is not registered in the package.");
        }

        private static VisioMaster? TryGetEffectiveMaster(IReadOnlyDictionary<string, VisioMaster> effectiveMasters, VisioShape shape) {
            return effectiveMasters.TryGetValue(shape.Id, out VisioMaster? master) ? master : null;
        }

        private static void WriteRawMasterRelationships(Package package, PackagePart masterPart, VisioMaster master, int masterPartNumber) {
            for (int i = 0; i < master.RawMasterRelationships.Count; i++) {
                VisioAssets.MasterRelationshipContent relationship = master.RawMasterRelationships[i];
                if (string.IsNullOrWhiteSpace(relationship.Id) || string.IsNullOrWhiteSpace(relationship.Type) || string.IsNullOrWhiteSpace(relationship.Target)) {
                    continue;
                }

                if (relationship.IsExternal || relationship.Data == null || relationship.Data.Length == 0) {
                    masterPart.CreateRelationship(new Uri(relationship.Target, UriKind.RelativeOrAbsolute), TargetMode.External, relationship.Type, relationship.Id);
                    continue;
                }

                string extension = string.IsNullOrWhiteSpace(relationship.Extension) ? ".bin" : relationship.Extension;
                if (!extension.StartsWith(".", StringComparison.Ordinal)) {
                    extension = "." + extension;
                }

                string mediaName = $"officeimo-master{masterPartNumber}-rel{i + 1}{extension}";
                Uri mediaUri = new($"/visio/media/{mediaName}", UriKind.Relative);
                if (!package.PartExists(mediaUri)) {
                    PackagePart mediaPart = package.CreatePart(mediaUri, string.IsNullOrWhiteSpace(relationship.ContentType) ? "application/octet-stream" : relationship.ContentType);
                    using Stream mediaStream = mediaPart.GetStream(FileMode.Create, FileAccess.Write);
                    mediaStream.Write(relationship.Data, 0, relationship.Data.Length);
                }

                masterPart.CreateRelationship(new Uri("../media/" + mediaName, UriKind.Relative), TargetMode.Internal, relationship.Type, relationship.Id);
            }
        }

        private static void WriteMasterGeometry(XmlWriter writer, string ns, string? masterNameU, double width, double height) {
            if (TryGetBuiltinMasterDefinition(masterNameU, out BuiltinMasterDefinition? definition) && definition != null) {
                 switch (definition.GeometryKind) {
                     case BuiltinGeometryKind.Ellipse:
                         WriteEllipseGeometry(writer, ns, width, height);
                         return;
                     case BuiltinGeometryKind.Diamond:
                         WriteDiamondGeometry(writer, ns, width, height);
                         return;
                     case BuiltinGeometryKind.Triangle:
                         WriteTriangleGeometry(writer, ns, width, height);
                         return;
                     case BuiltinGeometryKind.Pentagon:
                         WritePentagonGeometry(writer, ns, width, height);
                         return;
                     case BuiltinGeometryKind.Parallelogram:
                         WriteParallelogramGeometry(writer, ns, width, height);
                         return;
                     case BuiltinGeometryKind.Hexagon:
                         WriteHexagonGeometry(writer, ns, width, height);
                         return;
                     case BuiltinGeometryKind.Trapezoid:
                         WriteTrapezoidGeometry(writer, ns, width, height);
                         return;
                     case BuiltinGeometryKind.OffPageReference:
                         WriteOffPageReferenceGeometry(writer, ns, width, height);
                         return;
                 }
             }

            WriteRectangleGeometry(writer, ns, width, height);
        }

        private static void WriteMasterPageSheet(XmlWriter writer, string ns, VisioMaster master, BuiltinMasterDefinition? definition) {
            writer.WriteStartElement("PageSheet", ns);
            writer.WriteAttributeString("LineStyle", "0");
            writer.WriteAttributeString("FillStyle", "0");
            writer.WriteAttributeString("TextStyle", "0");
            foreach (XAttribute preservedAttribute in master.PreservedPageSheetAttributes) {
                writer.WriteAttributeString(
                    preservedAttribute.Name.LocalName,
                    preservedAttribute.Name.NamespaceName.Length == 0 ? null : preservedAttribute.Name.NamespaceName,
                    preservedAttribute.Value);
            }
            WritePageCell(writer, ns, "PageWidth", 3.937007874015748, "MM");
            WritePageCell(writer, ns, "PageHeight", 3.937007874015748, "MM");
            WritePageCell(writer, ns, "ShdwOffsetX", 0.1181102362204724);
            WritePageCell(writer, ns, "ShdwOffsetY", -0.1181102362204724);
            WritePageCell(writer, ns, "PageScale", 0.03937007874015748, "MM");
            WritePageCell(writer, ns, "DrawingScale", 0.03937007874015748, "MM");
            WritePageCell(writer, ns, "DrawingSizeType", definition?.GeometryKind == BuiltinGeometryKind.DynamicConnector ? 4 : 0);
            WritePageCell(writer, ns, "DrawingScaleType", 0);
            WritePageCell(writer, ns, "InhibitSnap", 0);
            WritePageCell(writer, ns, "PageLockReplace", 0, "BOOL");
            WritePageCell(writer, ns, "PageLockDuplicate", 0, "BOOL");
            WritePageCell(writer, ns, "UIVisibility", 0);
            WritePageCell(writer, ns, "ShdwType", 0);
            WritePageCell(writer, ns, "ShdwObliqueAngle", 0);
            WritePageCell(writer, ns, "ShdwScaleFactor", 1);
            WritePageCell(writer, ns, "DrawingResizeType", definition?.GeometryKind == BuiltinGeometryKind.DynamicConnector ? 0 : 1);
            WritePreservedElements(writer, master.PreservedPageSheetCells);
            string? shapeKeywords = definition?.ShapeKeywords;
            if (!string.IsNullOrWhiteSpace(shapeKeywords)) {
                WriteStringCell(writer, ns, "ShapeKeywords", shapeKeywords!);
            }
            bool hasPreservedLayerSection = master.PreservedPageSheetSections.Any(section =>
                string.Equals(section.Attribute("N")?.Value, "Layer", StringComparison.OrdinalIgnoreCase));
            bool hasPreservedUserSection = master.PreservedPageSheetSections.Any(section =>
                string.Equals(section.Attribute("N")?.Value, "User", StringComparison.OrdinalIgnoreCase));
            IReadOnlyList<VisioUserCell> masterUserCells = CreateMasterPageSheetUserCells(master);
            if (!hasPreservedUserSection) {
                WriteUserSection(writer, ns, masterUserCells.ToList());
            }
            if (definition?.AddConnectorLayer == true && !hasPreservedLayerSection) {
                writer.WriteStartElement("Section", ns);
                writer.WriteAttributeString("N", "Layer");
                writer.WriteStartElement("Row", ns);
                writer.WriteAttributeString("IX", "1");
                WriteStringCell(writer, ns, "Name", "Connector");
                WriteCell(writer, ns, "Color", 255);
                WriteCell(writer, ns, "Status", 0);
                WriteCell(writer, ns, "Visible", 1);
                WriteCell(writer, ns, "Print", 1);
                WriteCell(writer, ns, "Active", 0);
                WriteCell(writer, ns, "Lock", 0);
                WriteCell(writer, ns, "Snap", 1);
                WriteCell(writer, ns, "Glue", 1);
                WriteStringCell(writer, ns, "NameUniv", "Connector");
                WriteCell(writer, ns, "ColorTrans", 0);
                writer.WriteEndElement();
                writer.WriteEndElement();
            }
            WriteMasterPageSheetSections(writer, ns, master.PreservedPageSheetSections, masterUserCells);
            writer.WriteEndElement();
        }

        private static IReadOnlyList<VisioUserCell> CreateMasterPageSheetUserCells(VisioMaster master) {
            List<VisioUserCell> userCells = new();
            if (master.IsPackageBacked) {
                userCells.Add(new VisioUserCell("OfficeIMO.PackageBackedMaster", "1") {
                    Prompt = "OfficeIMO persisted package-backed stencil provenance"
                });
            }

            userCells.AddRange(VisioStencilMetadata.CreateMasterUserCells(master));
            return userCells.AsReadOnly();
        }

        private static XDocument MergeRawMasterMetadata(XDocument rawMasterContentXml, VisioMaster master) {
            XDocument merged = new(rawMasterContentXml);
            IReadOnlyList<VisioUserCell> masterUserCells = CreateMasterPageSheetUserCells(master);
            if (masterUserCells.Count == 0 || merged.Root == null) {
                return merged;
            }

            XNamespace ns = merged.Root.Name.Namespace;
            XElement? pageSheet = merged.Root.Elements()
                .FirstOrDefault(element => string.Equals(element.Name.LocalName, "PageSheet", StringComparison.OrdinalIgnoreCase));
            if (pageSheet == null) {
                pageSheet = new XElement(ns + "PageSheet");
                merged.Root.Add(pageSheet);
            }

            XElement? userSection = pageSheet.Elements()
                .FirstOrDefault(element =>
                    string.Equals(element.Name.LocalName, "Section", StringComparison.OrdinalIgnoreCase) &&
                    string.Equals((string?)element.Attribute("N"), "User", StringComparison.OrdinalIgnoreCase));
            if (userSection == null) {
                userSection = new XElement(ns + "Section", new XAttribute("N", "User"));
                pageSheet.Add(userSection);
            }

            HashSet<string> existingRows = new(
                userSection.Elements()
                    .Where(element => string.Equals(element.Name.LocalName, "Row", StringComparison.OrdinalIgnoreCase))
                    .Select(row => row.Attribute("N")?.Value)
                    .Where(name => !string.IsNullOrWhiteSpace(name))!,
                StringComparer.OrdinalIgnoreCase);
            foreach (VisioUserCell userCell in masterUserCells) {
                if (existingRows.Add(userCell.Name)) {
                    userSection.Add(CreateUserRowElement(ns.NamespaceName, userCell));
                }
            }

            return merged;
        }

        private static void WriteMasterPageSheetSections(XmlWriter writer, string ns, IEnumerable<XElement> sections, IReadOnlyList<VisioUserCell> masterUserCells) {
            foreach (XElement section in sections) {
                if (!string.Equals(section.Attribute("N")?.Value, "User", StringComparison.OrdinalIgnoreCase) ||
                    masterUserCells.Count == 0) {
                    section.WriteTo(writer);
                    continue;
                }

                XElement mergedSection = new(section);
                HashSet<string> existingRows = new(
                    mergedSection.Elements(XName.Get("Row", ns))
                        .Select(row => row.Attribute("N")?.Value)
                        .Where(name => !string.IsNullOrWhiteSpace(name))!,
                    StringComparer.OrdinalIgnoreCase);
                foreach (VisioUserCell userCell in masterUserCells) {
                    if (existingRows.Add(userCell.Name)) {
                        mergedSection.Add(CreateUserRowElement(ns, userCell));
                    }
                }

                mergedSection.WriteTo(writer);
            }
        }

        private static XElement CreateUserRowElement(string ns, VisioUserCell userCell) {
            XNamespace xns = ns;
            XElement row = new(xns + "Row", new XAttribute("N", userCell.Name));
            XElement value = new(xns + "Cell",
                new XAttribute("N", "Value"),
                new XAttribute("V", userCell.Value ?? string.Empty));
            if (!string.IsNullOrEmpty(userCell.Unit)) {
                value.Add(new XAttribute("U", userCell.Unit));
            }
            if (!string.IsNullOrEmpty(userCell.Formula)) {
                value.Add(new XAttribute("F", userCell.Formula));
            }
            row.Add(value);

            if (userCell.Prompt != null || !string.IsNullOrEmpty(userCell.PromptFormula)) {
                XElement prompt = new(xns + "Cell",
                    new XAttribute("N", "Prompt"),
                    new XAttribute("V", userCell.Prompt ?? string.Empty));
                if (!string.IsNullOrEmpty(userCell.PromptFormula)) {
                    prompt.Add(new XAttribute("F", userCell.PromptFormula));
                }
                row.Add(prompt);
            }

            return row;
        }

        private static void WriteMasterUserSection(XmlWriter writer, string ns) {
            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "User");
            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("N", "visVersion");
            WriteCell(writer, ns, "Value", 15);
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        private static void WriteMasterCharacterSection(XmlWriter writer, string ns) {
            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "Character");
            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("IX", "0");
            WriteCell(writer, ns, "Size", 0.1388888888888889, "PT", null);
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        private static void WriteDefaultTextBlock(XmlWriter writer, string ns, double width, double height) {
            WriteCell(writer, ns, "TxtPinX", width / 2.0, "MM", "Width*0.5");
            WriteCell(writer, ns, "TxtPinY", height / 2.0, "MM", "Height*0.5");
            WriteCell(writer, ns, "TxtWidth", width * 0.875, "MM", "Width*0.875");
            WriteCell(writer, ns, "TxtHeight", height * 0.75, "MM", "Height*0.75");
            WriteCell(writer, ns, "TxtLocPinX", width * 0.4375, "MM", "TxtWidth*0.5");
            WriteCell(writer, ns, "TxtLocPinY", height * 0.375, "MM", "TxtHeight*0.5");
            WriteCell(writer, ns, "TxtAngle", 0);
        }

        private static void WriteConnectorControlSection(XmlWriter writer, string ns, double height) {
            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "Control");
            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("N", "TextPosition");
            WriteCell(writer, ns, "X", 0);
            WriteCell(writer, ns, "Y", -height, null, "Controls.TextPosition.Y");
            WriteCell(writer, ns, "XDyn", 0, null, "Controls.TextPosition");
            WriteCell(writer, ns, "YDyn", -height, null, "Controls.TextPosition.Y");
            WriteCell(writer, ns, "XCon", 5);
            WriteCell(writer, ns, "YCon", 0);
            WriteCell(writer, ns, "CanGlue", 0);
            WriteStringCell(writer, ns, "Prompt", "Reposition Text");
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        private sealed class PackageMasterEntry {
            public VisioMaster Master { get; set; } = null!;
            public string PackageId { get; set; } = string.Empty;
            public int PartNumber { get; set; }
        }
    }
}
