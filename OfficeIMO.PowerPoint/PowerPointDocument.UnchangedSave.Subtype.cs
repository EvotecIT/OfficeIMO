using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private static bool CanSaveCurrentSubtype(OfficeOpcEditablePackage package, string? renamedPart) {
            // The SDK clone writes only its relationship graph. Keep the same behavior for opaque or orphaned ZIP parts.
            var reachable = new HashSet<string>(OfficeAsciiIgnoreCaseComparer.Instance) { OfficeOpcEditablePackage.ContentTypesPartName };
            try {
                IReadOnlyList<OfficeOpcRelationship>? Relationships(string? owner) {
                    IReadOnlyList<OfficeOpcRelationship>? edges = ReadEditablePackageRelationships(package, owner, strictXml: true);
                    if (edges is { Count: > 0 }) {
                        string part = OfficeOpcPartNames.GetRelationshipsPartName(owner);
                        if (renamedPart != null && edges.Any(edge => !edge.IsExternal && edge.TargetPartName is { } target &&
                            OfficeAsciiIgnoreCaseComparer.Instance.Equals(PartIdentity(target), PartIdentity(renamedPart)) &&
                            (owner != null || edge.Type != OfficeOpcPackage.OfficeDocumentRelationshipType))) return null;
                        reachable.Add(PartIdentity(part));
                    }
                    return edges;
                }

                return TryVisitPackageGraph(Relationships, package.GetPartName, package.GetContentType, default,
                    edge => reachable.Add(PartIdentity(edge.Part)), validateForSave: true)
                    && package.PartNames.All(name => reachable.Contains(PartIdentity(name)));
            } catch (Exception exception) when (exception is InvalidDataException or System.Xml.XmlException or IOException
                or ArgumentException or NotSupportedException or InvalidOperationException or OverflowException) {
                return false;
            }
        }

        private static bool HasStandardRelationshipMarkup(OfficeXmlNode root) {
            if (root.HasAttributes || HasRelationshipText(root)) return false;
            foreach (OfficeXmlNode entry in root.Children()) {
                if (!entry.Is(OfficeXmlNamespaces.PackageRelationships, "Relationship") || entry.FirstChild.Exists ||
                    HasRelationshipText(entry)) return false;
                foreach (OfficeXmlAttribute attribute in entry.Attributes()) {
                    if (attribute.NamespaceUri.Length != 0 || attribute.LocalName is not ("Id" or "Type" or "Target" or "TargetMode")) {
                        return false;
                    }
                }
            }
            return true;
        }

        private static bool HasRelationshipText(OfficeXmlNode parent) {
            OfficeXmlTree tree = parent.Tree!;
            for (int child = tree.FirstChildOf(parent.Node); child != OfficeXmlTree.NoNode; child = tree.NextSiblingOf(child)) {
                if (tree.KindOf(child) == OfficeXmlTreeNodeKind.Text && !string.IsNullOrWhiteSpace(tree.GetInnerText(child))) return true;
            }
            return false;
        }

        // The SDK has fixed content types for these parts. Files with other part kinds use the existing clone path.
        private static bool MatchesSubtypeSavePart(string relationship, string type) {
            const string r = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix;
            const string ppt = "application/vnd.openxmlformats-officedocument.presentationml.";
            const string drawing = "application/vnd.openxmlformats-officedocument.drawingml.";
            const string msOffice = "http://schemas.microsoft.com/office/";
            switch (relationship) {
                case r + "officeDocument": return PresentationExtension(type) != null;
                case r + "slide": return type == PowerPointPackageModel.SlideContentType;
                case r + "slideLayout": return type == PowerPointPackageModel.SlideLayoutContentType;
                case r + "slideMaster": return type == PowerPointPackageModel.SlideMasterContentType;
                case r + "notesSlide": return type == NotesSlideContentType;
                case r + "notesMaster": return type == NotesMasterContentType;
                case r + "handoutMaster": return type == ppt + "handoutMaster+xml";
                case r + "presProps": return type == PowerPointPackageModel.PresentationPropertiesContentType;
                case ViewRelationshipType: return type == ViewContentType;
                case r + "tableStyles": return type == PowerPointPackageModel.TableStylesContentType;
                case r + "commentAuthors": return type == ppt + "commentAuthors+xml";
                case r + "comments": return type == ppt + "comments+xml";
                case r + "tags": return type == ppt + "tags+xml";
                case r + "theme": return type == PowerPointPackageModel.ThemeContentType;
                case r + "themeOverride": return type == "application/vnd.openxmlformats-officedocument.themeOverride+xml";
                case OfficeOpcPackage.CorePropertiesRelationshipType: return type == PowerPointPackageModel.CorePropertiesContentType;
                case ExtendedPropertiesRelationshipType: return type == PowerPointPackageModel.ExtendedPropertiesContentType;
                case r + "custom-properties": return type == "application/vnd.openxmlformats-officedocument.custom-properties+xml";
                case r + "image":
                case "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail":
                    return type.StartsWith("image/", StringComparison.Ordinal);
                case AudioDataRelationshipType: return type.StartsWith("audio/", StringComparison.Ordinal);
                case VideoDataRelationshipType: return type.StartsWith("video/", StringComparison.Ordinal);
                case MediaDataRelationshipType:
                    return type.StartsWith("audio/", StringComparison.Ordinal) || type.StartsWith("video/", StringComparison.Ordinal);
                case EmbeddedPackageRelationshipType:
                case EmbeddedObjectRelationshipType:
                    return true; // Embedded packages and OLE objects have variable content types.
                case r + "chart": return type == drawing + "chart+xml";
                case r + "diagramColors": return type == drawing + "diagramColors+xml";
                case r + "diagramData": return type == drawing + "diagramData+xml";
                case r + "diagramLayout": return type == drawing + "diagramLayout+xml";
                case r + "diagramQuickStyle": return type == drawing + "diagramStyle+xml";
                case msOffice + "2007/relationships/diagramDrawing": return type == "application/vnd.ms-office.drawingml.diagramDrawing+xml";
                case msOffice + "2011/relationships/chartStyle": return type == "application/vnd.ms-office.chartstyle+xml";
                case msOffice + "2011/relationships/chartColorStyle": return type == "application/vnd.ms-office.chartcolorstyle+xml";
                case VbaProjectRelationshipType: return type == VbaProjectContentType;
                case msOffice + "2006/relationships/vbaProjectSignature": return type == "application/vnd.ms-office.vbaProjectSignature";
                case msOffice + "2014/relationships/vbaProjectSignatureAgile":
                case msOffice + "2006/relationships/vbaProjectSignatureAgile":
                    return type == "application/vnd.ms-office.vbaProjectSignatureAgile";
                case msOffice + "2020/07/relationships/vbaProjectSignatureV3":
                    return type == "application/vnd.ms-office.vbaProjectSignatureV3";
                default: return false;
            }
        }
    }
}
