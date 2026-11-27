using System.Xml.Linq;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// The diagram parts of a SmartArt diagram, reached through the relationships of its slide part wherever the part
    /// is held. Their content is read and written as bytes, so a part keeps the exact markup it was saved with until an
    /// edit replaces it.
    /// </summary>
    public partial class PowerPointSmartArt {
        internal const string DiagramDataRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "diagramData";
        internal const string DiagramLayoutRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "diagramLayout";
        internal const string DiagramQuickStyleRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "diagramQuickStyle";
        internal const string DiagramColorsRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "diagramColors";
        internal const string DiagramDrawingRelationshipType = "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing";

        // A diagram built over an SDK frame names its slide part; one read from the package belongs to its slide's part.
        private readonly PowerPointPartRef _initialSlidePart;

        /// <summary>The slide part the diagram's relationships belong to.</summary>
        private PowerPointPartRef HostPartRef => _initialSlidePart.IsNone ? OwnerSlide!.PartRef : _initialSlidePart;

        /// <summary>The frame's <c>dgm:relIds</c> at its place in the graphic data, as the data part is found.</summary>
        private OfficeXmlNode GraphicDataRelationshipIds => XmlElement
            .Child(OfficeXmlNamespaces.DrawingML, "graphic")
            .Child(OfficeXmlNamespaces.DrawingML, "graphicData")
            .Child(OfficeXmlNamespaces.Diagram, "relIds");

        /// <summary>The single <c>dgm:relIds</c> anywhere in the frame, or none; more than one throws.</summary>
        private OfficeXmlNode SingleRelationshipIds() =>
            XmlElement.Descendants(OfficeXmlNamespaces.Diagram, "relIds").SingleOrDefault();

        /// <summary>The part a relationship of the slide targets when the relationship is of the given type.</summary>
        private bool TryGetDiagramPart(string? relationshipId, string relationshipType, out PowerPointPartRef part) {
            part = default;
            if (string.IsNullOrEmpty(relationshipId)
                || !HostPartRef.TryGetRelationship(relationshipId!, out PowerPointRelationship relationship)
                || relationship.IsExternal || relationship.TargetPart.IsNone
                || !string.Equals(OfficeXmlNamespaces.NormalizeRelationshipType(relationship.Type), relationshipType, StringComparison.Ordinal)) {
                return false;
            }
            part = relationship.TargetPart;
            return true;
        }

        /// <summary>The part's bytes as they would be saved now.</summary>
        private static byte[] ReadPartBytes(PowerPointPartRef part) {
            PowerPointPartRef current = part.Current;
            if (current.IsSdk) return ReadSdkPartBytes(current.SdkPart!);
            OfficeOpcEditablePackage package = current.Document?.ReadablePackage
                ?? throw new InvalidOperationException("The SmartArt diagram part can't be read.");
            return package.ReadPartBytes(current.Name);
        }

        /// <summary>Replaces the part's content, in the package while the document edits it there, through the SDK otherwise.</summary>
        private static void WritePartBytes(PowerPointPartRef part, byte[] bytes) {
            PowerPointPartRef current = part.Current;
            if (!current.IsSdk && current.Document is { } document) {
                if (document.EditablePackage is { } package) {
                    package.ReplacePart(current.Name, bytes, bytes.Length);
                    current.MarkEdited();
                    return;
                }
            }
            // A package that can't be edited in place is edited through the SDK, as the model edits it.
            WriteSdkPartBytes(OpenSdkPart(part) ?? throw new InvalidOperationException("The SmartArt diagram part can't be edited."), bytes);
        }

        /// <summary>
        /// The markup of the part's root element as the part holds it now: its tree in the package, or the SDK's element,
        /// which also holds content set in memory and not yet saved. Null when the part has no root element.
        /// </summary>
        private static string? ReadRootMarkup(PowerPointPartRef part) {
            OfficeXmlNode root = part.Root;
            return root.IsNone ? null : System.Text.Encoding.UTF8.GetString(root.OuterXml());
        }

        private static XDocument LoadPartXml(PowerPointPartRef part, LoadOptions options = LoadOptions.None) {
            PowerPointPartRef current = part.Current;
            if (current.IsSdk) return LoadSdkPartXml(current.SdkPart!, options);
            using var stream = new MemoryStream(ReadPartBytes(current), writable: false);
            return PowerPointXmlReader.LoadPackagePartXml(stream, options);
        }

        /// <summary>Writes the document as the part's content: into the package's part, or straight into the SDK's.</summary>
        private static void WritePartXml(PowerPointPartRef part, XDocument document) {
            PowerPointPartRef current = part.Current;
            if (!current.IsSdk && current.Document?.EditablePackage != null) {
                WritePartBytes(current, SerializeXml(document));
                return;
            }
            SaveSdkPartXml(OpenSdkPart(part) ?? throw new InvalidOperationException("The SmartArt diagram part can't be edited."), document);
        }

        private static byte[] SerializeXml(XDocument document) {
            using var buffer = new MemoryStream();
            document.Save(buffer);
            return buffer.ToArray();
        }
    }
}
