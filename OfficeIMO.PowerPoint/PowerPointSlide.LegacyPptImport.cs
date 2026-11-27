using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        // Names of connectors the binary import added while the SDK held the slide, which the shape list does not track.
        private HashSet<string>? _legacyConnectorNames;

        /// <summary>
        /// Adds a connection shape with a preset geometry to the slide's shape tree, named and numbered as the slide names
        /// new connectors, and returns its element. On the package the slide tracks the shape; while the SDK holds the
        /// slide the element is added to its tree and the slide reads it again when it is next loaded.
        /// </summary>
        internal OfficeXmlNode AddLegacyConnectionShape(OfficePresetShapeType shapeType, long left,
            long top, long width, long height) {
            if (TryGetEditablePackageShapeTree(out OfficeXmlTree packageSlide, out int packageTree)) {
                uint packageId = AllocateShapeId();
                string packageName = GenerateUniqueName("Connector");
                OfficeXmlNode added = new OfficeXmlNode(packageSlide, packageTree).Append(
                    CreateLegacyConnectionShape(new OfficeXmlNode(packageSlide, packageTree), packageId, packageName,
                        shapeType, left, top, width, height));
                OwnerDocument!.MarkPackageEdited();
                TrackShape(new PowerPointConnectionShape(this, packageSlide, added.Node));
                return added;
            }

            OfficeXmlNode tree = PartRef.Root.Path(LegacyPptXml.P, "cSld", "spTree");
            if (tree.IsNone) throw new InvalidDataException("The projected slide has no shape tree.");
            uint id = AllocateShapeId();
            string name = GenerateLegacyConnectorName();
            return tree.Append(CreateLegacyConnectionShape(tree, id, name, shapeType, left, top, width, height));
        }

        // The slide's unique name for a new connector, counting connectors added while the shape list could not track them.
        private string GenerateLegacyConnectorName() {
            var names = new HashSet<string>(StringComparer.Ordinal);
            foreach (PowerPointShape shape in ShapeList) {
                if (shape.Name is string existing) names.Add(existing);
            }
            _legacyConnectorNames ??= new HashSet<string>(StringComparer.Ordinal);
            int index = 1;
            string name;
            do {
                name = "Connector " + index.ToString(CultureInfo.InvariantCulture);
                index++;
            } while (names.Contains(name) || _legacyConnectorNames.Contains(name));
            _legacyConnectorNames.Add(name);
            return name;
        }

        private static OfficeXmlNode CreateLegacyConnectionShape(OfficeXmlNode context, uint id, string name,
            OfficePresetShapeType shapeType, long left, long top, long width, long height) {
            OfficeXmlNode connector = context.CreateElement(LegacyPptXml.P, "cxnSp");
            OfficeXmlNode nonVisual = connector.AppendNew(LegacyPptXml.P, "nvCxnSpPr");
            OfficeXmlNode properties = nonVisual.AppendNew(LegacyPptXml.P, "cNvPr");
            properties.SetAttribute("id", id.ToString(CultureInfo.InvariantCulture));
            properties.SetAttribute("name", name);
            nonVisual.AppendNew(LegacyPptXml.P, "cNvCxnSpPr");
            nonVisual.AppendNew(LegacyPptXml.P, "nvPr");
            OfficeXmlNode shapeProperties = connector.AppendNew(LegacyPptXml.P, "spPr");
            OfficeXmlNode transform = shapeProperties.AppendNew(LegacyPptXml.A, "xfrm");
            OfficeXmlNode offset = transform.AppendNew(LegacyPptXml.A, "off");
            offset.SetAttribute("x", left.ToString(CultureInfo.InvariantCulture));
            offset.SetAttribute("y", top.ToString(CultureInfo.InvariantCulture));
            OfficeXmlNode extents = transform.AppendNew(LegacyPptXml.A, "ext");
            extents.SetAttribute("cx", width.ToString(CultureInfo.InvariantCulture));
            extents.SetAttribute("cy", height.ToString(CultureInfo.InvariantCulture));
            OfficeXmlNode geometry = shapeProperties.AppendNew(LegacyPptXml.A, "prstGeom");
            geometry.SetAttribute("prst", shapeType.ToOpenXmlToken());
            geometry.AppendNew(LegacyPptXml.A, "avLst");
            return connector;
        }
    }
}
