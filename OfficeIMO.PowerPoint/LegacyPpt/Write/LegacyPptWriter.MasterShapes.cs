using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        internal static IReadOnlyList<PowerPointShape> ReadMasterShapesForWrite(
            SlideMasterPart masterPart, out string? unsupportedReason) {
            if (masterPart == null) throw new ArgumentNullException(nameof(masterPart));
            return ReadMasterShapesForWrite(masterPart,
                masterPart.SlideMaster?.CommonSlideData?.ShapeTree,
                out unsupportedReason);
        }

        private static IReadOnlyList<PowerPointShape> ReadMasterShapesForWrite(
            OpenXmlPartContainer ownerPart, P.ShapeTree? tree,
            out string? unsupportedReason) {
            unsupportedReason = null;
            if (tree == null) return Array.Empty<PowerPointShape>();
            var shapes = new List<PowerPointShape>(tree.ChildElements.Count);
            foreach (OpenXmlElement element in tree.ChildElements) {
                if (element is P.NonVisualGroupShapeProperties
                    or P.GroupShapeProperties) continue;
                PowerPointShape? shape = WrapInheritedShape(element, ownerPart);
                if (shape == null) {
                    unsupportedReason ??=
                        $"The slide master contains '{element.LocalName}' content that is not yet encoded by the native binary writer.";
                    continue;
                }
                shapes.Add(shape);
            }
            return shapes;
        }

        private static uint ReadDrawingId(LegacyPptRecord prototype) {
            LegacyPptRecord? drawing = prototype.DescendantsAndSelf()
                .FirstOrDefault(record => record.Type == OfficeArtDg);
            if (drawing == null || drawing.Instance == 0) {
                throw new InvalidDataException(
                    "The embedded binary PowerPoint master has no drawing identifier.");
            }
            return drawing.Instance;
        }
    }
}
