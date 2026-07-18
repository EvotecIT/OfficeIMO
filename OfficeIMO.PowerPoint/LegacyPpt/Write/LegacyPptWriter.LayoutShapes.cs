using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Dgm = DocumentFormat.OpenXml.Drawing.Diagrams;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        internal static IReadOnlyList<PowerPointShape> ReadSlideShapesForWrite(
            PowerPointSlide slide, out string? unsupportedReason) {
            if (slide == null) throw new ArgumentNullException(nameof(slide));
            unsupportedReason = null;
            PowerPointShape[] slideShapes = slide.Shapes.ToArray();
            if (!ShowsLayoutShapes(slide.SlidePart.Slide)) {
                return slideShapes;
            }

            SlideLayoutPart? layoutPart = slide.SlidePart.SlideLayoutPart;
            P.ShapeTree? tree = layoutPart?.SlideLayout?.CommonSlideData?.ShapeTree;
            if (layoutPart == null || tree == null) return slideShapes;

            var result = new List<PowerPointShape>(
                checked(tree.ChildElements.Count + slideShapes.Length));
            foreach (OpenXmlElement element in tree.ChildElements) {
                if (element is P.NonVisualGroupShapeProperties
                    or P.GroupShapeProperties) continue;
                PowerPointShape? shape = WrapShapeForWrite(element,
                    layoutPart);
                if (shape == null) {
                    unsupportedReason ??=
                        $"The slide layout contains '{element.LocalName}' content that is not yet materialized by the native binary writer.";
                    continue;
                }
                if (!IsPlaceholderOverridden(shape, slideShapes)) result.Add(shape);
            }
            result.AddRange(slideShapes);
            return result;
        }

        internal static IReadOnlyList<PowerPointShape>
            FlattenShapeTreeForWrite(IEnumerable<PowerPointShape> shapes,
                out string? unsupportedReason) {
            if (shapes == null) throw new ArgumentNullException(nameof(shapes));
            var result = new List<PowerPointShape>();
            unsupportedReason = null;
            foreach (PowerPointShape shape in shapes) {
                result.Add(shape);
                if (shape is not PowerPointGroupShape group) continue;
                IReadOnlyList<PowerPointShape> children =
                    ReadGroupChildrenForWrite(group,
                        out string? childReason);
                if (childReason != null) {
                    unsupportedReason = childReason;
                    return result;
                }
                IReadOnlyList<PowerPointShape> descendants =
                    FlattenShapeTreeForWrite(children, out childReason);
                if (childReason != null) {
                    unsupportedReason = childReason;
                    return result;
                }
                result.AddRange(descendants);
            }
            return result;
        }

        internal static IReadOnlyList<PowerPointShape>
            ReadGroupChildrenForWrite(PowerPointGroupShape group,
                out string? unsupportedReason) {
            if (group == null) throw new ArgumentNullException(nameof(group));
            unsupportedReason = null;
            if (group.OwnerPart == null) {
                unsupportedReason = "The group shape has no owning package part for its nested content.";
                return Array.Empty<PowerPointShape>();
            }
            var result = new List<PowerPointShape>();
            foreach (OpenXmlElement element in group.GroupShape.ChildElements) {
                if (element is P.NonVisualGroupShapeProperties
                    or P.GroupShapeProperties) continue;
                PowerPointShape? child = WrapShapeForWrite(element,
                    group.OwnerPart);
                if (child == null) {
                    unsupportedReason = $"The group contains '{element.LocalName}' content that is not yet encoded by the native binary writer.";
                    return result;
                }
                if (group.OwnerSlide != null) child.AttachTo(group.OwnerSlide);
                result.Add(child);
            }
            if (result.Count == 0) {
                unsupportedReason = "Binary PowerPoint groups must contain at least one drawable child.";
            }
            return result;
        }

        private static PowerPointShape? WrapShapeForWrite(
            OpenXmlElement element, OpenXmlPartContainer ownerPart) =>
            element switch {
                P.Shape shape when shape.TextBody != null =>
                    new PowerPointTextBox(shape, ownerPart),
                P.Shape shape => new PowerPointAutoShape(shape),
                P.ConnectionShape connection =>
                    new PowerPointConnectionShape(connection),
                P.Picture picture when ownerPart is SlidePart slidePart
                    && PowerPointMedia.TryGetMediaKind(picture,
                        out PowerPointMediaKind kind) =>
                    new PowerPointMedia(picture, slidePart, kind),
                P.Picture picture when PowerPointMedia.TryGetMediaKind(
                    picture, out _) => null,
                P.Picture picture => new PowerPointPicture(picture,
                    ownerPart),
                P.GroupShape nested => new PowerPointGroupShape(nested,
                    ownerPart),
                P.GraphicFrame frame when frame.Graphic?.GraphicData?
                    .GetFirstChild<A.Table>() != null =>
                    new PowerPointTable(frame, ownerPart as SlidePart),
                P.GraphicFrame frame when frame.Graphic?.GraphicData?
                    .GetFirstChild<C.ChartReference>() != null =>
                    new PowerPointChart(frame, ownerPart),
                P.GraphicFrame frame when frame.Graphic?.GraphicData?
                    .GetFirstChild<Dgm.RelationshipIds>() != null
                    && ownerPart is SlidePart slidePart =>
                    new PowerPointSmartArt(frame, slidePart),
                P.GraphicFrame frame when frame.Graphic?.GraphicData?
                    .GetFirstChild<P.OleObject>() != null
                    && ownerPart is SlidePart slidePart =>
                    new PowerPointOleObject(frame, slidePart),
                _ => null
            };

        private static bool ShowsLayoutShapes(P.Slide? slide) =>
            slide?.ShowMasterShapes?.Value != false;

        internal static bool IsLayoutShape(PowerPointShape shape) =>
            shape != null && shape.Element.Ancestors<P.SlideLayout>().Any();

        internal static bool FollowsMasterObjects(PowerPointSlide slide,
            bool layoutIsIndependentMaster = false) {
            if (slide == null) throw new ArgumentNullException(nameof(slide));
            return slide.SlidePart.Slide?.ShowMasterShapes?.Value != false
                && (layoutIsIndependentMaster
                    || slide.SlidePart.SlideLayoutPart?.SlideLayout?
                        .ShowMasterShapes?.Value != false);
        }

        private static bool IsPlaceholderOverridden(PowerPointShape inheritedShape,
            IReadOnlyList<PowerPointShape> slideShapes) {
            if (!TryGetPlaceholderSignature(inheritedShape,
                    out P.PlaceholderValues? inheritedType,
                    out uint? inheritedIndex)) return false;
            return slideShapes.Any(shape => TryGetPlaceholderSignature(shape,
                    out P.PlaceholderValues? slideType, out uint? slideIndex)
                && PlaceholderSignaturesMatch(inheritedType, inheritedIndex,
                    slideType, slideIndex));
        }

        private static bool TryGetPlaceholderSignature(PowerPointShape shape,
            out P.PlaceholderValues? type, out uint? index) {
            type = shape.ShapePlaceholderType;
            index = shape.ShapePlaceholderIndex;
            return type.HasValue || index.HasValue;
        }

        private static bool PlaceholderSignaturesMatch(P.PlaceholderValues? leftType,
            uint? leftIndex, P.PlaceholderValues? rightType, uint? rightIndex) {
            if (leftIndex.HasValue && rightIndex.HasValue) {
                return leftIndex.Value == rightIndex.Value
                    && (!leftType.HasValue || !rightType.HasValue
                        || leftType.Value == rightType.Value);
            }
            return leftType.HasValue && rightType.HasValue
                && leftType.Value == rightType.Value;
        }
    }
}
