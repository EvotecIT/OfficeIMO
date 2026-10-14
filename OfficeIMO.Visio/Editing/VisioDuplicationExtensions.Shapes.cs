using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    public static partial class VisioDuplicationExtensions {
        private static VisioShape CloneShape(
            VisioShape source,
            IdAllocator ids,
            double offsetX,
            double offsetY,
            bool applyOffset,
            Dictionary<VisioShape, VisioShape> shapeMap,
            Dictionary<VisioConnectionPoint, VisioConnectionPoint> connectionPointMap,
            VisioShapeDuplicationOptions? options = null,
            VisioDocument? sourceDocument = null) {
            VisioShape clone = new(ResolveShapeCloneId(source, ids, options), source.PinX + (applyOffset ? offsetX : 0D), source.PinY + (applyOffset ? offsetY : 0D), source.Width, source.Height, source.Text ?? string.Empty) {
                Name = source.Name,
                NameU = source.NameU,
                Type = source.Type,
                Master = source.Master,
                MasterShapeId = source.MasterShapeId,
                MasterShape = source.MasterShape,
                LineWeight = source.LineWeight,
                LocPinX = source.LocPinX,
                LocPinY = source.LocPinY,
                HasExplicitLocPinX = source.HasExplicitLocPinX,
                HasExplicitLocPinY = source.HasExplicitLocPinY,
                Angle = source.Angle,
                LineColor = source.LineColor,
                FillColor = source.FillColor,
                LinePattern = source.LinePattern,
                FillPattern = source.FillPattern,
                PlacementStyle = source.PlacementStyle,
                PlacementFlip = source.PlacementFlip,
                PlowCode = source.PlowCode,
                AllowPlacementOnTop = source.AllowPlacementOnTop,
                AllowHorizontalConnectorRoutingThrough = source.AllowHorizontalConnectorRoutingThrough,
                AllowVerticalConnectorRoutingThrough = source.AllowVerticalConnectorRoutingThrough,
                CanSplitShapes = source.CanSplitShapes,
                CanBeSplit = source.CanBeSplit,
                RelationshipsValue = source.RelationshipsValue,
                RelationshipsFormula = source.RelationshipsFormula,
                TextStyle = source.TextStyle?.Clone(),
                PreservedTextElement = source.PreservedTextElement == null ? null : new XElement(source.PreservedTextElement),
                PreservedTextValue = source.PreservedTextValue,
                HasInheritedText = source.HasInheritedText,
                HasModeledCharSection = source.HasModeledCharSection,
                HasModeledParaSection = source.HasModeledParaSection,
                CharacterSectionSource = source.CharacterSectionSource?.Clone(),
                ParagraphSectionSource = source.ParagraphSectionSource?.Clone(),
                ShapeDataSectionName = source.ShapeDataSectionName,
                NativeStyleReferences = source.NativeStyleReferences,
                NativeFontScope = source.NativeFontScope,
                NativeCellMetadata = source.NativeCellMetadata?.Clone()
            };

            CopyStringSet(source.LayerNames, clone.LayerNames);
            clone.NativeLayerMembership = source.NativeLayerMembership?.Clone();
            CopyStringList(source.ContainerMemberIds, clone.ContainerMemberIds);
            CopyStringList(source.ContainerOwnerIds, clone.ContainerOwnerIds);
            CopyDictionary(source.Data, clone.Data);
            CopyConnectionPoints(source, clone, connectionPointMap);
            CopyHyperlinks(source.Hyperlinks, clone.Hyperlinks, VisioHyperlinkRowNames.Inherited(source, sourceDocument));
            CopyUserCells(source.UserCells, clone.UserCells);
            CopyShapeData(source.ShapeData, clone.ShapeData, clone.Data);
            CopyProtection(source.Protection, clone.Protection);
            CopyElements(source.PreservedGeometrySections, clone.PreservedGeometrySections);
            CopyElements(source.PreservedCellElements, clone.PreservedCellElements);
            CopyElements(source.PreservedNonGeometrySections, clone.PreservedNonGeometrySections);
            CopyElements(source.PreservedDataRows, clone.PreservedDataRows);

            clone.ForeignResources.AddRange(source.ForeignResources);
            if (source.ForeignResources.Count > 0) {
                // Preserve the complete sequence: selecting the ordered writer requires
                // the sibling raw cells/sections as well as the ForeignData reference.
                foreach (VisioShape.PreservedShapeChildEntry entry in source.PreservedShapeChildren)
                    clone.PreservedShapeChildren.Add(entry.RawElement != null
                        ? new VisioShape.PreservedShapeChildEntry(entry.RawElement)
                        : new VisioShape.PreservedShapeChildEntry(entry.Token!));
            }

            shapeMap[source] = clone;

            foreach (VisioShape child in source.Children) {
                clone.Children.Add(CloneShape(child, ids, offsetX, offsetY, applyOffset: false, shapeMap, connectionPointMap, options, sourceDocument));
            }

            return clone;
        }

    }
}
