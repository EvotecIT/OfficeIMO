using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    public partial class VisioDocument {

        private static void ApplyMasterReferences(VisioShape shape, XElement shapeElement, XNamespace ns, Dictionary<string, VisioMaster> masters, VisioMaster? inheritedMaster = null, VisioShape? inheritedMasterShape = null) {
            VisioMaster? effectiveMaster = inheritedMaster;
            VisioShape? effectiveMasterShape = inheritedMasterShape;

            string? masterIdAttr = shapeElement.Attribute("Master")?.Value;
            if (!string.IsNullOrEmpty(masterIdAttr) && masters.TryGetValue(masterIdAttr!, out VisioMaster? resolvedMaster)) {
                effectiveMaster = resolvedMaster;
                effectiveMasterShape = resolvedMaster.Shape;
            }

            if (effectiveMaster != null) {
                shape.Master = effectiveMaster;
                if (string.IsNullOrWhiteSpace(shape.NameU)) {
                    shape.NameU = effectiveMaster.NameU;
                }
            }

            if (!string.IsNullOrEmpty(shape.MasterShapeId) && effectiveMaster != null) {
                VisioShape? referencedMasterShape = effectiveMaster.Shape.FindDescendantById(shape.MasterShapeId!);
                if (referencedMasterShape != null) {
                    shape.MasterShape = referencedMasterShape;
                    effectiveMasterShape = referencedMasterShape;
                }
            }

            VisioShape? fallbackMasterShape = shape.MasterShape ?? effectiveMasterShape ?? effectiveMaster?.Shape;
            if (fallbackMasterShape != null) {
                if (!shape.HasExplicitWidth) {
                    shape.Width = fallbackMasterShape.Width;
                }
                if (!shape.HasExplicitHeight) {
                    shape.Height = fallbackMasterShape.Height;
                }
                if (!shape.HasExplicitLocPinX) {
                    shape.LocPinX = fallbackMasterShape.LocPinX;
                }
                if (!shape.HasExplicitLocPinY) {
                    shape.LocPinY = fallbackMasterShape.LocPinY;
                }
            }

            XElement? childShapes = shapeElement.Element(ns + "Shapes");
            if (childShapes != null && shape.Children.Count > 0) {
                List<XElement> childElements = childShapes.Elements(ns + "Shape").ToList();
                int count = Math.Min(childElements.Count, shape.Children.Count);
                for (int i = 0; i < count; i++) {
                    VisioShape? inheritedChildMasterShape = null;
                    if (fallbackMasterShape != null && i < fallbackMasterShape.Children.Count) {
                        inheritedChildMasterShape = fallbackMasterShape.Children[i];
                    }

                    ApplyMasterReferences(shape.Children[i], childElements[i], ns, masters, effectiveMaster, inheritedChildMasterShape ?? fallbackMasterShape);
                }
            }
        }

        private static void RegisterShapeHierarchy(VisioShape shape, Dictionary<string, VisioShape> shapeMap) {
            if (!string.IsNullOrEmpty(shape.PersistedId)) {
                shapeMap[shape.PersistedId!] = shape;
            }
            if (!shapeMap.ContainsKey(shape.Id)) {
                shapeMap[shape.Id] = shape;
            }
            foreach (VisioShape child in shape.Children) {
                RegisterShapeHierarchy(child, shapeMap);
            }
        }

        private static void HydrateContainerRelationships(VisioPage page, Dictionary<string, VisioShape> shapeMap) {
            foreach (VisioShape shape in page.Shapes) {
                HydrateContainerRelationships(shape, shapeMap);
            }
        }

        private static void HydrateContainerRelationships(VisioShape shape, Dictionary<string, VisioShape> shapeMap) {
            if (!string.IsNullOrWhiteSpace(shape.RelationshipsFormula)) {
                foreach ((int relationshipType, string sheetId) in ParseRelationshipDependencies(shape.RelationshipsFormula!)) {
                    if (!shapeMap.TryGetValue(sheetId, out VisioShape? relatedShape)) {
                        continue;
                    }

                    if (relationshipType == 1) {
                        AddUnique(shape.ContainerMemberIds, relatedShape.Id);
                        AddUnique(relatedShape.ContainerOwnerIds, shape.Id);
                    } else if (relationshipType == 4) {
                        AddUnique(shape.ContainerOwnerIds, relatedShape.Id);
                        AddUnique(relatedShape.ContainerMemberIds, shape.Id);
                    }
                }
            }

            foreach (VisioShape child in shape.Children) {
                HydrateContainerRelationships(child, shapeMap);
            }
        }

        private static IEnumerable<(int relationshipType, string sheetId)> ParseRelationshipDependencies(string formula) {
            foreach (Match match in Regex.Matches(formula, @"DEPENDSON\(\s*(\d+)\s*,\s*Sheet\.([^!]+)!SheetRef\(\)\s*\)", RegexOptions.IgnoreCase)) {
                if (int.TryParse(match.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int relationshipType)) {
                    string sheetId = match.Groups[2].Value.Trim();
                    if (!string.IsNullOrWhiteSpace(sheetId)) {
                        yield return (relationshipType, sheetId);
                    }
                }
            }
        }

        private static void AddUnique(IList<string> values, string value) {
            if (!values.Contains(value, StringComparer.OrdinalIgnoreCase)) {
                values.Add(value);
            }
        }
    }
}
