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

        private static void WriteRemainingModeledShapeCells(
            XmlWriter writer,
            string ns,
            VisioShape shape,
            bool isGroup,
            double width,
            double height,
            double locPinX,
            double locPinY,
            ISet<string> emittedTokens,
            IReadOnlyDictionary<string, string> persistedIds,
            IReadOnlyDictionary<string, int> layerIndexes) {
            foreach (string cellName in ShapeModeledCellOrder) {
                string token = $"Cell:{cellName}";
                if (emittedTokens.Add(token)) {
                    TryWriteModeledShapeCell(writer, ns, shape, cellName, isGroup, width, height, locPinX, locPinY, persistedIds, layerIndexes);
                }
            }
        }

        private static void WriteShapeLayoutCells(XmlWriter writer, string ns, VisioShape shape) {
            if (shape.PlacementStyle.HasValue) {
                WriteCell(writer, ns, "ShapePlaceStyle", (int)shape.PlacementStyle.Value);
            }

            if (shape.PlacementFlip.HasValue) {
                WriteCell(writer, ns, "ShapePlaceFlip", (int)shape.PlacementFlip.Value);
            }

            if (shape.PlowCode.HasValue) {
                WriteCell(writer, ns, "ShapePlowCode", (int)shape.PlowCode.Value);
            }

            if (shape.AllowPlacementOnTop.HasValue) {
                WriteCell(writer, ns, "ShapePermeablePlace", shape.AllowPlacementOnTop.Value ? 1 : 0, "BOOL", null);
            }

            if (shape.AllowHorizontalConnectorRoutingThrough.HasValue) {
                WriteCell(writer, ns, "ShapePermeableX", shape.AllowHorizontalConnectorRoutingThrough.Value ? 1 : 0, "BOOL", null);
            }

            if (shape.AllowVerticalConnectorRoutingThrough.HasValue) {
                WriteCell(writer, ns, "ShapePermeableY", shape.AllowVerticalConnectorRoutingThrough.Value ? 1 : 0, "BOOL", null);
            }

            if (shape.CanSplitShapes.HasValue) {
                WriteCell(writer, ns, "ShapeSplit", shape.CanSplitShapes.Value ? 1 : 0);
            }

            if (shape.CanBeSplit.HasValue) {
                WriteCell(writer, ns, "ShapeSplittable", shape.CanBeSplit.Value ? 1 : 0);
            }
        }

        private static bool TryWriteModeledShapeCell(
            XmlWriter writer,
            string ns,
            VisioShape shape,
            string cellName,
            bool isGroup,
            double width,
            double height,
            double locPinX,
            double locPinY,
            IReadOnlyDictionary<string, string> persistedIds,
            IReadOnlyDictionary<string, int> layerIndexes) {
            switch (cellName) {
                case "PinX":
                    WriteCell(writer, ns, "PinX", shape.PinX);
                    return true;
                case "PinY":
                    WriteCell(writer, ns, "PinY", shape.PinY);
                    return true;
                case "Width":
                    WriteCell(writer, ns, "Width", width);
                    return true;
                case "Height":
                    WriteCell(writer, ns, "Height", height);
                    return true;
                case "LocPinX":
                    WriteCell(writer, ns, "LocPinX", locPinX);
                    return true;
                case "LocPinY":
                    WriteCell(writer, ns, "LocPinY", locPinY);
                    return true;
                case "Angle":
                    WriteCell(writer, ns, "Angle", shape.Angle);
                    return true;
                case "LineWeight":
                    WriteCell(writer, ns, "LineWeight", shape.LineWeight);
                    return true;
                case "LinePattern":
                    WriteCell(writer, ns, "LinePattern", shape.LinePattern);
                    return true;
                case "LineColor":
                    WriteCellValue(writer, ns, "LineColor", shape.LineColor.ToVisioHex());
                    return true;
                case "FillPattern":
                    WriteCell(writer, ns, "FillPattern", shape.FillPattern);
                    return true;
                case "FillForegnd":
                    WriteCellValue(writer, ns, "FillForegnd", shape.FillColor.ToVisioHex());
                    return true;
                case "ObjType":
                    if (!isGroup) {
                        WriteCell(writer, ns, "ObjType", 1);
                    }
                    return true;
                case "LeftMargin":
                    if (shape.TextStyle?.LeftMargin.HasValue == true) {
                        WriteCell(writer, ns, "LeftMargin", shape.TextStyle.LeftMargin.Value);
                    }
                    return true;
                case "RightMargin":
                    if (shape.TextStyle?.RightMargin.HasValue == true) {
                        WriteCell(writer, ns, "RightMargin", shape.TextStyle.RightMargin.Value);
                    }
                    return true;
                case "TopMargin":
                    if (shape.TextStyle?.TopMargin.HasValue == true) {
                        WriteCell(writer, ns, "TopMargin", shape.TextStyle.TopMargin.Value);
                    }
                    return true;
                case "BottomMargin":
                    if (shape.TextStyle?.BottomMargin.HasValue == true) {
                        WriteCell(writer, ns, "BottomMargin", shape.TextStyle.BottomMargin.Value);
                    }
                    return true;
                case "VerticalAlign":
                    if (shape.TextStyle?.VerticalAlignment.HasValue == true) {
                        WriteCell(writer, ns, "VerticalAlign", (int)shape.TextStyle.VerticalAlignment.Value);
                    }
                    return true;
                case "TextBkgnd":
                    if (shape.TextStyle?.BackgroundColor.HasValue == true) {
                        WriteCellValue(writer, ns, "TextBkgnd", shape.TextStyle.BackgroundColor.Value.ToVisioHex());
                    }
                    return true;
                case "TextBkgndTrans":
                    if (shape.TextStyle?.BackgroundTransparency.HasValue == true) {
                        WriteCell(writer, ns, "TextBkgndTrans", shape.TextStyle.BackgroundTransparency.Value);
                    }
                    return true;
                case "TxtPinX":
                    if (shape.TextStyle?.TextPinX.HasValue == true) {
                        WriteCell(writer, ns, "TxtPinX", shape.TextStyle.TextPinX.Value);
                    }
                    return true;
                case "TxtPinY":
                    if (shape.TextStyle?.TextPinY.HasValue == true) {
                        WriteCell(writer, ns, "TxtPinY", shape.TextStyle.TextPinY.Value);
                    }
                    return true;
                case "TxtWidth":
                    if (shape.TextStyle?.TextWidth.HasValue == true) {
                        WriteCell(writer, ns, "TxtWidth", shape.TextStyle.TextWidth.Value);
                    }
                    return true;
                case "TxtHeight":
                    if (shape.TextStyle?.TextHeight.HasValue == true) {
                        WriteCell(writer, ns, "TxtHeight", shape.TextStyle.TextHeight.Value);
                    }
                    return true;
                case "TxtLocPinX":
                    if (shape.TextStyle?.TextLocPinX.HasValue == true) {
                        WriteCell(writer, ns, "TxtLocPinX", shape.TextStyle.TextLocPinX.Value);
                    }
                    return true;
                case "TxtLocPinY":
                    if (shape.TextStyle?.TextLocPinY.HasValue == true) {
                        WriteCell(writer, ns, "TxtLocPinY", shape.TextStyle.TextLocPinY.Value);
                    }
                    return true;
                case "TxtAngle":
                    if (shape.TextStyle?.TextAngle.HasValue == true) {
                        WriteCell(writer, ns, "TxtAngle", shape.TextStyle.TextAngle.Value);
                    }
                    return true;
                case "LayerMember":
                    WriteLayerMemberCell(writer, ns, shape.LayerNames, layerIndexes);
                    return true;
                case "Relationships":
                    WriteRelationshipCell(writer, ns, shape, persistedIds);
                    return true;
                case "ShapePlaceStyle":
                    if (shape.PlacementStyle.HasValue) {
                        WriteCell(writer, ns, "ShapePlaceStyle", (int)shape.PlacementStyle.Value);
                    }
                    return true;
                case "ShapePlaceFlip":
                    if (shape.PlacementFlip.HasValue) {
                        WriteCell(writer, ns, "ShapePlaceFlip", (int)shape.PlacementFlip.Value);
                    }
                    return true;
                case "ShapePlowCode":
                    if (shape.PlowCode.HasValue) {
                        WriteCell(writer, ns, "ShapePlowCode", (int)shape.PlowCode.Value);
                    }
                    return true;
                case "ShapePermeablePlace":
                    if (shape.AllowPlacementOnTop.HasValue) {
                        WriteCell(writer, ns, "ShapePermeablePlace", shape.AllowPlacementOnTop.Value ? 1 : 0, "BOOL", null);
                    }
                    return true;
                case "ShapePermeableX":
                    if (shape.AllowHorizontalConnectorRoutingThrough.HasValue) {
                        WriteCell(writer, ns, "ShapePermeableX", shape.AllowHorizontalConnectorRoutingThrough.Value ? 1 : 0, "BOOL", null);
                    }
                    return true;
                case "ShapePermeableY":
                    if (shape.AllowVerticalConnectorRoutingThrough.HasValue) {
                        WriteCell(writer, ns, "ShapePermeableY", shape.AllowVerticalConnectorRoutingThrough.Value ? 1 : 0, "BOOL", null);
                    }
                    return true;
                case "ShapeSplit":
                    if (shape.CanSplitShapes.HasValue) {
                        WriteCell(writer, ns, "ShapeSplit", shape.CanSplitShapes.Value ? 1 : 0);
                    }
                    return true;
                case "ShapeSplittable":
                    if (shape.CanBeSplit.HasValue) {
                        WriteCell(writer, ns, "ShapeSplittable", shape.CanBeSplit.Value ? 1 : 0);
                    }
                    return true;
                default:
                    if (TryWriteProtectionCell(writer, ns, shape.Protection, cellName)) {
                        return true;
                    }

                    return false;
            }
        }

        private static bool HasAnyShapeTransformCellTokens(ISet<string> emittedTokens) {
            foreach (string cellName in ShapeTransformCellNames) {
                if (emittedTokens.Contains($"Cell:{cellName}")) {
                    return true;
                }
            }

            return false;
        }

        private static void MarkShapeTransformCellTokens(ISet<string> emittedTokens) {
            foreach (string cellName in ShapeTransformCellNames) {
                emittedTokens.Add($"Cell:{cellName}");
            }
        }

        private static void WriteRelationshipCell(XmlWriter writer, string ns, VisioShape shape, IReadOnlyDictionary<string, string> persistedIds) {
            string? formula = BuildRelationshipFormula(shape, persistedIds) ?? shape.RelationshipsFormula;
            if (string.IsNullOrWhiteSpace(formula) && string.IsNullOrWhiteSpace(shape.RelationshipsValue)) {
                return;
            }

            writer.WriteStartElement("Cell", ns);
            writer.WriteAttributeString("N", "Relationships");
            writer.WriteAttributeString("V", string.IsNullOrWhiteSpace(shape.RelationshipsValue) ? "0" : shape.RelationshipsValue);
            if (!string.IsNullOrWhiteSpace(formula)) {
                writer.WriteAttributeString("F", formula);
            }

            writer.WriteEndElement();
        }

        private static void WriteProtectionCells(XmlWriter writer, string ns, VisioProtection protection) {
            foreach (string cellName in VisioProtection.CellNames) {
                TryWriteProtectionCell(writer, ns, protection, cellName);
            }
        }

        private static bool TryWriteProtectionCell(XmlWriter writer, string ns, VisioProtection protection, string cellName) {
            if (!protection.TryGetCellValue(cellName, out bool? value)) {
                return false;
            }

            if (value.HasValue) {
                WriteCell(writer, ns, cellName, value.Value ? 1 : 0);
            }

            return true;
        }

        private static string? BuildRelationshipFormula(VisioShape shape, IReadOnlyDictionary<string, string> persistedIds) {
            List<string> dependencies = new();
            foreach (string memberId in shape.ContainerMemberIds.Where(id => !string.IsNullOrWhiteSpace(id)).Distinct(StringComparer.OrdinalIgnoreCase)) {
                dependencies.Add($"DEPENDSON(1,Sheet.{GetPersistedId(persistedIds, memberId)}!SheetRef())");
            }

            foreach (string ownerId in shape.ContainerOwnerIds.Where(id => !string.IsNullOrWhiteSpace(id)).Distinct(StringComparer.OrdinalIgnoreCase)) {
                dependencies.Add($"DEPENDSON(4,Sheet.{GetPersistedId(persistedIds, ownerId)}!SheetRef())");
            }

            return dependencies.Count == 0 ? null : $"SUM({string.Join(",", dependencies)})";
        }
    }
}
