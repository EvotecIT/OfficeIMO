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

        private static Dictionary<string, int> BuildLayerIndexMap(VisioPage page, out List<VisioLayer> layers) {
            List<VisioLayer> resolvedLayers = new();
            Dictionary<string, int> indexes = new(StringComparer.OrdinalIgnoreCase);

            void AddLayer(VisioLayer layer) {
                string key = GetLayerKey(layer);
                if (indexes.ContainsKey(key)) {
                    return;
                }

                indexes[key] = resolvedLayers.Count;
                if (!string.Equals(layer.Name, key, StringComparison.OrdinalIgnoreCase)) {
                    indexes[layer.Name] = resolvedLayers.Count;
                }

                resolvedLayers.Add(layer);
            }

            void AddSyntheticLayer(string layerName) {
                if (string.IsNullOrWhiteSpace(layerName) || indexes.ContainsKey(layerName)) {
                    return;
                }

                AddLayer(new VisioLayer(layerName));
            }

            void VisitShape(VisioShape shape) {
                foreach (string layerName in shape.LayerNames) {
                    AddSyntheticLayer(layerName);
                }

                foreach (VisioShape child in shape.Children) {
                    VisitShape(child);
                }
            }

            foreach (VisioLayer layer in page.Layers) {
                AddLayer(layer);
            }

            foreach (VisioShape shape in page.Shapes) {
                VisitShape(shape);
            }

            foreach (VisioConnector connector in page.Connectors) {
                foreach (string layerName in connector.LayerNames) {
                    AddSyntheticLayer(layerName);
                }
            }

            layers = resolvedLayers;
            return indexes;
        }

        private static string GetLayerKey(VisioLayer layer) {
            return string.IsNullOrWhiteSpace(layer.NameU) ? layer.Name : layer.NameU;
        }

        private static void WriteLayerSection(XmlWriter writer, string ns, IReadOnlyList<VisioLayer> layers) {
            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "Layer");
            int[] rowIndexes = VisioLayerRowIndexes.Create(layers);
            for (int i = 0; i < layers.Count; i++) {
                VisioLayer layer = layers[i];
                writer.WriteStartElement("Row", ns);
                writer.WriteAttributeString("IX", XmlConvert.ToString(rowIndexes[i]));
                WritePreservedAttributes(writer, layer.PreservedRowAttributes);
                WriteLayerCell(writer, ns, layer, "Name", layer.Name);
                WriteLayerCell(writer, ns, layer, "Color", layer.Color);
                WriteLayerCell(writer, ns, layer, "Status", layer.Status);
                WriteLayerCell(writer, ns, layer, "Visible", layer.Visible ? 1 : 0);
                WriteLayerCell(writer, ns, layer, "Print", layer.Print ? 1 : 0);
                WriteLayerCell(writer, ns, layer, "Active", layer.Active ? 1 : 0);
                WriteLayerCell(writer, ns, layer, "Lock", layer.Lock ? 1 : 0);
                WriteLayerCell(writer, ns, layer, "Snap", layer.Snap ? 1 : 0);
                WriteLayerCell(writer, ns, layer, "Glue", layer.Glue ? 1 : 0);
                WriteLayerCell(writer, ns, layer, "NameUniv", string.IsNullOrWhiteSpace(layer.NameU) ? layer.Name : layer.NameU);
                WriteLayerCell(writer, ns, layer, "ColorTrans", layer.ColorTransparency);
                WritePreservedElements(writer, layer.PreservedCells);
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }

        private static void WriteLayerCell(XmlWriter writer, string ns, VisioLayer layer, string cellName, int value) {
            WriteLayerCell(writer, ns, layer, cellName, value.ToString(CultureInfo.InvariantCulture));
        }

        private static void WriteLayerCell(XmlWriter writer, string ns, VisioLayer layer, string cellName, string value) {
            layer.PreservedKnownCells.TryGetValue(cellName, out XElement? template);
            writer.WriteStartElement("Cell", ns);
            writer.WriteAttributeString("N", cellName);
            writer.WriteAttributeString("V", value);
            if (template != null) {
                foreach (XAttribute attribute in template.Attributes()) {
                    if (attribute.IsNamespaceDeclaration ||
                        string.Equals(attribute.Name.LocalName, "N", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(attribute.Name.LocalName, "V", StringComparison.OrdinalIgnoreCase)) {
                        continue;
                    }

                    WriteAttribute(writer, attribute);
                }
            }

            writer.WriteEndElement();
        }

        private static void WriteLayerMemberCell(XmlWriter writer, string ns, IEnumerable<string> layerNames, IReadOnlyDictionary<string, int> layerIndexes, VisioLayerMembership? source = null) {
            List<int> memberIndexes = new();
            foreach (string layerName in layerNames) {
                if (layerIndexes.TryGetValue(layerName, out int index)) {
                    memberIndexes.Add(index);
                }
            }

            if (source?.IsCurrent(layerNames, memberIndexes) == true) {
                source.Cell.WriteTo(writer);
                return;
            }
            if (memberIndexes.Count == 0 && (source == null || !source.HasBoundNames)) {
                return;
            }

            string value = string.Join(";", memberIndexes.Distinct().OrderBy(index => index).Select(index => index.ToString(CultureInfo.InvariantCulture)));
            if (source == null) WriteCellValue(writer, ns, "LayerMember", value);
            else {
                source.ReplaceProducerState();
                var cell = new XElement(source.Cell);
                cell.SetAttributeValue("V", value);
                cell.Attribute("F")?.Remove();
                cell.Attribute("E")?.Remove();
                cell.Attribute("Err")?.Remove();
                cell.WriteTo(writer);
            }
        }

        private static void WriteMarginCells(XmlWriter writer, string ns, VisioPage page, bool useUnits) {
            if (useUnits || page.HasExplicitMargins && page.MarginUnit != VisioMeasurementUnit.Inches) {
                VisioMeasurementUnit unit = page.HasExplicitMargins ? page.MarginUnit : page.DefaultUnit;
                string unitCode = unit.ToVisioUnitCode();
                WritePageCell(writer, ns, "PageLeftMargin", page.LeftMargin, unitCode);
                WritePageCell(writer, ns, "PageRightMargin", page.RightMargin, unitCode);
                WritePageCell(writer, ns, "PageTopMargin", page.TopMargin, unitCode);
                WritePageCell(writer, ns, "PageBottomMargin", page.BottomMargin, unitCode);
                return;
            }

            WritePageCell(writer, ns, "PageLeftMargin", page.LeftMargin);
            WritePageCell(writer, ns, "PageRightMargin", page.RightMargin);
            WritePageCell(writer, ns, "PageTopMargin", page.TopMargin);
            WritePageCell(writer, ns, "PageBottomMargin", page.BottomMargin);
        }

        private static void WritePageLayoutRoutingCells(XmlWriter writer, string ns, VisioPage page) {
            if (page.ConnectorRouteStyle.HasValue) {
                WritePageCell(writer, ns, "RouteStyle", (int)page.ConnectorRouteStyle.Value);
            }

            if (page.ConnectorRouteAppearance.HasValue) {
                WritePageCell(writer, ns, "LineRouteExt", (int)page.ConnectorRouteAppearance.Value);
            }

            if (page.LineJumpStyle.HasValue) {
                WritePageCell(writer, ns, "LineJumpStyle", (int)page.LineJumpStyle.Value);
            }

            if (page.LineJumpCode.HasValue) {
                WritePageCell(writer, ns, "LineJumpCode", (int)page.LineJumpCode.Value);
            }

            if (page.HorizontalLineJumpDirection.HasValue) {
                WritePageCell(writer, ns, "PageLineJumpDirX", (int)page.HorizontalLineJumpDirection.Value);
            }

            if (page.VerticalLineJumpDirection.HasValue) {
                WritePageCell(writer, ns, "PageLineJumpDirY", (int)page.VerticalLineJumpDirection.Value);
            }
        }

        private static void WritePagePlacementCells(XmlWriter writer, string ns, VisioPage page) {
            if (page.PlacementStyle.HasValue) {
                WritePageCell(writer, ns, "PlaceStyle", (int)page.PlacementStyle.Value);
            }

            if (page.PlacementDepth.HasValue) {
                WritePageCell(writer, ns, "PlaceDepth", (int)page.PlacementDepth.Value);
            }

            if (page.PlacementFlip.HasValue) {
                WritePageCell(writer, ns, "PlaceFlip", (int)page.PlacementFlip.Value);
            }

            if (page.MoveShapesAwayOnDrop.HasValue) {
                WritePageCell(writer, ns, "PlowCode", page.MoveShapesAwayOnDrop.Value ? 1 : 0);
            }

            if (page.ResizePageToFitLayout.HasValue) {
                WritePageCell(writer, ns, "ResizePage", page.ResizePageToFitLayout.Value ? 1 : 0, "BOOL");
            }
        }

        private static void WritePageLayoutGridCells(XmlWriter writer, string ns, VisioPage page) {
            if (page.EnableLayoutGrid.HasValue) {
                WritePageCell(writer, ns, "EnableGrid", page.EnableLayoutGrid.Value ? 1 : 0, "BOOL");
            }

            if (!page.HasLayoutGridSizing) {
                return;
            }

            VisioMeasurementUnit unit = page.LayoutGridUnit;
            string unitCode = unit.ToVisioUnitCode();
            if (page.LayoutBlockSizeX.HasValue) {
                WritePageCell(writer, ns, "BlockSizeX", page.LayoutBlockSizeX.Value, unitCode);
            }

            if (page.LayoutBlockSizeY.HasValue) {
                WritePageCell(writer, ns, "BlockSizeY", page.LayoutBlockSizeY.Value, unitCode);
            }

            if (page.LayoutAvenueSizeX.HasValue) {
                WritePageCell(writer, ns, "AvenueSizeX", page.LayoutAvenueSizeX.Value, unitCode);
            }

            if (page.LayoutAvenueSizeY.HasValue) {
                WritePageCell(writer, ns, "AvenueSizeY", page.LayoutAvenueSizeY.Value, unitCode);
            }
        }

        private static void WritePageRoutingSpacingCells(XmlWriter writer, string ns, VisioPage page) {
            if (!page.HasConnectorSpacing) {
                return;
            }

            VisioMeasurementUnit unit = page.ConnectorSpacingUnit;
            string unitCode = unit.ToVisioUnitCode();
            if (page.LineToLineX.HasValue) {
                WritePageCell(writer, ns, "LineToLineX", page.LineToLineX.Value, unitCode);
            }

            if (page.LineToLineY.HasValue) {
                WritePageCell(writer, ns, "LineToLineY", page.LineToLineY.Value, unitCode);
            }

            if (page.LineToNodeX.HasValue) {
                WritePageCell(writer, ns, "LineToNodeX", page.LineToNodeX.Value, unitCode);
            }

            if (page.LineToNodeY.HasValue) {
                WritePageCell(writer, ns, "LineToNodeY", page.LineToNodeY.Value, unitCode);
            }
        }

        private static readonly string[] ConnectorModeledCellOrder = {
            "LineWeight",
            "LinePattern",
            "LineColor",
            "FillPattern",
            "FillForegnd",
            "LayerMember",
            "ShapeRouteStyle",
            "ConLineRouteExt",
            "ConLineJumpStyle",
            "ConLineJumpCode",
            "ConLineJumpDirX",
            "ConLineJumpDirY",
            "ConFixedCode",
            "BeginArrow",
            "EndArrow",
            "LeftMargin",
            "RightMargin",
            "TopMargin",
            "BottomMargin",
            "VerticalAlign",
            "TextBkgnd",
            "TextBkgndTrans",
            "TxtPinX",
            "TxtPinY",
            "TxtWidth",
            "TxtHeight",
            "TxtLocPinX",
            "TxtLocPinY",
            "TxtAngle",
            "LockWidth",
            "LockHeight",
            "LockAspect",
            "LockMoveX",
            "LockMoveY",
            "LockDelete",
            "LockTextEdit",
            "LockFormat",
            "LockGroup",
            "LockUngroup",
            "LockSelect",
            "LockRotate",
            "LockCrop",
            "LockVtxEdit",
            "LockBegin",
            "LockEnd",
            "LockCalcWH",
            "LockCustProp",
            "LockFromGroupFormat",
            "LockThemeColors",
            "LockThemeEffects"
        };

        private static readonly string[] ShapeModeledCellOrder = {
            "PinX",
            "PinY",
            "Width",
            "Height",
            "LocPinX",
            "LocPinY",
            "Angle",
            "LineWeight",
            "LinePattern",
            "LineColor",
            "FillPattern",
            "FillForegnd",
            "ObjType",
            "LeftMargin",
            "RightMargin",
            "TopMargin",
            "BottomMargin",
            "VerticalAlign",
            "TextBkgnd",
            "TextBkgndTrans",
            "TxtPinX",
            "TxtPinY",
            "TxtWidth",
            "TxtHeight",
            "TxtLocPinX",
            "TxtLocPinY",
            "TxtAngle",
            "LayerMember",
            "Relationships",
            "ShapePlaceStyle",
            "ShapePlaceFlip",
            "ShapePlowCode",
            "ShapePermeablePlace",
            "ShapePermeableX",
            "ShapePermeableY",
            "ShapeSplit",
            "ShapeSplittable",
            "LockWidth",
            "LockHeight",
            "LockAspect",
            "LockMoveX",
            "LockMoveY",
            "LockDelete",
            "LockTextEdit",
            "LockFormat",
            "LockGroup",
            "LockUngroup",
            "LockSelect",
            "LockRotate",
            "LockCrop",
            "LockVtxEdit",
            "LockBegin",
            "LockEnd",
            "LockCalcWH",
            "LockCustProp",
            "LockFromGroupFormat",
            "LockThemeColors",
            "LockThemeEffects"
        };

        // The same modeled cell names occur on thousands of loaded shapes and connectors.
        // Keep their order tokens shared while retaining the original token for unusual casing.
        private static readonly IReadOnlyDictionary<string, string> ModeledCellTokens = CreateModeledCellTokens();

        private static IReadOnlyDictionary<string, string> CreateModeledCellTokens() {
            Dictionary<string, string> tokens = new(StringComparer.Ordinal);
            foreach (string name in ShapeModeledCellOrder) {
                tokens[name] = $"Cell:{name}";
            }
            foreach (string name in ConnectorModeledCellOrder) {
                tokens[name] = $"Cell:{name}";
            }
            return tokens;
        }

        private static string GetModeledCellToken(string cellName) =>
            ModeledCellTokens.TryGetValue(cellName, out string? token) ? token : $"Cell:{cellName}";

        private static readonly string[] ShapeTransformCellNames = {
            "PinX",
            "PinY",
            "Width",
            "Height",
            "LocPinX",
            "LocPinY",
            "Angle"
        };

        private static Dictionary<string, string> BuildPersistedIdMap(VisioPage page, IReadOnlyDictionary<string, VisioMaster> effectiveMasters) =>
            BuildPersistedIdMap(page.Shapes, page.Connectors, effectiveMasters);
    }
}
