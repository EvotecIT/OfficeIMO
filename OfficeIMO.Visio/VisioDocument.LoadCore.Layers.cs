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

        private static void ParseLayerSection(VisioPage page, XElement layerSection, XNamespace ns) {
            foreach (XElement row in layerSection.Elements(ns + "Row")) {
                int? sourceIndex = null;
                if (int.TryParse(row.Attribute("IX")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedIndex) &&
                    parsedIndex >= 0) {
                    sourceIndex = parsedIndex;
                }

                string? name = null;
                string? nameU = null;
                int color = 255;
                int status = 0;
                bool visible = true;
                bool print = true;
                bool active = false;
                bool locked = false;
                bool snap = true;
                bool glue = true;
                int colorTransparency = 0;
                Dictionary<string, XElement> preservedKnownCells = new(StringComparer.OrdinalIgnoreCase);
                List<XElement> preservedCells = new();

                foreach (XElement cell in row.Elements(ns + "Cell")) {
                    string? cellName = cell.Attribute("N")?.Value;
                    string? value = cell.Attribute("V")?.Value;
                    if (IsKnownLayerCell(cellName)) {
                        preservedKnownCells[cellName!] = new XElement(cell);
                    }

                    switch (cellName) {
                        case "Name":
                            name = value;
                            break;
                        case "NameUniv":
                            nameU = value;
                            break;
                        case "Color":
                            if (TryParseCellIntValue(value, out int parsedColor)) {
                                color = parsedColor;
                            }
                            break;
                        case "Status":
                            if (TryParseCellIntValue(value, out int parsedStatus)) {
                                status = parsedStatus;
                            }
                            break;
                        case "Visible":
                            visible = ParseBoolCell(value, visible);
                            break;
                        case "Print":
                            print = ParseBoolCell(value, print);
                            break;
                        case "Active":
                            active = ParseBoolCell(value, active);
                            break;
                        case "Lock":
                            locked = ParseBoolCell(value, locked);
                            break;
                        case "Snap":
                            snap = ParseBoolCell(value, snap);
                            break;
                        case "Glue":
                            glue = ParseBoolCell(value, glue);
                            break;
                        case "ColorTrans":
                            if (TryParseCellIntValue(value, out int parsedTransparency)) {
                                colorTransparency = parsedTransparency;
                            }
                            break;
                        default:
                            preservedCells.Add(new XElement(cell));
                            break;
                    }
                }

                string resolvedName = string.IsNullOrWhiteSpace(name) ? $"Layer {sourceIndex ?? page.Layers.Count}" : name!;
                VisioLayer layer = new(resolvedName, nameU) {
                    Color = color,
                    Status = status,
                    Visible = visible,
                    Print = print,
                    Active = active,
                    Lock = locked,
                    Snap = snap,
                    Glue = glue,
                    ColorTransparency = colorTransparency,
                    SourceIndex = sourceIndex
                };

                foreach (XAttribute attribute in row.Attributes().Where(attribute =>
                             !string.Equals(attribute.Name.LocalName, "IX", StringComparison.OrdinalIgnoreCase))) {
                    layer.PreservedRowAttributes.Add(new XAttribute(attribute));
                }

                foreach (KeyValuePair<string, XElement> knownCell in preservedKnownCells) {
                    layer.PreservedKnownCells[knownCell.Key] = knownCell.Value;
                }

                foreach (XElement preservedCell in preservedCells) {
                    layer.PreservedCells.Add(preservedCell);
                }

                page.Layers.Add(layer);
            }
        }

        private static void ApplyLayerNamesFromIndexes(VisioPage page, VisioShape shape) {
            foreach (int layerIndex in shape.LayerIndexes) {
                VisioLayer? layer = FindLayerBySourceIndex(page, layerIndex);
                if (layer != null) {
                    shape.LayerNames.Add(string.IsNullOrWhiteSpace(layer.NameU) ? layer.Name : layer.NameU);
                }
            }

            foreach (VisioShape child in shape.Children) {
                ApplyLayerNamesFromIndexes(page, child);
            }
        }

        private static void ApplyLayerNamesFromIndexes(VisioPage page, VisioConnector connector) {
            foreach (int layerIndex in connector.LayerIndexes) {
                VisioLayer? layer = FindLayerBySourceIndex(page, layerIndex);
                if (layer != null) {
                    connector.LayerNames.Add(string.IsNullOrWhiteSpace(layer.NameU) ? layer.Name : layer.NameU);
                }
            }
        }

        private static VisioLayer? FindLayerBySourceIndex(VisioPage page, int layerIndex) {
            if (layerIndex >= 0 && layerIndex < page.Layers.Count) {
                return page.Layers[layerIndex];
            }

            foreach (VisioLayer layer in page.Layers) {
                if (layer.SourceIndex == layerIndex) {
                    return layer;
                }
            }

            return null;
        }

        private static bool ParseBoolCell(string? value, bool fallback) {
            return ParseNullableBoolCell(value) ?? fallback;
        }

        private static bool IsKnownLayerCell(string? cellName) {
            switch (cellName) {
                case "Name":
                case "Color":
                case "Status":
                case "Visible":
                case "Print":
                case "Active":
                case "Lock":
                case "Snap":
                case "Glue":
                case "NameUniv":
                case "ColorTrans":
                    return true;
                default:
                    return false;
            }
        }

        private static void ParseLayerIndexes(string? value, IList<int> target) {
            target.Clear();
            if (string.IsNullOrWhiteSpace(value)) {
                return;
            }

            string[] parts = value!.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string part in parts) {
                if (int.TryParse(part.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int index) &&
                    index >= 0) {
                    target.Add(index);
                }
            }
        }
    }
}
