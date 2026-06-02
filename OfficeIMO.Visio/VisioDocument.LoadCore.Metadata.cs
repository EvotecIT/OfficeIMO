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

        private static void ParseHyperlinks(XElement hyperlinkSection, XNamespace ns, IList<VisioHyperlink> target) {
            target.Clear();
            foreach (XElement row in hyperlinkSection.Elements(ns + "Row")) {
                VisioHyperlink hyperlink = new() {
                    RowName = row.Attribute("N")?.Value
                };

                if (int.TryParse(row.Attribute("IX")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int rowIndex) &&
                    rowIndex >= 0) {
                    hyperlink.RowIndex = rowIndex;
                }

                foreach (XAttribute attribute in row.Attributes()) {
                    if (attribute.IsNamespaceDeclaration ||
                        string.Equals(attribute.Name.LocalName, "N", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(attribute.Name.LocalName, "IX", StringComparison.OrdinalIgnoreCase)) {
                        continue;
                    }

                    hyperlink.PreservedRowAttributes.Add(new XAttribute(attribute));
                }

                foreach (XElement cell in row.Elements(ns + "Cell")) {
                    string? cellName = cell.Attribute("N")?.Value;
                    string? value = cell.Attribute("V")?.Value;
                    if (IsHyperlinkCell(cellName)) {
                        hyperlink.PreservedKnownCells[cellName!] = new XElement(cell);
                    }

                    switch (cellName) {
                        case "Description":
                            hyperlink.Description = value;
                            break;
                        case "Address":
                            hyperlink.Address = value;
                            break;
                        case "SubAddress":
                            hyperlink.SubAddress = value;
                            break;
                        case "ExtraInfo":
                            hyperlink.ExtraInfo = value;
                            break;
                        case "Frame":
                            hyperlink.Frame = value;
                            break;
                        case "NewWindow":
                            hyperlink.NewWindow = TryParseTruthyCellValue(value);
                            break;
                        case "Default":
                            hyperlink.Default = TryParseTruthyCellValue(value);
                            break;
                        case "Invisible":
                            hyperlink.Invisible = TryParseTruthyCellValue(value);
                            break;
                        case "SortKey":
                            hyperlink.SortKey = value;
                            break;
                        default:
                            hyperlink.PreservedCells.Add(new XElement(cell));
                            break;
                    }
                }

                target.Add(hyperlink);
            }
        }

        private static void ParseUserCells(XElement userSection, XNamespace ns, IList<VisioUserCell> target) {
            target.Clear();
            foreach (XElement row in userSection.Elements(ns + "Row")) {
                string? rowName = row.Attribute("N")?.Value;
                int? rowIndex = null;
                if (int.TryParse(row.Attribute("IX")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedIndex) &&
                    parsedIndex >= 0) {
                    rowIndex = parsedIndex;
                }

                string resolvedName = string.IsNullOrWhiteSpace(rowName)
                    ? "User" + (rowIndex?.ToString(CultureInfo.InvariantCulture) ?? target.Count.ToString(CultureInfo.InvariantCulture))
                    : rowName!;
                VisioUserCell userCell = new(resolvedName) {
                    RowIndex = rowIndex
                };

                foreach (XAttribute attribute in row.Attributes()) {
                    if (attribute.IsNamespaceDeclaration ||
                        string.Equals(attribute.Name.LocalName, "N", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(attribute.Name.LocalName, "IX", StringComparison.OrdinalIgnoreCase)) {
                        continue;
                    }

                    userCell.PreservedRowAttributes.Add(new XAttribute(attribute));
                }

                foreach (XElement cell in row.Elements(ns + "Cell")) {
                    string? cellName = cell.Attribute("N")?.Value;
                    if (string.Equals(cellName, "Value", StringComparison.OrdinalIgnoreCase)) {
                        userCell.Value = cell.Attribute("V")?.Value;
                        userCell.Unit = cell.Attribute("U")?.Value;
                        userCell.Formula = cell.Attribute("F")?.Value;
                        CopyPreservedCellAttributes(cell, userCell.PreservedValueAttributes);
                    } else if (string.Equals(cellName, "Prompt", StringComparison.OrdinalIgnoreCase)) {
                        userCell.Prompt = cell.Attribute("V")?.Value;
                        userCell.PromptFormula = cell.Attribute("F")?.Value;
                        CopyPreservedCellAttributes(cell, userCell.PreservedPromptAttributes);
                    } else {
                        userCell.PreservedCells.Add(new XElement(cell));
                    }
                }

                target.Add(userCell);
            }
        }

        private static bool IsPackageBackedMasterUserCell(VisioUserCell cell) {
            return string.Equals(cell.Name, "OfficeIMO.PackageBackedMaster", StringComparison.OrdinalIgnoreCase) &&
                   (string.Equals(cell.Value, "1", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(cell.Value, "true", StringComparison.OrdinalIgnoreCase));
        }

        private static void CopyPreservedCellAttributes(XElement cell, IList<XAttribute> target) {
            foreach (XAttribute attribute in cell.Attributes()) {
                if (attribute.IsNamespaceDeclaration ||
                    string.Equals(attribute.Name.LocalName, "N", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(attribute.Name.LocalName, "V", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(attribute.Name.LocalName, "U", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(attribute.Name.LocalName, "F", StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                target.Add(new XAttribute(attribute));
            }
        }

        private static bool IsHyperlinkCell(string? cellName) {
            if (string.IsNullOrWhiteSpace(cellName)) {
                return false;
            }

            return VisioHyperlink.CellOrder.Any(known => string.Equals(known, cellName, StringComparison.OrdinalIgnoreCase));
        }
    }
}
