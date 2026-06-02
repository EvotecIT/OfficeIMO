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

        private static bool ShouldPreserveMasterAttribute(XAttribute attribute) {
            string localName = attribute.Name.LocalName;
            string namespaceName = attribute.Name.NamespaceName;

            if (namespaceName == "http://www.w3.org/XML/1998/namespace") {
                return false;
            }

            return !string.Equals(localName, "ID", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "Name", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "NameU", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "IsCustomNameU", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "IsCustomName", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "Prompt", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "IconSize", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "AlignName", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "MatchByName", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "IconUpdate", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "UniqueID", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "BaseID", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "PatternFlags", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "Hidden", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "MasterType", StringComparison.OrdinalIgnoreCase);
        }

        private static bool ShouldPreserveMasterPageSheetAttribute(XAttribute attribute) {
            string localName = attribute.Name.LocalName;
            string namespaceName = attribute.Name.NamespaceName;

            if (namespaceName == "http://www.w3.org/XML/1998/namespace") {
                return false;
            }

            return !string.Equals(localName, "LineStyle", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "FillStyle", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "TextStyle", StringComparison.OrdinalIgnoreCase);
        }

        private static bool ShouldPreserveMasterPageSheetCell(XElement cell) {
            string? cellName = cell.Attribute("N")?.Value;
            return !string.IsNullOrWhiteSpace(cellName) &&
                   !string.Equals(cellName, "PageWidth", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "PageHeight", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "ShdwOffsetX", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "ShdwOffsetY", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "PageScale", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "DrawingScale", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "DrawingSizeType", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "DrawingScaleType", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "InhibitSnap", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "PageLockReplace", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "PageLockDuplicate", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "UIVisibility", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "ShdwType", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "ShdwObliqueAngle", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "ShdwScaleFactor", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "DrawingResizeType", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "ShapeKeywords", StringComparison.OrdinalIgnoreCase);
        }

        private static bool ShouldPreserveMasterElement(XElement element) {
            string localName = element.Name.LocalName;
            return !string.Equals(localName, "PageSheet", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "Rel", StringComparison.OrdinalIgnoreCase);
        }

        private static bool ShouldPreserveMastersRootAttribute(XAttribute attribute) {
            if (attribute.IsNamespaceDeclaration) {
                return false;
            }

            string namespaceName = attribute.Name.NamespaceName;
            return namespaceName != "http://www.w3.org/XML/1998/namespace";
        }

        private static bool ShouldPreserveMastersRootElement(XElement element) {
            return !string.Equals(element.Name.LocalName, "Master", StringComparison.OrdinalIgnoreCase);
        }

        private static bool ShouldPreserveDocumentAttribute(XAttribute attribute) {
            if (attribute.IsNamespaceDeclaration) {
                return false;
            }

            return attribute.Name.NamespaceName != "http://www.w3.org/XML/1998/namespace";
        }

        private static bool ShouldPreserveDocumentElement(XElement element) {
            string localName = element.Name.LocalName;
            return !string.Equals(localName, "DocumentSettings", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "Colors", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "FaceNames", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "StyleSheets", StringComparison.OrdinalIgnoreCase);
        }

        private static bool ShouldPreserveDocumentSettingsAttribute(XAttribute attribute) {
            if (attribute.IsNamespaceDeclaration) {
                return false;
            }

            string localName = attribute.Name.LocalName;
            string namespaceName = attribute.Name.NamespaceName;

            if (namespaceName == "http://www.w3.org/XML/1998/namespace") {
                return false;
            }

            return !string.Equals(localName, "TopPage", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "DefaultTextStyle", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "DefaultLineStyle", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "DefaultFillStyle", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "DefaultGuideStyle", StringComparison.OrdinalIgnoreCase);
        }

        private static bool ShouldPreserveDocumentSettingsElement(XElement element) {
            string localName = element.Name.LocalName;
            return !string.Equals(localName, "GlueSettings", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "SnapSettings", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "SnapExtensions", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "SnapAngles", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "DynamicGridEnabled", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "ProtectStyles", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "ProtectShapes", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "ProtectMasters", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "ProtectBkgnds", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "RelayoutAndRerouteUponOpen", StringComparison.OrdinalIgnoreCase);
        }

        private static bool ShouldPreservePageContentsAttribute(XAttribute attribute) {
            if (attribute.IsNamespaceDeclaration) {
                return false;
            }

            string localName = attribute.Name.LocalName;
            string namespaceName = attribute.Name.NamespaceName;

            return !(namespaceName == "http://www.w3.org/XML/1998/namespace" &&
                     string.Equals(localName, "space", StringComparison.OrdinalIgnoreCase));
        }

        private static bool ShouldPreservePageContentsElement(XElement element) {
            string localName = element.Name.LocalName;
            return !string.Equals(localName, "Shapes", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "Connects", StringComparison.OrdinalIgnoreCase);
        }

        private static bool ShouldPreserveShapesContainerAttribute(XAttribute attribute) {
            return !attribute.IsNamespaceDeclaration &&
                   attribute.Name.NamespaceName != "http://www.w3.org/XML/1998/namespace";
        }

        private static bool ShouldPreserveShapesContainerElement(XElement element) {
            return !string.Equals(element.Name.LocalName, "Shape", StringComparison.OrdinalIgnoreCase);
        }

        private static bool ShouldPreserveConnectsAttribute(XAttribute attribute) {
            return !attribute.IsNamespaceDeclaration &&
                   attribute.Name.NamespaceName != "http://www.w3.org/XML/1998/namespace";
        }

        private static bool ShouldPreserveConnectsElement(XElement element) {
            return !string.Equals(element.Name.LocalName, "Connect", StringComparison.OrdinalIgnoreCase);
        }

        private static bool ShouldPreserveColorsAttribute(XAttribute attribute) {
            return !attribute.IsNamespaceDeclaration &&
                   attribute.Name.NamespaceName != "http://www.w3.org/XML/1998/namespace";
        }

        private static bool ShouldPreserveColorsElement(XElement element) {
            return true;
        }

        private static bool ShouldPreserveFaceNamesAttribute(XAttribute attribute) {
            return !attribute.IsNamespaceDeclaration &&
                   attribute.Name.NamespaceName != "http://www.w3.org/XML/1998/namespace";
        }

        private static bool ShouldPreserveFaceNamesElement(XElement element) {
            return true;
        }

        private static bool ShouldPreserveStyleSheetsAttribute(XAttribute attribute) {
            return !attribute.IsNamespaceDeclaration &&
                   attribute.Name.NamespaceName != "http://www.w3.org/XML/1998/namespace";
        }

        private static bool ShouldPreserveStyleSheetsElement(XElement element) {
            return !string.Equals(element.Name.LocalName, "StyleSheet", StringComparison.OrdinalIgnoreCase);
        }

        private static bool ShouldPreserveStyleSheetAttribute(XAttribute attribute, string styleSheetId) {
            if (attribute.IsNamespaceDeclaration) {
                return false;
            }

            string localName = attribute.Name.LocalName;
            string namespaceName = attribute.Name.NamespaceName;

            if (namespaceName == "http://www.w3.org/XML/1998/namespace") {
                return false;
            }

            if (string.Equals(localName, "ID", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(localName, "Name", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(localName, "NameU", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            if ((string.Equals(styleSheetId, "0", StringComparison.Ordinal) ||
                 string.Equals(styleSheetId, "1", StringComparison.Ordinal) ||
                 string.Equals(styleSheetId, "2", StringComparison.Ordinal)) &&
                (string.Equals(localName, "BasedOn", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(localName, "LineStyle", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(localName, "FillStyle", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(localName, "TextStyle", StringComparison.OrdinalIgnoreCase))) {
                return false;
            }

            return true;
        }

        private static bool ShouldPreserveStyleSheetElement(XElement element, string styleSheetId) {
            if (!string.Equals(element.Name.LocalName, "Cell", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            if (!(element.Attribute("N")?.Value is string cellName) ||
                string.IsNullOrWhiteSpace(cellName)) {
                return true;
            }

            return !GetGeneratedStyleSheetCellNames(styleSheetId).Contains(cellName);
        }

        private static bool IsGeneratedStyleSheet(string styleSheetId) {
            return string.Equals(styleSheetId, "0", StringComparison.Ordinal) ||
                   string.Equals(styleSheetId, "1", StringComparison.Ordinal) ||
                   string.Equals(styleSheetId, "2", StringComparison.Ordinal);
        }

        private static ISet<string> GetGeneratedStyleSheetCellNames(string styleSheetId) {
            return styleSheetId switch {
                "0" => new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
                    "EnableLineProps", "EnableFillProps", "EnableTextProps", "LineWeight", "LineColor", "LinePattern", "FillForegnd", "FillPattern"
                },
                "1" => new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
                    "LinePattern", "LineColor", "FillPattern", "FillForegnd"
                },
                "2" => new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
                    "EndArrow"
                },
                _ => new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            };
        }

        private static PreservedStyleSheetData GetOrCreatePreservedStyleSheet(VisioDocument document, string styleSheetId) {
            if (!document.PreservedGeneratedStyleSheets.TryGetValue(styleSheetId, out PreservedStyleSheetData? preserved)) {
                preserved = new PreservedStyleSheetData();
                document.PreservedGeneratedStyleSheets[styleSheetId] = preserved;
            }

            return preserved;
        }

        private static bool ShouldPreserveMasterContentAttribute(XAttribute attribute) {
            if (attribute.IsNamespaceDeclaration) {
                return false;
            }

            string localName = attribute.Name.LocalName;
            string namespaceName = attribute.Name.NamespaceName;

            return !(namespaceName == "http://www.w3.org/XML/1998/namespace" &&
                     string.Equals(localName, "space", StringComparison.OrdinalIgnoreCase));
        }

        private static bool ShouldPreserveMasterShapesAttribute(XAttribute attribute) {
            return !attribute.IsNamespaceDeclaration;
        }
    }
}
