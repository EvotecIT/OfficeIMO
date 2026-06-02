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

        private static bool ShouldPreservePageCell(string? cellName) {
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
                   !string.Equals(cellName, "PageShapeSplit", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "PlaceStyle", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "PlaceDepth", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "PlaceFlip", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "PlowCode", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "ResizePage", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "EnableGrid", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "BlockSizeX", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "BlockSizeY", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "AvenueSizeX", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "AvenueSizeY", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "RouteStyle", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "LineRouteExt", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "LineJumpStyle", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "LineJumpCode", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "PageLineJumpDirX", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "PageLineJumpDirY", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "LineToLineX", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "LineToLineY", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "LineToNodeX", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "LineToNodeY", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "ColorSchemeIndex", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "EffectSchemeIndex", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "ConnectorSchemeIndex", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "FontSchemeIndex", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "ThemeIndex", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "PageLeftMargin", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "PageRightMargin", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "PageTopMargin", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "PageBottomMargin", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(cellName, "PrintPageOrientation", StringComparison.OrdinalIgnoreCase);
        }

        private static bool ShouldPreservePageSection(XElement section) {
            return true;
        }

        private static bool ShouldPreservePageAttribute(XAttribute attribute) {
            string localName = attribute.Name.LocalName;
            string namespaceName = attribute.Name.NamespaceName;

            if (namespaceName == "http://www.w3.org/XML/1998/namespace") {
                return false;
            }

            return !string.Equals(localName, "ID", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "Name", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "NameU", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "Background", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "BackPage", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "ViewScale", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "ViewCenterX", StringComparison.OrdinalIgnoreCase) &&
                   !string.Equals(localName, "ViewCenterY", StringComparison.OrdinalIgnoreCase);
        }
    }
}
