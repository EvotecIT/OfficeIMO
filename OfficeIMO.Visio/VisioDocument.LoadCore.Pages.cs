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

        private static void ResolvePageBackgrounds(VisioDocument document) {
            Dictionary<int, VisioPage> pagesById = document.Pages
                .GroupBy(page => page.Id)
                .Where(group => group.Count() == 1)
                .ToDictionary(group => group.Key, group => group.First());

            foreach (VisioPage page in document.Pages) {
                page.ResolveBackgroundPage(pagesById);
            }
        }

        private static VisioConnectorLabelPlacement EnsureConnectorLabelPlacement(VisioConnector connector) {
            if (connector.LabelPlacement == null) {
                connector.LabelPlacement = new VisioConnectorLabelPlacement();
            }

            return connector.LabelPlacement;
        }

        private static VisioTextStyle EnsureConnectorTextStyle(VisioConnector connector) {
            if (connector.TextStyle == null) {
                connector.TextStyle = new VisioTextStyle();
            }

            return connector.TextStyle;
        }

        private const int MaxShapeNestingDepth = 100;
        private static readonly double DefaultLineWeight = VisioShape.DefaultLineWeight;

        private static double ParseDouble(string? value) {
            string? normalized = NormalizeCellLiteral(value);
            return double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out double result) ? result : 0;
        }

        private static double? ParseNonNegativeDouble(string? value) {
            double parsed = ParseDouble(value);
            return double.IsNaN(parsed) || double.IsInfinity(parsed) || parsed < 0 ? null : parsed;
        }

        private static bool IsValidPlacementFlip(int value) {
            const int allKnownFlags = (int)(VisioPlacementFlip.Horizontal |
                                            VisioPlacementFlip.Vertical |
                                            VisioPlacementFlip.Rotate90 |
                                            VisioPlacementFlip.None);
            return value >= 0 && (value & ~allKnownFlags) == 0;
        }

        private static OfficeIMO.Drawing.OfficeColor ParseColor(string? value, OfficeIMO.Drawing.OfficeColor fallback) {
            string? normalized = NormalizeCellLiteral(value);
            return string.IsNullOrWhiteSpace(normalized) ? fallback : VisioHelpers.FromVisioColor(normalized!);
        }

    }
}
