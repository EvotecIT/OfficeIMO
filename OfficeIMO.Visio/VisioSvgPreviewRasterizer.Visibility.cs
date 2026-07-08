using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    internal static partial class VisioSvgPreviewRasterizer {
        private static bool IsElementDisplayNone(XElement element, SvgRenderContext context) {
            Dictionary<string, string> style = context.StyleSheet.CreateStyle(element);
            string? display = ReadStyleValue(element, style, "display");
            return string.Equals(display, "none", StringComparison.OrdinalIgnoreCase);
        }

        private static bool? ReadVisibilityOverride(XElement element, SvgRenderContext context) {
            Dictionary<string, string> style = context.StyleSheet.CreateStyle(element);
            return ReadVisibilityValue(element, style);
        }

        private static bool? ReadVisibilityValue(XElement element, Dictionary<string, string> style) {
            string? visibility = ReadStyleValue(element, style, "visibility");
            if (string.Equals(visibility, "visible", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            if (string.Equals(visibility, "hidden", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(visibility, "collapse", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            return null;
        }

        private static string? ReadStyleValue(XElement element, Dictionary<string, string> style, string name) =>
            style.TryGetValue(name, out string? value) ? value : element.Attribute(name)?.Value;
    }
}
