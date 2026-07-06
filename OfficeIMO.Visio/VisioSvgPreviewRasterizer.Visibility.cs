using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    internal static partial class VisioSvgPreviewRasterizer {
        private static bool IsElementHidden(XElement element, SvgRenderContext context) {
            Dictionary<string, string> style = context.StyleSheet.CreateStyle(element);
            string? display = ReadStyleValue(element, style, "display");
            if (string.Equals(display, "none", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            string? visibility = ReadStyleValue(element, style, "visibility");
            return string.Equals(visibility, "hidden", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(visibility, "collapse", StringComparison.OrdinalIgnoreCase);
        }

        private static string? ReadStyleValue(XElement element, Dictionary<string, string> style, string name) =>
            element.Attribute(name)?.Value ?? (style.TryGetValue(name, out string? value) ? value : null);
    }
}
