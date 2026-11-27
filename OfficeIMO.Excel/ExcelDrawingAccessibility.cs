using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel {
    /// <summary>Alternative-text and decorative metadata on worksheet drawing objects.</summary>
    internal static class ExcelDrawingAccessibility {
        private const string DecorativeExtensionUri = "{C183D7F6-B498-43B3-948B-1728B52AA6E4}";
        private const string DecorativeNamespace = "http://schemas.microsoft.com/office/drawing/2017/decorative";

        /// <summary>Clears the title and description and sets the Office decorative flag that Excel and checkers read.</summary>
        /// <param name="properties">The drawing's <c>xdr:cNvPr</c>.</param>
        internal static void MarkDecorative(OfficeXmlNode properties) {
            properties.SetAttribute("descr", string.Empty);
            properties.SetAttribute("title", string.Empty);
            OfficeXmlNode list = properties.Child(A, "extLst");
            if (list.IsNone) list = properties.Append(properties.New(A, "extLst"));
            OfficeXmlNode extension = default;
            foreach (OfficeXmlNode item in list.Children(A, "ext")) {
                if (string.Equals(item.Attribute("uri"), DecorativeExtensionUri, StringComparison.OrdinalIgnoreCase)) {
                    extension = item;
                    break;
                }
            }
            if (extension.IsNone) extension = list.Append(list.New(A, "ext").Attr("uri", DecorativeExtensionUri));
            for (OfficeXmlNode child = extension.FirstChild; !child.IsNone;) {
                OfficeXmlNode next = child.NextSibling;
                if (child.Is(DecorativeNamespace, "decorative")) child.Remove();
                child = next;
            }
            extension.Append(extension.CreateElement(DecorativeNamespace, "decorative", "adec").SetVal(true));
        }

        /// <summary>Removes the decorative flag, for example when alternative text is set.</summary>
        /// <param name="properties">The drawing's <c>xdr:cNvPr</c>.</param>
        internal static void ClearDecorative(OfficeXmlNode properties) {
            OfficeXmlNode list = properties.Child(A, "extLst");
            if (list.IsNone) return;
            foreach (OfficeXmlNode extension in list.Children(A, "ext").ToList()) {
                if (string.Equals(extension.Attribute("uri"), DecorativeExtensionUri, StringComparison.OrdinalIgnoreCase)) extension.Remove();
            }
            if (list.FirstChild.IsNone) list.Remove();
        }

        /// <summary>Whether any element below the drawing properties is a <c>decorative</c> flag set to true.</summary>
        internal static bool IsDecorative(OfficeXmlNode properties) {
            if (properties.IsNone) return false;
            for (OfficeXmlNode current = properties.FirstChild; !current.IsNone; current = current.NextInSubtree(properties)) {
                if (current.LocalName != "decorative") continue;
                string? value = current.Attribute("val");
                if (value == "1" || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
        }
    }
}
