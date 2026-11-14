using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using A19 = DocumentFormat.OpenXml.Office2019.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>Alternative-text and decorative metadata on worksheet drawing objects.</summary>
    internal static class ExcelDrawingAccessibility {
        private const string DecorativeExtensionUri = "{C183D7F6-B498-43B3-948B-1728B52AA6E4}";

        /// <summary>Clears the title and description and sets the Office decorative flag that Excel and checkers read.</summary>
        internal static void MarkDecorative(Xdr.NonVisualDrawingProperties properties) {
            properties.Description = string.Empty;
            properties.Title = string.Empty;
            A.NonVisualDrawingPropertiesExtensionList list = properties.GetFirstChild<A.NonVisualDrawingPropertiesExtensionList>()
                ?? properties.AppendChild(new A.NonVisualDrawingPropertiesExtensionList());
            A.NonVisualDrawingPropertiesExtension? extension = list.Elements<A.NonVisualDrawingPropertiesExtension>()
                .FirstOrDefault(item => string.Equals(item.Uri?.Value, DecorativeExtensionUri, StringComparison.OrdinalIgnoreCase));
            if (extension == null) {
                extension = new A.NonVisualDrawingPropertiesExtension { Uri = DecorativeExtensionUri };
                list.Append(extension);
            }
            extension.RemoveAllChildren<A19.Decorative>();
            extension.Append(new A19.Decorative { Val = true });
        }

        /// <summary>Removes the decorative flag, for example when alternative text is set.</summary>
        internal static void ClearDecorative(Xdr.NonVisualDrawingProperties properties) {
            A.NonVisualDrawingPropertiesExtensionList? list = properties.GetFirstChild<A.NonVisualDrawingPropertiesExtensionList>();
            if (list == null) return;
            foreach (A.NonVisualDrawingPropertiesExtension extension in list.Elements<A.NonVisualDrawingPropertiesExtension>()
                .Where(item => string.Equals(item.Uri?.Value, DecorativeExtensionUri, StringComparison.OrdinalIgnoreCase)).ToList()) {
                extension.Remove();
            }
            if (!list.HasChildren) list.Remove();
        }

        internal static bool IsDecorative(OpenXmlElement? properties) =>
            properties != null && properties.Descendants().Any(element => element.LocalName == "decorative" &&
                element.GetAttributes().Any(attribute => attribute.LocalName == "val" &&
                    (attribute.Value == "1" || string.Equals(attribute.Value, "true", StringComparison.OrdinalIgnoreCase))));
    }
}
