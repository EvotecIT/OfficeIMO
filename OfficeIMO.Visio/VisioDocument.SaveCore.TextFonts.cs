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

        private void PrepareTextFontFaceNames(IEnumerable<VisioPage> pagesToSave) {
            Dictionary<string, int> faceIdsByName = new(StringComparer.OrdinalIgnoreCase);
            HashSet<int> usedIds = new();
            XNamespace ns = VisioNamespace;

            foreach (XElement faceName in PreservedFaceNamesElements.Where(element => string.Equals(element.Name.LocalName, "FaceName", StringComparison.OrdinalIgnoreCase))) {
                if (int.TryParse(faceName.Attribute("ID")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int id)) {
                    usedIds.Add(id);
                    string? name = faceName.Attribute("Name")?.Value;
                    if (!string.IsNullOrWhiteSpace(name) && !faceIdsByName.ContainsKey(name!)) {
                        faceIdsByName[name!] = id;
                    }
                }
            }

            foreach (VisioTextStyle textStyle in EnumerateTextStyles(pagesToSave)) {
                string? fontFamily = textStyle.FontFamily?.Trim();
                if (string.IsNullOrWhiteSpace(fontFamily)) {
                    textStyle.FontFaceId = null;
                    continue;
                }

                if (!faceIdsByName.TryGetValue(fontFamily!, out int faceId)) {
                    faceId = NextFaceNameId(usedIds);
                    usedIds.Add(faceId);
                    faceIdsByName[fontFamily!] = faceId;
                    PreservedFaceNamesElements.Add(new XElement(ns + "FaceName",
                        new XAttribute("ID", faceId),
                        new XAttribute("Name", fontFamily!),
                        new XAttribute("UnicodeRanges", "0-255"),
                        new XAttribute("CharSets", "0")));
                }

                textStyle.FontFamily = fontFamily;
                textStyle.FontFaceId = faceId;
            }
        }

        private static IEnumerable<VisioTextStyle> EnumerateTextStyles(IEnumerable<VisioPage> pages) {
            foreach (VisioPage page in pages) {
                foreach (VisioShape shape in page.Shapes) {
                    foreach (VisioTextStyle textStyle in EnumerateTextStyles(shape)) {
                        yield return textStyle;
                    }
                }

                foreach (VisioConnector connector in page.Connectors) {
                    if (connector.TextStyle != null) {
                        yield return connector.TextStyle;
                    }
                }
            }
        }

        private static IEnumerable<VisioTextStyle> EnumerateTextStyles(VisioShape shape) {
            if (shape.TextStyle != null) {
                yield return shape.TextStyle;
            }

            foreach (VisioShape child in shape.Children) {
                foreach (VisioTextStyle textStyle in EnumerateTextStyles(child)) {
                    yield return textStyle;
                }
            }
        }

        private static int NextFaceNameId(HashSet<int> usedIds) {
            int id = 0;
            while (usedIds.Contains(id)) {
                id++;
            }

            return id;
        }
    }
}
