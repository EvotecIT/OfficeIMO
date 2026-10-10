using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    public partial class VisioDocument {
        private const int MaximumLoadedFontFamilyCharacters = 256;

        private static void PreserveDocumentMetadata(VisioDocument document, XElement root,
            IDictionary<int, string> faceNamesById, CancellationToken cancellationToken) {
            foreach (XAttribute attribute in root.Attributes().Where(ShouldPreserveDocumentAttribute)) {
                cancellationToken.ThrowIfCancellationRequested();
                document.PreservedDocumentAttributes.Add(new XAttribute(attribute));
            }
            foreach (XElement element in root.Elements().Where(ShouldPreserveDocumentElement)) {
                cancellationToken.ThrowIfCancellationRequested();
                document.PreservedDocumentElements.Add(new XElement(element));
            }

            XElement? documentSettings = root.Element(XName.Get("DocumentSettings", VisioNamespace));
            if (documentSettings != null) {
                foreach (XAttribute attribute in documentSettings.Attributes().Where(ShouldPreserveDocumentSettingsAttribute)) {
                    cancellationToken.ThrowIfCancellationRequested();
                    document.PreservedDocumentSettingsAttributes.Add(new XAttribute(attribute));
                }
                foreach (XElement element in documentSettings.Elements().Where(ShouldPreserveDocumentSettingsElement)) {
                    cancellationToken.ThrowIfCancellationRequested();
                    document.PreservedDocumentSettingsElements.Add(new XElement(element));
                }

                XElement? relayout = documentSettings.Element(XName.Get("RelayoutAndRerouteUponOpen", VisioNamespace));
                if (relayout != null && !string.Equals(relayout.Value, "0", StringComparison.OrdinalIgnoreCase)) {
                    document._requestRecalcOnOpen = true;
                }
            }

            XElement? colors = root.Element(XName.Get("Colors", VisioNamespace));
            if (colors != null) {
                foreach (XAttribute attribute in colors.Attributes().Where(ShouldPreserveColorsAttribute)) {
                    cancellationToken.ThrowIfCancellationRequested();
                    document.PreservedColorsAttributes.Add(new XAttribute(attribute));
                }

                foreach (XElement element in colors.Elements().Where(ShouldPreserveColorsElement)) {
                    cancellationToken.ThrowIfCancellationRequested();
                    document.PreservedColorsElements.Add(new XElement(element));
                }
            }

            XElement? faceNames = root.Element(XName.Get("FaceNames", VisioNamespace));
            if (faceNames != null) {
                foreach (XAttribute attribute in faceNames.Attributes().Where(ShouldPreserveFaceNamesAttribute)) {
                    cancellationToken.ThrowIfCancellationRequested();
                    document.PreservedFaceNamesAttributes.Add(new XAttribute(attribute));
                }

                foreach (XElement element in faceNames.Elements().Where(ShouldPreserveFaceNamesElement)) {
                    cancellationToken.ThrowIfCancellationRequested();
                    document.PreservedFaceNamesElements.Add(new XElement(element));
                    if (string.Equals(element.Name.LocalName, "FaceName", StringComparison.OrdinalIgnoreCase) &&
                        int.TryParse(element.Attribute("ID")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int faceId)) {
                        string? name = element.Attribute("Name")?.Value;
                        if (!string.IsNullOrWhiteSpace(name) && !faceNamesById.ContainsKey(faceId)) {
                            string normalizedName = name!.Trim();
                            faceNamesById[faceId] = normalizedName.Length <= MaximumLoadedFontFamilyCharacters
                                ? normalizedName
                                : normalizedName.Substring(0, MaximumLoadedFontFamilyCharacters);
                        }
                    }
                }
            }

            XElement? styleSheets = root.Element(XName.Get("StyleSheets", VisioNamespace));
            if (styleSheets != null) {
                foreach (XAttribute attribute in styleSheets.Attributes().Where(ShouldPreserveStyleSheetsAttribute)) {
                    cancellationToken.ThrowIfCancellationRequested();
                    document.PreservedStyleSheetsAttributes.Add(new XAttribute(attribute));
                }

                foreach (XElement element in styleSheets.Elements().Where(ShouldPreserveStyleSheetsElement)) {
                    cancellationToken.ThrowIfCancellationRequested();
                    document.PreservedStyleSheetsElements.Add(new XElement(element));
                }

                foreach (XElement styleSheet in styleSheets.Elements(XName.Get("StyleSheet", VisioNamespace))) {
                    cancellationToken.ThrowIfCancellationRequested();
                    string id = NormalizeStyleSheetId(styleSheet.Attribute("ID")?.Value ?? string.Empty);
                    if (!IsGeneratedStyleSheet(id)) {
                        document.PreservedAdditionalStyleSheets.Add(new XElement(styleSheet));
                        continue;
                    }

                    PreservedStyleSheetData preserved = GetOrCreatePreservedStyleSheet(document, id);
                    foreach (XAttribute attribute in styleSheet.Attributes().Where(attribute => ShouldPreserveStyleSheetAttribute(attribute, id))) {
                        cancellationToken.ThrowIfCancellationRequested();
                        preserved.Attributes.Add(new XAttribute(attribute));
                    }

                    foreach (XElement element in styleSheet.Elements().Where(element => ShouldPreserveStyleSheetElement(element, id))) {
                        cancellationToken.ThrowIfCancellationRequested();
                        preserved.ChildElements.Add(new XElement(element));
                    }
                }
            }
        }
    }
}
