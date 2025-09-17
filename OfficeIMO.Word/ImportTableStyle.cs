using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides helper methods for copying table styles between Word documents.
    /// </summary>
    internal class ImportTableStyle {
        private static XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        /// <summary>
        /// Copies table styles from the <paramref name="sourcefilepath"/> document to the
        /// <paramref name="destinationfilepath"/> document.
        /// </summary>
        /// <param name="sourcefilepath">Path to the source document containing styles to copy.</param>
        /// <param name="destinationfilepath">Path to the destination document that receives the styles.</param>
        internal static void ImportTableStyles(string sourcefilepath, string destinationfilepath) {
            using WordprocessingDocument sourceDocument = WordprocessingDocument.Open(sourcefilepath, true);
            MainDocumentPart sourceMainPart = sourceDocument.MainDocumentPart ?? throw new InvalidOperationException("Source document is missing a main document part.");
            StyleDefinitionsPart sourceStylePart = sourceMainPart.StyleDefinitionsPart ?? throw new InvalidOperationException("Source document does not contain styles.");

            XDocument sourceStyleDoc;
            using (TextReader reader = new StreamReader(sourceStylePart.GetStream())) {
                sourceStyleDoc = XDocument.Load(reader);
            }

            XName stylesName = XName.Get("styles", w.NamespaceName);
            XName styleName = XName.Get("style", w.NamespaceName);
            XName styleIdName = XName.Get("styleId");
            XName typeName = XName.Get("type", w.NamespaceName);

            List<XElement> sourceTableStyles = sourceStyleDoc
                .Element(stylesName)?
                .Elements(styleName)
                .Where(style => string.Equals((string?)style.Attribute(typeName), "table", StringComparison.Ordinal))
                .Select(style => new XElement(style))
                .ToList() ?? new List<XElement>();

            using WordprocessingDocument destinationDocument = WordprocessingDocument.Open(destinationfilepath, true);
            MainDocumentPart destinationMainPart = destinationDocument.MainDocumentPart ?? throw new InvalidOperationException("Destination document is missing a main document part.");
            StyleDefinitionsPart destinationStylePart = destinationMainPart.StyleDefinitionsPart ?? throw new InvalidOperationException("Destination document does not contain styles.");

            XDocument destinationStyleDoc;
            using (TextReader reader = new StreamReader(destinationStylePart.GetStream())) {
                destinationStyleDoc = XDocument.Load(reader);
            }

            XElement stylesRoot = destinationStyleDoc.Element(stylesName) ?? new XElement(stylesName);
            if (stylesRoot.Parent == null) {
                destinationStyleDoc.Add(stylesRoot);
            }

            foreach (XElement styleElement in sourceTableStyles) {
                string? styleId = styleElement.Attribute(styleIdName)?.Value;
                if (string.IsNullOrEmpty(styleId)) {
                    continue;
                }

                bool exists = stylesRoot
                    .Elements(styleName)
                    .Any(existing => string.Equals((string?)existing.Attribute(styleIdName), styleId, StringComparison.Ordinal));

                if (!exists) {
                    stylesRoot.Add(new XElement(styleElement));
                }
            }

            using TextWriter writer = new StreamWriter(destinationStylePart.GetStream(FileMode.Create));
            destinationStyleDoc.Save(writer, SaveOptions.None);
        }
    }
}
