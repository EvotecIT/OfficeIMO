using DocumentFormat.OpenXml.Packaging;
using System.Globalization;
using System.Xml.Linq;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static XDocument LoadOrCreateHeaderFooterVmlDocument(VmlDrawingPart vmlPart) {
            try {
                using Stream stream = vmlPart.GetStream(FileMode.Open, FileAccess.Read);
                if (!stream.CanSeek || stream.Length > 0) {
                    XDocument existing = LoadVmlXDocument(stream);
                    if (existing.Root != null) {
                        EnsureHeaderFooterVmlNamespaces(existing.Root);
                        EnsureHeaderFooterVmlShapeType(existing.Root);
                        return existing;
                    }
                }
            } catch {
                // Recreate malformed writer-owned VML rather than preserving broken markup.
            }

            var document = new XDocument(new XElement("xml"));
            EnsureHeaderFooterVmlNamespaces(document.Root!);
            EnsureHeaderFooterVmlShapeType(document.Root!);
            return document;
        }

        private static void UpsertHeaderFooterVmlShape(XDocument document, string shapeId, string imageRelationshipId, double widthPoints, double heightPoints) {
            XElement root = document.Root ?? new XElement("xml");
            if (document.Root == null) {
                document.Add(root);
            }

            EnsureHeaderFooterVmlNamespaces(root);
            EnsureHeaderFooterVmlShapeType(root);

            document.Descendants()
                .Where(element => string.Equals(element.Name.LocalName, "shape", StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(element.Attribute("id")?.Value, shapeId, StringComparison.OrdinalIgnoreCase))
                .Remove();

            XNamespace v = HeaderFooterVmlNamespace;
            XNamespace o = HeaderFooterOfficeNamespace;
            XNamespace r = HeaderFooterRelationshipNamespace;
            root.Add(new XElement(v + "shape",
                new XAttribute("id", shapeId),
                new XAttribute(o + "spid", "_x0000_s" + HeaderFooterShapeSpid(shapeId).ToString(CultureInfo.InvariantCulture)),
                new XAttribute("type", "#_x0000_t75"),
                new XAttribute("style", "position:absolute;margin-left:0;margin-top:0;width:" +
                    widthPoints.ToString(CultureInfo.InvariantCulture) + "pt;height:" +
                    heightPoints.ToString(CultureInfo.InvariantCulture) + "pt;z-index:1"),
                new XElement(v + "imagedata",
                    new XAttribute(r + "id", imageRelationshipId),
                    new XAttribute(o + "relid", imageRelationshipId),
                    new XAttribute(o + "title", string.Empty))));
        }

        private static void EnsureHeaderFooterVmlNamespaces(XElement root) {
            AddNamespaceDeclaration(root, "v", HeaderFooterVmlNamespace);
            AddNamespaceDeclaration(root, "o", HeaderFooterOfficeNamespace);
            AddNamespaceDeclaration(root, "x", HeaderFooterExcelNamespace);
            AddNamespaceDeclaration(root, "r", HeaderFooterRelationshipNamespace);
        }

        private static void EnsureHeaderFooterVmlShapeType(XElement root) {
            XNamespace v = HeaderFooterVmlNamespace;
            XNamespace o = HeaderFooterOfficeNamespace;
            if (root.Elements(v + "shapetype").Any(element => string.Equals(element.Attribute("id")?.Value, "_x0000_t75", StringComparison.OrdinalIgnoreCase))) {
                return;
            }

            root.AddFirst(
                new XElement(v + "shapetype",
                    new XAttribute("id", "_x0000_t75"),
                    new XAttribute("coordsize", "21600,21600"),
                    new XAttribute(o + "spt", "75"),
                    new XAttribute(o + "preferrelative", "t"),
                    new XAttribute("path", "m@4@5l@4@11@9@11@9@5xe"),
                    new XAttribute("filled", "f"),
                    new XAttribute("stroked", "f"),
                    new XElement(v + "stroke", new XAttribute("joinstyle", "miter")),
                    new XElement(v + "formulas",
                        new XElement(v + "f", new XAttribute("eqn", "if lineDrawn pixelLineWidth 0")),
                        new XElement(v + "f", new XAttribute("eqn", "sum @0 1 0")),
                        new XElement(v + "f", new XAttribute("eqn", "sum 0 0 @1")),
                        new XElement(v + "f", new XAttribute("eqn", "prod @2 1 2")),
                        new XElement(v + "f", new XAttribute("eqn", "prod @3 21600 pixelWidth")),
                        new XElement(v + "f", new XAttribute("eqn", "prod @3 21600 pixelHeight")),
                        new XElement(v + "f", new XAttribute("eqn", "sum @0 0 1")),
                        new XElement(v + "f", new XAttribute("eqn", "prod @6 1 2")),
                        new XElement(v + "f", new XAttribute("eqn", "prod @7 21600 pixelWidth")),
                        new XElement(v + "f", new XAttribute("eqn", "sum @8 21600 0")),
                        new XElement(v + "f", new XAttribute("eqn", "prod @7 21600 pixelHeight")),
                        new XElement(v + "f", new XAttribute("eqn", "sum @10 21600 0"))),
                    new XElement(v + "path",
                        new XAttribute(o + "extrusionok", "f"),
                        new XAttribute("gradientshapeok", "t"),
                        new XAttribute(o + "connecttype", "rect")),
                    new XElement(o + "lock",
                        new XAttribute(HeaderFooterVmlNamespace + "ext", "edit"),
                        new XAttribute("aspectratio", "t"))));
        }

        private static void AddNamespaceDeclaration(XElement root, string prefix, XNamespace ns) {
            XName attributeName = XNamespace.Xmlns + prefix;
            if (root.Attribute(attributeName) == null) {
                root.SetAttributeValue(attributeName, ns.NamespaceName);
            }
        }

        private static int HeaderFooterShapeSpid(string shapeId) =>
            shapeId.ToUpperInvariant() switch {
                "LH" => 1025,
                "CH" => 1026,
                "RH" => 1027,
                "LF" => 1028,
                "CF" => 1029,
                "RF" => 1030,
                _ => 1031
            };

        private static readonly XNamespace HeaderFooterVmlNamespace = "urn:schemas-microsoft-com:vml";
        private static readonly XNamespace HeaderFooterOfficeNamespace = "urn:schemas-microsoft-com:office:office";
        private static readonly XNamespace HeaderFooterExcelNamespace = "urn:schemas-microsoft-com:office:excel";
        private static readonly XNamespace HeaderFooterRelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    }
}
