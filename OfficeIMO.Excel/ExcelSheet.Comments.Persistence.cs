using DocumentFormat.OpenXml.Packaging;
using System.Xml.Linq;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private XDocument LoadOrCreateVmlDocument(VmlDrawingPart part) {
            if (PendingCommentVmlDocuments.TryGetValue(part, out XDocument? pending)) return pending;
            try {
                using var stream = part.GetStream();
                if (stream.Length > 0) {
                    return LoadVmlXDocument(stream);
                }
            } catch (Exception ex) {
                System.Diagnostics.Debug.WriteLine($"Failed to load VML drawing part stream: {ex}");
            }

            var v = XNamespace.Get("urn:schemas-microsoft-com:vml");
            var o = XNamespace.Get("urn:schemas-microsoft-com:office:office");
            var x = XNamespace.Get("urn:schemas-microsoft-com:office:excel");

            var root = new XElement("xml",
                new XAttribute(XNamespace.Xmlns + "v", v),
                new XAttribute(XNamespace.Xmlns + "o", o),
                new XAttribute(XNamespace.Xmlns + "x", x),
                new XElement(o + "shapelayout",
                    new XAttribute(v + "ext", "edit"),
                    new XElement(o + "idmap", new XAttribute(v + "ext", "edit"), new XAttribute("data", "1"))
                ),
                new XElement(v + "shapetype",
                    new XAttribute("id", "_x0000_t202"),
                    new XAttribute("coordsize", "21600,21600"),
                    new XAttribute(o + "spt", "202"),
                    new XAttribute("path", "m,l,21600r21600,l21600,xe"),
                    new XElement(v + "stroke", new XAttribute("joinstyle", "miter")),
                    new XElement(v + "path", new XAttribute("gradientshapeok", "t"), new XAttribute(o + "connecttype", "rect"))
                )
            );

            return new XDocument(root);
        }

        private void SaveVmlDocument(VmlDrawingPart part, XDocument doc) {
            if (!_isBatchOperation) {
                using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
                doc.Save(stream);
                return;
            }
            PendingCommentVmlDocuments[part] = doc;
            MarkRequiresSavePreparation();
        }

        /// <summary>Flushes batched note drawings before a native writer or stream-based planner inspects them.</summary>
        internal void CommitCommentVmlChanges() {
            foreach (var pair in PendingCommentVmlDocuments) {
                if (!_worksheetPart.Parts.Any(item => ReferenceEquals(item.OpenXmlPart, pair.Key))) continue;
                using var stream = pair.Key.GetStream(FileMode.Create, FileAccess.Write);
                pair.Value.Save(stream);
            }
            PendingCommentVmlDocuments.Clear();
        }

    }
}
