using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioGeometryTests {
        private static string AssetsPath => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets"));

        [Fact]
        public void AllShapesHaveSizeAndGeometry() {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            foreach (string file in Directory.GetFiles(AssetsPath, "*.vsdx")) {
                VisioDocument doc = VisioDocument.Load(file);
                using ZipArchive pkg = ZipFile.OpenRead(file);
                for (int p = 0; p < doc.Pages.Count; p++) {
                    VisioPage page = doc.Pages[p];
                    ZipArchiveEntry? pageEntry = pkg.GetEntry($"visio/pages/page{p + 1}.xml");
                    Assert.NotNull(pageEntry);
                    XDocument pageDoc = XDocument.Load(pageEntry!.Open());
                    var shapeElements = pageDoc.Root?.Element(ns + "Shapes")?.Elements(ns + "Shape") ?? Enumerable.Empty<XElement>();
                    foreach (XElement shapeElement in shapeElements) {
                        string id = shapeElement.Attribute("ID")?.Value ?? string.Empty;
                        VisioShape shape = page.Shapes.First(s => s.Id == id);
                        bool referencesMaster = shapeElement.Attribute("Master") != null;
                        bool hasGeom = shapeElement.Element(ns + "Geom") != null;
                        Assert.True((shape.Width > 0 && shape.Height > 0) || referencesMaster || hasGeom,
                            $"{Path.GetFileName(file)} shape {id} has zero size");
                        if (!hasGeom && !referencesMaster) {
                            continue;
                        }
                    }
                }
            }
        }
    }
}

