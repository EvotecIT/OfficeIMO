using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        private const string SmartArtExtensionUri = "{6F0E8D7C-1B2A-4C3D-9E8F-0A1B2C3D4E5F}";

        [Fact]
        public void Test_SmartArt_IgnoresRelationshipIdsInsideDrawingExtensions() {
            string filePath = Path.Combine(_directoryWithFiles, "SmartArtExtensionRelIds.docx");
            using (var document = WordDocument.Create(filePath)) {
                document.AddSmartArt(WordSmartArtType.BasicProcess).SetNodeText(0, "Kept");
                document.Save();
            }

            // A dgm:relIds in an extension of the drawing properties names parts that do not exist. Reading it
            // instead of the graphic data's relIds would lose the diagram's text and type.
            RewritePart(filePath, "word/document.xml", xml => Regex.Replace(xml, "(<wp:docPr [^>]*?)\\s*/>",
                "$1><a:extLst xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:ext uri=\"" + SmartArtExtensionUri + "\">" +
                "<dgm:relIds xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\" " +
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                "r:dm=\"rIdMissing1\" r:lo=\"rIdMissing2\" r:qs=\"rIdMissing3\" r:cs=\"rIdMissing4\" />" +
                "</a:ext></a:extLst></wp:docPr>", RegexOptions.None, TimeSpan.FromSeconds(5)));

            using (var document = WordDocument.Load(filePath)) {
                WordSmartArt smartArt = Assert.Single(document.SmartArts);
                Assert.Equal("Kept", smartArt.GetNodeText(0));
                smartArt.SetNodeText(0, "Edited");
                document.Save();
            }

            using (var document = WordDocument.Load(filePath)) {
                Assert.Equal("Edited", Assert.Single(document.SmartArts).GetNodeText(0));
            }
        }

        [Fact]
        public void Test_SmartArt_PersistedLayoutIgnoresShapesInsideExtensions() {
            string filePath = Path.Combine(_directoryWithFiles, "SmartArtExtensionShapes.docx");
            int expected;
            using (var document = WordDocument.Create(filePath)) {
                WordSmartArt smartArt = document.AddSmartArt(WordSmartArtType.CustomSmartArt1);
                expected = smartArt.GetPersistedLayoutShapes().Count;
                document.Save();
            }
            Assert.True(expected > 0);

            // A copy of a laid-out shape inside the shape tree's extension list is not part of the layout.
            RewritePart(filePath, part => part.IndexOf("diagrams/drawing", StringComparison.Ordinal) >= 0, xml => {
                Match shape = Regex.Match(xml, "<dsp:sp .*?</dsp:sp>", RegexOptions.Singleline, TimeSpan.FromSeconds(5));
                Assert.True(shape.Success);
                return xml.Replace("</dsp:spTree>",
                    "<dsp:extLst><a:ext uri=\"" + SmartArtExtensionUri + "\"><dsp:spTree>" + shape.Value + "</dsp:spTree>" + shape.Value +
                    "</a:ext></dsp:extLst></dsp:spTree>");
            });

            using var loaded = WordDocument.Load(filePath);
            Assert.Equal(expected, Assert.Single(loaded.SmartArts).GetPersistedLayoutShapes().Count);
        }

        private static void RewritePart(string filePath, string partName, Func<string, string> rewrite) =>
            RewritePart(filePath, name => name == partName, rewrite);

        private static void RewritePart(string filePath, Func<string, bool> match, Func<string, string> rewrite) {
            using var archive = ZipFile.Open(filePath, ZipArchiveMode.Update);
            ZipArchiveEntry entry = archive.Entries.Single(e => match(e.FullName));
            string name = entry.FullName;
            string xml;
            using (var reader = new StreamReader(entry.Open(), Encoding.UTF8)) xml = reader.ReadToEnd();
            string updated = rewrite(xml);
            Assert.NotEqual(xml, updated);
            entry.Delete();
            using var writer = new StreamWriter(archive.CreateEntry(name).Open(), new UTF8Encoding(false));
            writer.Write(updated);
        }
    }
}
