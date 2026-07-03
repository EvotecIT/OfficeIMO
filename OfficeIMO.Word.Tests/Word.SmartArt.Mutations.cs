using System.Linq;
using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void SmartArt_BasicProcess_AddInsertRemove_Clear() {
            string filePath = Path.Combine(_directoryWithFiles, "SmartArt_Basic_Mutations.docx");

            using (var document = WordDocument.Create(filePath)) {
                var sa = document.AddSmartArt(SmartArtType.BasicProcess);
                Assert.Equal(1, sa.NodeCount);

                sa.AddNode("B");
                sa.AddNode("C");
                Assert.Equal(3, sa.NodeCount);

                sa.InsertNodeAt(1, "AB");
                Assert.Equal(4, sa.NodeCount);

                sa.RemoveNodeAt(sa.NodeCount - 1);
                Assert.Equal(3, sa.NodeCount);

                sa.ReplaceTexts("A", "AB", "C");
                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                var sa = document.SmartArts.Single();
                Assert.Equal(3, sa.NodeCount);
                Assert.Equal("A", sa.GetNodeText(0));
                Assert.Equal("AB", sa.GetNodeText(1));
                Assert.Equal("C", sa.GetNodeText(2));
            }
        }

        [Fact]
        public void SmartArt_Cycle_AddInsertRemove() {
            string filePath = Path.Combine(_directoryWithFiles, "SmartArt_Cycle_Mutations.docx");

            using (var document = WordDocument.Create(filePath)) {
                var sa = document.AddSmartArt(SmartArtType.Cycle);
                Assert.Equal(1, sa.NodeCount);
                sa.AddNode("B");
                sa.AddNode("C");
                sa.AddNode("D");
                Assert.Equal(4, sa.NodeCount);
                sa.InsertNodeAt(1, "AB");
                Assert.Equal(5, sa.NodeCount);
                sa.RemoveNodeAt(3);
                Assert.Equal(4, sa.NodeCount);
                sa.ReplaceTexts(new [] {"Start", "AB", "C", "D"});
                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                var sa = document.SmartArts.Single();
                Assert.Equal(4, sa.NodeCount);
                Assert.Equal("Start", sa.GetNodeText(0));
                Assert.Equal("AB", sa.GetNodeText(1));
                Assert.Equal("C", sa.GetNodeText(2));
                Assert.Equal("D", sa.GetNodeText(3));
            }
        }

        [Fact]
        public void SmartArt_ReplaceTextsFormatted_Persists() {
            string filePath = Path.Combine(_directoryWithFiles, "SmartArt_Basic_Formatted.docx");

            using (var document = WordDocument.Create(filePath)) {
                var sa = document.AddSmartArt(SmartArtType.BasicProcess);
                // Ensure two nodes
                sa.AddNode("Second");
                sa.ReplaceTexts(bold: true, italic: true, underline: true, colorHex: "#FF0000", sizePt: 12, texts: new [] {"First", "Second"});
                document.Save(false);
            }

            // Inspect the XML to assert run properties persisted
            using (var zip = System.IO.Compression.ZipFile.OpenRead(filePath)) {
                var dataEntry = zip.Entries.First(e => e.FullName.StartsWith("graphics/") && e.FullName.Contains("data"));
                using var es = dataEntry.Open();
                var xdoc = System.Xml.Linq.XDocument.Load(es);
                var dgm = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/drawingml/2006/diagram";
                var a = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/drawingml/2006/main";
                var nodes = xdoc.Descendants(dgm + "pt").Where(p => (string?)p.Attribute("type") == null || (string?)p.Attribute("type") == "node").ToList();
                Assert.True(nodes.Count >= 2);
                foreach (var node in nodes.Take(2)) {
                    var p = (node.Element(dgm + "t") ?? node.Element(dgm + "txBody"))?.Element(a + "p");
                    Assert.NotNull(p);
                    var r = p!.Element(a + "r");
                    Assert.NotNull(r);
                    var rPr = r!.Element(a + "rPr");
                    Assert.NotNull(rPr);
                    Assert.Equal("1", (string?)rPr!.Attribute("b"));
                    Assert.Equal("1", (string?)rPr!.Attribute("i"));
                    Assert.Equal("sng", (string?)rPr!.Attribute("u"));
                    Assert.Equal("1200", (string?)rPr!.Attribute("sz"));
                    var clr = rPr!.Descendants(a + "srgbClr").FirstOrDefault();
                    Assert.NotNull(clr);
                    Assert.Equal("FF0000", (string?)clr!.Attribute("val"));
                }
            }
        }
    }
}
