using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void SmartArt_MoveNode_ChangesLogicalOrderAcrossReload() {
            string filePath = Path.Combine(_directoryWithFiles, "SmartArt.MoveNode.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordSmartArt smartArt = document.AddSmartArt(SmartArtType.BasicProcess);
                smartArt.SetNodeText(0, "A");
                smartArt.AddNode("B");
                smartArt.AddNode("C");

                smartArt.MoveNode(0, 2);

                Assert.Equal("B", smartArt.GetNodeText(0));
                Assert.Equal("C", smartArt.GetNodeText(1));
                Assert.Equal("A", smartArt.GetNodeText(2));
                document.Save();
            }

            using WordDocument reloaded = WordDocument.Load(filePath);
            WordSmartArt imported = Assert.Single(reloaded.SmartArts);
            Assert.Equal(new[] { "B", "C", "A" },
                Enumerable.Range(0, imported.NodeCount).Select(imported.GetNodeText));
            Assert.Empty(new OpenXmlValidator().Validate(reloaded._wordprocessingDocument));
        }

        [Fact]
        public void SmartArt_DuplicateNodeModelIds_FallBackToDocumentOrder() {
            string filePath = Path.Combine(_directoryWithFiles, "SmartArt.DuplicateNodeModelIds.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordSmartArt createdSmartArt = document.AddSmartArt(SmartArtType.BasicProcess);
                createdSmartArt.SetNodeText(0, "A");
                createdSmartArt.AddNode("B");
                createdSmartArt.AddNode("C");

                DiagramDataPart dataPart = document._wordprocessingDocument.MainDocumentPart!.DiagramDataParts.Single();
                XDocument data = LoadDiagramData(dataPart);
                XNamespace dgm = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
                var nodes = data.Descendants(dgm + "pt")
                    .Where(point => (string?)point.Attribute("type") is null or "node")
                    .Where(point => point.Element(dgm + "t") != null)
                    .ToList();
                nodes[1].SetAttributeValue("modelId", (string)nodes[0].Attribute("modelId")!);
                SaveDiagramData(dataPart, data);
                document.Save();
            }

            using WordDocument reloaded = WordDocument.Load(filePath);
            WordSmartArt smartArt = Assert.Single(reloaded.SmartArts);
            Assert.Equal(3, smartArt.NodeCount);
            Assert.Equal(new[] { "A", "B", "C" },
                Enumerable.Range(0, smartArt.NodeCount).Select(smartArt.GetNodeText));
            Assert.Contains("non-empty and unique", Assert.Throws<InvalidOperationException>(() => smartArt.AddNode("D")).Message);
            Assert.Contains("non-empty and unique", Assert.Throws<InvalidOperationException>(() => smartArt.InsertNodeAt(1, "D")).Message);
            Assert.Contains("non-empty and unique", Assert.Throws<InvalidOperationException>(() => smartArt.RemoveNodeAt(1)).Message);
            Assert.Contains("non-empty and unique", Assert.Throws<InvalidOperationException>(() => smartArt.MoveNode(0, 2)).Message);
        }

        [Fact]
        public void SmartArt_MoveNode_UpdatesDuplicateDocumentChildConnections() {
            string filePath = Path.Combine(_directoryWithFiles, "SmartArt.DuplicateDocumentChildConnections.docx");
            using WordDocument document = WordDocument.Create(filePath);
            WordSmartArt smartArt = document.AddSmartArt(SmartArtType.BasicProcess);
            smartArt.SetNodeText(0, "A");
            smartArt.AddNode("B");
            smartArt.AddNode("C");

            DiagramDataPart dataPart = document._wordprocessingDocument.MainDocumentPart!.DiagramDataParts.Single();
            XDocument data = LoadDiagramData(dataPart);
            XNamespace dgm = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
            string docId = (string)data.Descendants(dgm + "pt")
                .Single(point => (string?)point.Attribute("type") == "doc")
                .Attribute("modelId")!;
            XElement connection = data.Descendants(dgm + "cxn")
                .First(item => (string?)item.Attribute("srcId") == docId);
            XElement duplicate = new XElement(connection);
            duplicate.SetAttributeValue("modelId", "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}");
            duplicate.SetAttributeValue("srcOrd", 99);
            connection.AddAfterSelf(duplicate);
            SaveDiagramData(dataPart, data);

            smartArt.MoveNode(0, 2);

            XDocument updated = LoadDiagramData(dataPart);
            string duplicateDestination = (string)duplicate.Attribute("destId")!;
            var duplicateConnections = updated.Descendants(dgm + "cxn")
                .Where(item => (string?)item.Attribute("srcId") == docId)
                .Where(item => (string?)item.Attribute("destId") == duplicateDestination)
                .ToList();
            Assert.Equal(2, duplicateConnections.Count);
            Assert.All(duplicateConnections,
                item => Assert.Equal(2, (int)item.Attribute("srcOrd")!));
            Assert.Equal(new[] { "B", "C", "A" },
                Enumerable.Range(0, smartArt.NodeCount).Select(smartArt.GetNodeText));
        }

        [Fact]
        public void SmartArt_InsertAndRemove_ResequenceDuplicateDocumentChildConnections() {
            string filePath = Path.Combine(_directoryWithFiles, "SmartArt.ResequenceDuplicateDocumentChildConnections.docx");
            string docId;
            string duplicateDestination;
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordSmartArt smartArt = document.AddSmartArt(SmartArtType.BasicProcess);
                smartArt.SetNodeText(0, "A");
                smartArt.AddNode("B");
                smartArt.AddNode("C");

                DiagramDataPart dataPart = document._wordprocessingDocument.MainDocumentPart!.DiagramDataParts.Single();
                XDocument data = LoadDiagramData(dataPart);
                XNamespace dgm = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
                docId = (string)data.Descendants(dgm + "pt")
                    .Single(point => (string?)point.Attribute("type") == "doc")
                    .Attribute("modelId")!;
                XElement connection = data.Descendants(dgm + "cxn")
                    .Single(item => (string?)item.Attribute("srcId") == docId && (int?)item.Attribute("srcOrd") == 1);
                duplicateDestination = (string)connection.Attribute("destId")!;
                XElement duplicate = new XElement(connection);
                duplicate.SetAttributeValue("modelId", "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}");
                duplicate.SetAttributeValue("srcOrd", 99);
                connection.AddAfterSelf(duplicate);
                SaveDiagramData(dataPart, data);

                smartArt.InsertNodeAt(1, "X");
                smartArt.RemoveNodeAt(0);

                Assert.Equal(new[] { "X", "B", "C" },
                    Enumerable.Range(0, smartArt.NodeCount).Select(smartArt.GetNodeText));
                AssertDuplicateConnectionsShareOrder(LoadDiagramData(dataPart), dgm, docId, duplicateDestination, 1);
                document.Save();
            }

            using WordDocument reloaded = WordDocument.Load(filePath);
            WordSmartArt imported = Assert.Single(reloaded.SmartArts);
            Assert.Equal(new[] { "X", "B", "C" },
                Enumerable.Range(0, imported.NodeCount).Select(imported.GetNodeText));
            DiagramDataPart reloadedDataPart = reloaded._wordprocessingDocument.MainDocumentPart!.DiagramDataParts.Single();
            AssertDuplicateConnectionsShareOrder(
                LoadDiagramData(reloadedDataPart),
                "http://schemas.openxmlformats.org/drawingml/2006/diagram",
                docId,
                duplicateDestination,
                1);
        }

        [Fact]
        public void SmartArt_RemoveNode_RemovesIncomingAndOutgoingConnections() {
            string filePath = Path.Combine(_directoryWithFiles, "SmartArt.RemoveIncidentConnections.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordSmartArt smartArt = document.AddSmartArt(SmartArtType.Cycle);
                smartArt.SetNodeText(0, "A");
                smartArt.AddNode("B");
                smartArt.AddNode("C");

                DiagramDataPart dataPart = document._wordprocessingDocument.MainDocumentPart!.DiagramDataParts.Single();
                XDocument data = LoadDiagramData(dataPart);
                XNamespace dgm = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
                var nodes = data.Descendants(dgm + "pt")
                    .Where(point => (string?)point.Attribute("type") is null or "node")
                    .Where(point => point.Element(dgm + "t") != null)
                    .ToList();
                string removedId = (string)nodes[1].Attribute("modelId")!;
                string remainingId = (string)nodes[2].Attribute("modelId")!;
                data.Descendants(dgm + "cxnLst").Single().Add(
                    new XElement(dgm + "cxn",
                        new XAttribute("modelId", "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}"),
                        new XAttribute("srcId", removedId),
                        new XAttribute("destId", remainingId),
                        new XAttribute("srcOrd", 0),
                        new XAttribute("destOrd", 1)));
                SaveDiagramData(dataPart, data);

                smartArt.RemoveNodeAt(1);

                XDocument updated = LoadDiagramData(dataPart);
                Assert.DoesNotContain(updated.Descendants(dgm + "cxn"), connection =>
                    (string?)connection.Attribute("srcId") == removedId ||
                    (string?)connection.Attribute("destId") == removedId);
                Assert.Equal(new[] { 0, 1 }, updated.Descendants(dgm + "cxn")
                    .Where(connection => (string?)connection.Attribute("srcId") ==
                        (string?)updated.Descendants(dgm + "pt").Single(point => (string?)point.Attribute("type") == "doc").Attribute("modelId"))
                    .Select(connection => (int)connection.Attribute("srcOrd")!)
                    .OrderBy(value => value));
                document.Save();
            }

            using WordDocument reloaded = WordDocument.Load(filePath);
            Assert.Equal(new[] { "A", "C" },
                Enumerable.Range(0, Assert.Single(reloaded.SmartArts).NodeCount)
                    .Select(Assert.Single(reloaded.SmartArts).GetNodeText));
            Assert.Empty(new OpenXmlValidator().Validate(reloaded._wordprocessingDocument));
        }

        private static XDocument LoadDiagramData(DiagramDataPart part) {
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            return XDocument.Load(stream);
        }

        private static void SaveDiagramData(DiagramDataPart part, XDocument data) {
            using var stream = new MemoryStream();
            data.Save(stream);
            stream.Position = 0;
            part.FeedData(stream);
        }

        private static void AssertDuplicateConnectionsShareOrder(
            XDocument data,
            XNamespace dgm,
            string docId,
            string destinationId,
            int expectedOrder) {
            var connections = data.Descendants(dgm + "cxn")
                .Where(item => (string?)item.Attribute("srcId") == docId)
                .Where(item => (string?)item.Attribute("destId") == destinationId)
                .ToList();
            Assert.Equal(2, connections.Count);
            Assert.All(connections, item => Assert.Equal(expectedOrder, (int)item.Attribute("srcOrd")!));
        }
    }
}
