using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_MergingDocumentsWithSeparateLists() {
            var filePath1 = Path.Combine(_directoryWithFiles, "MergeDoc1.docx");
            var filePath2 = Path.Combine(_directoryWithFiles, "MergeDoc2.docx");

            using (var doc1 = WordDocument.Create(filePath1)) {
                var list1 = doc1.AddList(WordListStyle.Numbered);
                list1.AddItem("Item 1");
                list1.AddItem("Item 2");
                doc1.Save();
            }

            using (var doc2 = WordDocument.Create(filePath2)) {
                var list2 = doc2.AddList(WordListStyle.Numbered);
                list2.AddItem("Second 1");
                list2.AddItem("Second 2");
                doc2.Save();
            }

            using (var doc1 = WordDocument.Load(filePath1))
            using (var doc2 = WordDocument.Load(filePath2)) {
                doc1.AppendDocument(doc2);
                doc1.Save();
            }

            using (var merged = WordDocument.Load(filePath1)) {
                Assert.Equal(2, merged.Lists.Count);
                var mainPart = merged._wordprocessingDocument.MainDocumentPart;
                Assert.NotNull(mainPart);
                var numberingPart = mainPart.NumberingDefinitionsPart;
                Assert.NotNull(numberingPart);
                var numbering = numberingPart.Numbering;
                Assert.NotNull(numbering);
                var numberingInstances = numbering.Elements<NumberingInstance>().ToList();
                Assert.All(numberingInstances, n => Assert.NotNull(n.NumberID));
                var ids = numberingInstances.Select(n => n.NumberID!.Value).Distinct().ToList();
                Assert.Equal(2, ids.Count);
            }
        }
    [Fact]

    public void Test_MergingDocumentsWithNestedLists() {
        var filePath1 = Path.Combine(_directoryWithFiles, "MergeDocNested1.docx");
        var filePath2 = Path.Combine(_directoryWithFiles, "MergeDocNested2.docx");

        using (var doc1 = WordDocument.Create(filePath1)) {
            var list1 = doc1.AddList(WordListStyle.Numbered);
            list1.AddItem("Item 1");
            list1.AddItem("Item 1.1", 1);
            doc1.Save();
        }

        using (var doc2 = WordDocument.Create(filePath2)) {
            var list2 = doc2.AddList(WordListStyle.Numbered);
            list2.AddItem("Item 2");
            list2.AddItem("Item 2.1", 1);
            doc2.Save();
        }

        using (var doc1 = WordDocument.Load(filePath1))
        using (var doc2 = WordDocument.Load(filePath2)) {
            doc1.AppendDocument(doc2);
            doc1.Save();
        }

        using (var merged = WordDocument.Load(filePath1)) {
            Assert.Equal(2, merged.Lists.Count);
            var mainPart = merged._wordprocessingDocument.MainDocumentPart;
            Assert.NotNull(mainPart);
            var numberingPart = mainPart.NumberingDefinitionsPart;
            Assert.NotNull(numberingPart);
            var numbering = numberingPart.Numbering;
            Assert.NotNull(numbering);
            var numberingInstances = numbering.Elements<NumberingInstance>().ToList();
            Assert.All(numberingInstances, n => Assert.NotNull(n.NumberID));
            var ids = numberingInstances.Select(n => n.NumberID!.Value).Distinct().ToList();
            Assert.Equal(2, ids.Count);
            Assert.Equal(4, merged.Paragraphs.Count(p => p.IsListItem));
        }
    }

    [Fact]
    public void Test_MergingDocumentsMultipleTimes() {
        var filePath1 = Path.Combine(_directoryWithFiles, "MergeDocMulti1.docx");
        var filePath2 = Path.Combine(_directoryWithFiles, "MergeDocMulti2.docx");
        var filePath3 = Path.Combine(_directoryWithFiles, "MergeDocMulti3.docx");

        using (var doc = WordDocument.Create(filePath1)) {
            var list = doc.AddList(WordListStyle.Numbered);
            list.AddItem("Item 1");
            doc.Save();
        }

        using (var doc = WordDocument.Create(filePath2)) {
            var list = doc.AddList(WordListStyle.Numbered);
            list.AddItem("Item 2");
            doc.Save();
        }

        using (var doc = WordDocument.Create(filePath3)) {
            var list = doc.AddList(WordListStyle.Numbered);
            list.AddItem("Item 3");
            doc.Save();
        }

        using (var baseDoc = WordDocument.Load(filePath1))
        using (var doc2 = WordDocument.Load(filePath2))
        using (var doc3 = WordDocument.Load(filePath3)) {
            baseDoc.AppendDocument(doc2);
            baseDoc.AppendDocument(doc3);
            baseDoc.Save();
        }

        using (var merged = WordDocument.Load(filePath1)) {
            var mainPart = merged._wordprocessingDocument.MainDocumentPart;
            Assert.NotNull(mainPart);
            var numberingPart = mainPart.NumberingDefinitionsPart;
            Assert.NotNull(numberingPart);
            var numbering = numberingPart.Numbering;
            Assert.NotNull(numbering);
            var numberingInstances = numbering.Elements<NumberingInstance>().ToList();
            Assert.All(numberingInstances, n => Assert.NotNull(n.NumberID));
            var ids = numberingInstances.Select(n => n.NumberID!.Value).Distinct().ToList();
            Assert.Equal(3, ids.Count);
        }
    }

    [Fact]
    public void Test_MergingDocumentWithoutLists() {
        var filePath1 = Path.Combine(_directoryWithFiles, "MergeDocNoLists1.docx");
        var filePath2 = Path.Combine(_directoryWithFiles, "MergeDocNoLists2.docx");

        using (var doc1 = WordDocument.Create(filePath1)) {
            var list = doc1.AddList(WordListStyle.Numbered);
            list.AddItem("Item 1");
            doc1.Save();
        }

        using (var doc2 = WordDocument.Create(filePath2)) {
            doc2.AddParagraph("Just text");
            doc2.Save();
        }

        using (var doc1 = WordDocument.Load(filePath1))
        using (var doc2 = WordDocument.Load(filePath2)) {
            doc1.AppendDocument(doc2);
            doc1.Save();
        }

        using (var merged = WordDocument.Load(filePath1)) {
            Assert.Single(merged.Lists);
            Assert.Contains(merged.Paragraphs, p => p.Text == "Just text");
        }
    }

    [Fact]
    public void Test_MergingDocumentsPreservesRelationships() {
        var destinationPath = Path.Combine(_directoryWithFiles, "MergeDocRelationshipsBase.docx");
        var sourcePath = Path.Combine(_directoryWithFiles, "MergeDocRelationshipsSource.docx");

        string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
        string excelFilePath = Path.Combine(_directoryDocuments, "SampleFileExcel.xlsx");
        string iconPath = Path.Combine(_directoryDocuments, "SampleExcelIcon.png");
        var hyperlinkUri = new Uri("https://openai.com/");

        using (var doc = WordDocument.Create(destinationPath)) {
            doc.AddParagraph("Base document");
            doc.Save();
        }

        using (var doc = WordDocument.Create(sourcePath)) {
            doc.AddParagraph("Image paragraph").AddImage(imagePath, width: 64, height: 64);
            doc.AddHyperLink("OpenAI", hyperlinkUri);
            doc.AddParagraph("Embedded object");
            doc.AddEmbeddedObject(excelFilePath, iconPath);
            doc.Save();
        }

        using (var destination = WordDocument.Load(destinationPath))
        using (var source = WordDocument.Load(sourcePath)) {
            destination.AppendDocument(source);
            destination.Save();
        }

        using (var merged = WordDocument.Load(destinationPath)) {
            var mainPart = merged._wordprocessingDocument.MainDocumentPart;
            Assert.NotNull(mainPart);

            Assert.NotNull(mainPart!.Document?.Body);

            var blip = mainPart.Document!.Body!.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
            Assert.NotNull(blip);
            Assert.False(string.IsNullOrEmpty(blip!.Embed?.Value));
            Assert.NotNull(mainPart.GetPartById(blip.Embed!.Value!));

            var hyperlink = mainPart.Document.Body.Descendants<Hyperlink>().FirstOrDefault();
            Assert.NotNull(hyperlink);
            Assert.False(string.IsNullOrEmpty(hyperlink!.Id));
            var hyperlinkRel = mainPart.HyperlinkRelationships.FirstOrDefault(h => h.Id == hyperlink.Id!);
            Assert.NotNull(hyperlinkRel);
            Assert.Equal(hyperlinkUri, hyperlinkRel!.Uri);

            var oleObject = mainPart.Document.Body.Descendants<OleObject>().FirstOrDefault();
            Assert.NotNull(oleObject);
            Assert.False(string.IsNullOrEmpty(oleObject!.Id?.Value));
            Assert.NotNull(mainPart.GetPartById(oleObject.Id!.Value!));

            Assert.True(mainPart.ImageParts.Any());
            Assert.True(mainPart.EmbeddedPackageParts.Any());
        }
    }
}
}
