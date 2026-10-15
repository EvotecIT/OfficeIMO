using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.OpenDocument;
using OfficeIMO.Reader.OpenDocument;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderOpenDocumentInspectionEvidenceTests {
    private static readonly byte[] Pixel = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
    private static OdfRect Box => new(OdfLength.Centimeters(1), OdfLength.Centimeters(1), OdfLength.Centimeters(2), OdfLength.Centimeters(2));
    private static ReaderOptions Inspection() => new() { Inspection = new ReaderInspectionOptions { IncludeContentSafety = false } };

    [Fact]
    public void TextStoriesRetainDistinctHeaderFootnoteLocationsAndShareInertImagePayloads() {
        OdtDocument source = OdtDocument.Create();
        OdtParagraph body = source.AddParagraph("Body"); body.AddImage(Pixel, "pixel.png", Box.Width, Box.Height);
        OdtNote note = body.AddFootnote("Note source"); note.Paragraphs[0].AddHyperlink("Note target", "https://example.invalid/note");
        OdtParagraph header = source.PageLayout.Header.AddParagraph("Header source");
        header.AddHyperlink("Header target", "https://example.invalid/header");
        header.AddImage(Pixel, "pixel.png", Box.Width, Box.Height);
        OfficeDocumentReadResult result = new OfficeDocumentReaderBuilder().AddOpenDocumentHandler().Build().ReadDocument(source.ToBytes(), "stories.odt", Inspection());
        Assert.Contains(result.Blocks, block => block.Text.Contains("Header source") && block.Location.SourceBlockKind == "header" && block.Location.BlockAnchor != null);
        Assert.Contains(result.Blocks, block => block.Text.Contains("Note source") && block.Location.SourceBlockKind == "footnote");
        Assert.Contains(result.Links, link => link.Uri == "https://example.invalid/note" && link.Location.SourceBlockKind == "footnote");
        Assert.Equal(2, result.Assets.Count);
        Assert.Same(result.Assets[0].PayloadBytes, result.Assets[1].PayloadBytes);
        Assert.Empty(result.Pages);
        Assert.Contains(OfficeDocumentReadResultJson.Deserialize(result.ToJson()).Blocks, block => block.Text.Contains("Header source") && block.Location.SourceBlockKind == "header");
    }

    [Fact]
    public void PresentationRetainsSlideNotesTablesImagesLinksAndUnreferencedMasters() {
        OdpPresentation source = OdpPresentation.Create(); OdpSlide slide = source.AddSlide("Slide source");
        OdpTextBox text = slide.AddTextBox(Box, "Body source"); text.Paragraphs[0].AddHyperlink("Slide target", "https://example.invalid/slide");
        OdpImage image = slide.AddImage(Pixel, "pixel.png", Box); slide.GetOrCreateSpeakerNotes().AddParagraph("Notes source").AddHyperlink("Notes target", "https://example.invalid/notes");
        OdpTable table = slide.AddTable(Box, 1, 1); table.Cell(0, 0).AddParagraph("Table source").AddHyperlink("Cell target", "https://example.invalid/cell");
        XNamespace office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0", style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        XNamespace draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0", tx = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        XNamespace xlink = "http://www.w3.org/1999/xlink", svg = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
        byte[] bytes = Rewrite(source.ToBytes(), "styles.xml", xml => xml.Root!.Element(office + "master-styles")!.Add(
            new XElement(style + "master-page", new XAttribute(style + "name", "UnusedMaster"),
                new XElement(draw + "frame", new XAttribute(draw + "name", "MasterText"), new XAttribute(svg + "width", "2cm"), new XAttribute(svg + "height", "2cm"),
                    new XElement(draw + "text-box", new XElement(tx + "p", "Master source", new XElement(tx + "a", new XAttribute(xlink + "href", "https://example.invalid/master"), "Master target")))),
                new XElement(draw + "frame", new XAttribute(draw + "name", "MasterImage"), new XAttribute(svg + "width", "2cm"), new XAttribute(svg + "height", "2cm"),
                    new XElement(draw + "image", new XAttribute(xlink + "href", image.Path))))));
        OfficeDocumentReadResult result = new OfficeDocumentReaderBuilder().AddOpenDocumentHandler().Build().ReadDocument(bytes, "stories.odp", Inspection());
        Assert.Contains(result.Blocks, block => block.Text.Contains("Master source") && block.Location.SourceBlockKind == "master-shape" && block.Location.Slide == null && block.Location.BlockAnchor!.StartsWith("styles.xml#UnusedMaster/"));
        Assert.Contains(result.Blocks, block => block.Text.Contains("Notes source") && block.Location.SourceBlockKind == "speaker-notes");
        Assert.Equal(4, result.Links.Count); Assert.Equal(2, result.Assets.Count); Assert.Single(result.Pages);
        Assert.Same(result.Assets[0].PayloadBytes, result.Assets[1].PayloadBytes);
        Assert.Contains(OfficeDocumentReadResultJson.Deserialize(result.ToJson()).Assets, asset => asset.Location.SourceBlockKind == "master-shape" && asset.PayloadHash != null);
        Assert.Equal(3, result.Pages[0].Links.Count); Assert.Single(result.Pages[0].Assets);
        Assert.Contains(result.Metadata, item => item.Name == "form-widgets" && item.Value == "owner-model-unavailable");
    }

    [Fact]
    public void SpreadsheetSparseInspectionKeepsFormulaCachesAnnotationsRepeatRangesAndVisibility() {
        OdsDocument source = OdsDocument.Create(); OdsSheet sheet = source.AddSheet("Sparse source");
        sheet.Cell(0, 0).Formula = "of:=1+1"; sheet.Cell(0, 0).AddAnnotation("Annotation source", "Ada");
        sheet.Row(0).Hidden = true; sheet.Column(0).Hidden = true;
        XNamespace table = "urn:oasis:names:tc:opendocument:xmlns:table:1.0", office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        byte[] bytes = Rewrite(source.ToBytes(), "content.xml", xml => {
            XElement row = xml.Descendants(table + "table-row").Single(); row.SetAttributeValue(table + "number-rows-repeated", 1000);
            XElement cell = row.Element(table + "table-cell")!; cell.SetAttributeValue(table + "number-columns-repeated", 3);
            cell.SetAttributeValue(office + "value-type", "float"); cell.SetAttributeValue(office + "value", "99");
        });
        OfficeDocumentReadResult result = new OfficeDocumentReaderBuilder().AddOpenDocumentHandler(new ReaderOpenDocumentOptions { A1Range = "B2:C3", HeadersInFirstRow = false }).Build().ReadDocument(bytes, "sparse.ods", Inspection());
        OfficeDocumentMetadataEntry formula = Assert.Single(result.Metadata, item => item.Category == OfficeDocumentInspectionCategories.Formula);
        Assert.Equal("of:=1+1", formula.Value); Assert.Equal("99", formula.Attributes["storedValue"]); Assert.Equal("false", formula.Attributes["calculated"]);
        Assert.Equal("B2:C3", formula.Location!.A1Range); Assert.Equal("1000", formula.Attributes["rowRepeatCount"]);
        Assert.Single(result.Metadata, item => item.Category == OfficeDocumentInspectionCategories.Comment && item.Value == "Annotation source");
        Assert.Contains(result.Metadata, item => item.Category == OfficeDocumentInspectionCategories.Visibility && item.Name == "Row prototype" && item.Attributes["repeatCount"] == "1000");
        Assert.Contains(result.Metadata, item => item.Category == OfficeDocumentInspectionCategories.Visibility && item.Name == "Column prototype");
    }

    [Fact]
    public void DrawingRetainsEveryMasterAndLayerStateAndPayloadBoundsAreTerminal() {
        OdgDocument source = OdgDocument.Create(); OdgPage page = source.AddPage("Page source");
        page.Shapes.AddImage(Pixel, "pixel.png", Box); page.MasterShapes.AddTextBox(Box, "Master source");
        page.MasterShapes.AddImage(Pixel, "pixel.png", Box);
        page.Layers.Add("Hidden source", OdgLayerDisplay.None); source.CloneMasterPage(page.MasterPageName, "UnusedMaster");
        byte[] bytes = source.ToBytes();
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddOpenDocumentHandler(new ReaderOpenDocumentOptions { MaxTotalImageBytes = Pixel.Length }).Build();
        OfficeDocumentReadResult result = reader.ReadDocument(bytes, "stories.odg", Inspection());
        Assert.Equal(3, result.Assets.Count); Assert.Same(result.Assets[0].PayloadBytes, result.Assets[2].PayloadBytes);
        Assert.Equal(2, result.Blocks.Count(block => block.Text == "Master source" && block.Location.SourceBlockKind == "master-shape"));
        Assert.Contains(result.Metadata, item => item.Name == "Layer" && item.SourceObjectId == "Hidden source" && item.Attributes["visibleOnScreen"] == "false");
        Assert.Single(result.Pages); Assert.True(result.Pages[0].Width > 0);
        OfficeDocumentReader bounded = new OfficeDocumentReaderBuilder().AddOpenDocumentHandler(new ReaderOpenDocumentOptions { MaxImageBytes = Pixel.Length - 1 }).Build();
        Assert.Equal("OpenDocument.MaxImageBytes", Assert.Throws<ReaderResourceLimitException>(() => bounded.ReadDocument(bytes, "stories.odg")).LimitName);
        OfficeDocumentReader noLinks = new OfficeDocumentReaderBuilder().AddOpenDocumentHandler(new ReaderOpenDocumentOptions { MaxLinks = 0 }).Build();
        OdtDocument linked = OdtDocument.Create(); linked.AddParagraph().AddHyperlink("Source", "https://example.invalid/target");
        Assert.Equal("OpenDocument.MaxLinks", Assert.Throws<ReaderResourceLimitException>(() => noLinks.ReadDocument(linked.ToBytes(), "links.odt")).LimitName);
    }

    private static byte[] Rewrite(byte[] source, string entryName, Action<XDocument> mutate) {
        using var memory = new MemoryStream(); memory.Write(source, 0, source.Length); memory.Position = 0;
        using (var archive = new ZipArchive(memory, ZipArchiveMode.Update, leaveOpen: true)) {
            ZipArchiveEntry entry = archive.GetEntry(entryName)!; XDocument xml;
            using (Stream input = entry.Open()) xml = XDocument.Load(input);
            mutate(xml); entry.Delete(); using Stream output = archive.CreateEntry(entryName).Open(); xml.Save(output);
        }
        return memory.ToArray();
    }
}
