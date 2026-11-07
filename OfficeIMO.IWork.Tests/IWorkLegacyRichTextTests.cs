using System.Xml.Linq;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkLegacyRichTextTests {
    private static readonly byte[] Image = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

    [Fact]
    public void Pages09InheritedStylesBulletsHyperlinksAndAnchoredImageSurviveSavedDocx() {
        using MemoryStream package = CreatePackage();
        using IWorkSourceDocument source = IWorkSourceDocument.Load(package);
        IWorkPagesProjection projection = source.ReadPages();
        IWorkTextParagraph first = projection.Body.Paragraphs[0];
        Assert.Equal(IWorkTextAlignment.Center, first.Style.Alignment);
        IWorkTextRun emphasis = first.Runs.Single(run => run.Text == "Emphasis");
        Assert.True(emphasis.Style.Bold);
        Assert.True(emphasis.Style.Italic);
        Assert.Equal(14, emphasis.Style.FontSizePoints);
        Assert.Equal("Arial", emphasis.Style.FontName);
        Assert.Equal("FF0000", emphasis.Style.Color!.RgbHex);
        Assert.Equal("https://example.com/resource", first.Runs.Single(run => run.Text == "Link").Hyperlink);
        IWorkTextParagraph bullet = projection.Body.Paragraphs[1];
        Assert.Equal(0, bullet.ListLevel);
        Assert.Equal("•", bullet.ListLabel);
        Assert.Equal(IWorkListMarkerKind.Text, bullet.ListMarkerKind);
        IWorkImageAsset image = Assert.Single(projection.Images);
        Assert.Equal(Image, image.GetBytes());
        Assert.Equal((72d, 36d), (image.Geometry!.WidthPoints, image.Geometry.HeightPoints));
        Assert.NotNull(projection.Body.Paragraphs[2].Runs.Single(run => run.InlineObject != null).InlineObject);

        using var result = source.ToWordResult(new IWorkConversionOptions {
            Mode = IWorkConversionMode.EditableOnly,
            AllowPartialEditableReconstruction = true
        });
        Assert.False(result.IsVisualFallback);
        using var output = new MemoryStream();
        result.Value.Save(output);
        output.Position = 0;
        using var zip = new ZipArchive(output, ZipArchiveMode.Read);
        using Stream documentXml = zip.GetEntry("word/document.xml")!.Open();
        XDocument document = XDocument.Load(documentXml);
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        XElement targetRun = Assert.Single(document.Descendants(w + "r"), run =>
            run.Element(w + "t")?.Value == "Emphasis");
        Assert.NotNull(targetRun.Element(w + "rPr")?.Element(w + "b"));
        Assert.NotNull(targetRun.Element(w + "rPr")?.Element(w + "i"));
        Assert.Equal("28", (string?)targetRun.Element(w + "rPr")?.Element(w + "sz")?.Attribute(w + "val"));
        Assert.Contains(document.Descendants(w + "p"), paragraph =>
            paragraph.Descendants(w + "t").Any(text => text.Value == "Bullet")
            && paragraph.Element(w + "pPr")?.Element(w + "numPr") != null);
        XElement anchored = Assert.Single(document.Descendants(w + "p"), paragraph =>
            paragraph.Descendants(w + "t").Any(text => text.Value == "Before"));
        Assert.Equal("BeforeAfter", string.Concat(anchored.Descendants(w + "t").Select(text => text.Value)));
        XElement extent = Assert.Single(anchored.Descendants(wp + "extent"));
        Assert.Equal("914400", (string?)extent.Attribute("cx"));
        Assert.Equal("457200", (string?)extent.Attribute("cy"));
        using var relationshipStream = zip.GetEntry("word/_rels/document.xml.rels")!.Open();
        XDocument relationships = XDocument.Load(relationshipStream);
        XElement imageRelationship = Assert.Single(relationships.Root!.Elements(), relationship =>
            ((string?)relationship.Attribute("Type"))?.EndsWith("/image", StringComparison.Ordinal) == true);
        var imageTarget = new Uri(new Uri("https://package.example/word/document.xml"),
            (string)imageRelationship.Attribute("Target")!);
        ZipArchiveEntry payload = zip.GetEntry(imageTarget.AbsolutePath.TrimStart('/'))!;
        using var imageBytes = new MemoryStream();
        using (Stream data = payload.Open()) data.CopyTo(imageBytes);
        Assert.Equal(Image, imageBytes.ToArray());
        Assert.Contains(relationships.Root.Elements(), relationship =>
            (string?)relationship.Attribute("Target") == "https://example.com/resource");
    }

    [Fact]
    public void SelectedXmlStyleCyclesFailBeforeEditableProjection() {
        using MemoryStream package = CreatePackage(cycle: true);
        using IWorkSourceDocument source = IWorkSourceDocument.Load(package);
        var error = Assert.Throws<InvalidDataException>(() => source.ReadPages());
        Assert.Contains("cycle", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void XmlImageProjectionUsesTheSourceWideDecodedImageBudget() {
        using MemoryStream package = CreatePackage();
        using IWorkSourceDocument source = IWorkSourceDocument.Load(package,
            new IWorkLoadOptions { MaxDecodedImageBytes = 3 });
        var error = Assert.Throws<InvalidDataException>(() => source.ReadPages());
        Assert.Contains("Decoded image data", error.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("count")]
    [InlineData("encoded")]
    [InlineData("decoded")]
    public void XmlSharedImageOccurrencesRetainBudgetsAndUniqueResourceSemantics(string boundary) {
        string[]? sharedDiagnostics = null;
        foreach (bool shared in new[] { true, false }) {
            using MemoryStream package = RepeatedImagePackage(shared);
            var limits = new IWorkLoadOptions {
                MaxProjectedImages = 2,
                MaxProjectedImageBytes = 2L * Image.Length,
                MaxDecodedImageBytes = 8
            };
            using IWorkSourceDocument source = IWorkSourceDocument.Load(package, limits);
            IWorkPagesProjection projection = source.ReadPages();
            Assert.Equal("EmphasisLink\nBullet\nBeforeAfter\nSecondTail", projection.Body.PlainText);
            Assert.Equal(2, projection.Images.Count);
            Assert.Equal(shared ? 1 : 2, projection.Images.Select(image => image.PackagePath).Distinct().Count());
            Assert.All(projection.Images, image => Assert.Equal(Image, image.GetBytes()));
            Assert.Equal(new[] { (72d, 36d), (36d, 18d) }, projection.Images
                .Select(image => (image.Geometry!.WidthPoints, image.Geometry.HeightPoints)));
            Assert.Equal(projection.Images.Select(image => image.SourceIdentity!.RecordIdentifier),
                projection.Body.Paragraphs.SelectMany(paragraph => paragraph.Runs)
                    .Where(run => run.InlineObject != null).Select(run => run.InlineObject!.Drawable.RecordIdentifier));
            using var result = source.ToWordResult(new IWorkConversionOptions {
                Mode = IWorkConversionMode.EditableOnly,
                AllowPartialEditableReconstruction = true
            });
            string[] diagnostics = result.Report.Diagnostics.Select(item => item.Code + ":" + item.LossKind).ToArray();
            if (shared) sharedDiagnostics = diagnostics;
            else Assert.Equal(sharedDiagnostics, diagnostics);
            using var output = new MemoryStream();
            result.Value.Save(output);
            output.Position = 0;
            using (var saved = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(output, false)) {
                var main = saved.MainDocumentPart!;
                Assert.Equal(projection.Body.Paragraphs.Select(paragraph => paragraph.Text),
                    main.Document!.Body!.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()
                        .Select(paragraph => paragraph.InnerText));
                Assert.Equal(new[] { (914400L, 457200L), (457200L, 228600L) },
                    main.Document!.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent>()
                        .Select(extent => (extent.Cx!.Value, extent.Cy!.Value)));
                var pictures = main.Document!.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().ToArray();
                Assert.Equal(2, pictures.Length);
                foreach (var picture in pictures) {
                    var part = Assert.IsType<DocumentFormat.OpenXml.Packaging.ImagePart>(main.GetPartById(picture.Embed!.Value!));
                    using var payload = new MemoryStream();
                    using (Stream image = part.GetStream()) image.CopyTo(payload);
                    Assert.Equal(Image, payload.ToArray());
                }
            }

            if (boundary == "count") limits.MaxProjectedImages--;
            else if (boundary == "encoded") limits.MaxProjectedImageBytes--;
            else limits.MaxDecodedImageBytes--;
            package.Position = 0;
            using IWorkSourceDocument limited = IWorkSourceDocument.Load(package, limits);
            var error = Assert.Throws<InvalidDataException>(() => limited.ReadPages());
            Assert.Contains(boundary == "count" ? "iWork image count" : boundary == "encoded"
                ? "Projected destination image data" : "Decoded image data", error.Message, StringComparison.Ordinal);
        }
    }

    private static MemoryStream RepeatedImagePackage(bool shared) {
        MemoryStream package = CreatePackage();
        XNamespace sf = "http://developer.apple.com/namespaces/sf";
        XNamespace sfa = "http://developer.apple.com/namespaces/sfa";
        using (var zip = new ZipArchive(package, ZipArchiveMode.Update, leaveOpen: true)) {
            ZipArchiveEntry index = zip.GetEntry("index.xml")!;
            XDocument document;
            using (Stream xml = index.Open()) document = XDocument.Load(xml, LoadOptions.PreserveWhitespace);
            XElement storage = document.Root!.Element(sf + "text-storage")!;
            XElement attachments = storage.Element(sf + "attachments")!;
            var second = new XElement(attachments.Elements().Single());
            second.SetAttributeValue(sfa + "ID", "second-attachment");
            second.Element(sf + "media")!.SetAttributeValue(sfa + "ID", "second-image");
            XElement size = second.Descendants(sf + "size").Single();
            size.SetAttributeValue(sfa + "w", "36");
            size.SetAttributeValue(sfa + "h", "18");
            if (!shared) second.Descendants(sf + "data").Single().SetAttributeValue(sf + "path", "second.png");
            attachments.Add(second);
            storage.Element(sf + "text-body")!.Add(new XElement(sf + "p", "Second",
                new XElement(sf + "attachment-ref", new XAttribute(sfa + "IDREF", "second-attachment")), "Tail"));
            index.Delete();
            using (Stream xml = zip.CreateEntry("index.xml").Open()) document.Save(xml);
            if (!shared) {
                using Stream image = zip.CreateEntry("second.png").Open();
                image.Write(Image, 0, Image.Length);
            }
        }
        package.Position = 0;
        return package;
    }

    private static MemoryStream CreatePackage(bool cycle = false) {
        string parent = cycle ? "loop" : "base";
        string xml = $$"""
            <sl:document xmlns:sl="http://developer.apple.com/namespaces/sl"
              xmlns:sf="http://developer.apple.com/namespaces/sf" xmlns:sfa="http://developer.apple.com/namespaces/sfa"
              sl:version="92008102400">
              <sl:slprint-info sl:page-width="612" sl:page-height="792"><sf:page-margins
                sf:left="72" sf:right="72" sf:top="72" sf:bottom="72" sf:header="36" sf:footer="36"/></sl:slprint-info>
              <sl:stylesheet sfa:ID="styles"><sf:styles>
                <sf:paragraphstyle sf:ident="base"><sf:property-map>
                  <sf:fontName><sf:string sfa:string="Arial"/></sf:fontName><sf:fontSize><sf:number sfa:number="14"/></sf:fontSize>
                  <sf:fontColor><sf:color sfa:r="1" sfa:g="0" sfa:b="0" sfa:a="1"/></sf:fontColor>
                  <sf:alignment><sf:number sfa:number="2"/></sf:alignment>
                </sf:property-map></sf:paragraphstyle>
                <sf:paragraphstyle sf:ident="selected" sf:parent-ident="{{parent}}"><sf:property-map/></sf:paragraphstyle>
                <sf:paragraphstyle sf:ident="loop" sf:parent-ident="selected"><sf:property-map/></sf:paragraphstyle>
                <sf:characterstyle sf:ident="emphasis-base"><sf:property-map><sf:italic><sf:number sfa:number="1"/></sf:italic></sf:property-map></sf:characterstyle>
                <sf:characterstyle sf:ident="emphasis" sf:parent-ident="emphasis-base"><sf:property-map><sf:bold><sf:number sfa:number="1"/></sf:bold></sf:property-map></sf:characterstyle>
                <sf:liststyle sfa:ID="bullets"><sf:property-map><sf:listLabelTypes><sf:array>
                  <sf:list-label-typeinfo sf:type="none"/><sf:list-label-typeinfo sf:type="text"><sf:text-label sf:type="bullet" sf:format="•" sf:first="0"/></sf:list-label-typeinfo>
                </sf:array></sf:listLabelTypes></sf:property-map></sf:liststyle>
                <sf:paragraphstyle sf:ident="listed" sf:parent-ident="base"><sf:property-map><sf:listStyle><sf:liststyle-ref sfa:IDREF="bullets"/></sf:listStyle></sf:property-map></sf:paragraphstyle>
              </sf:styles></sl:stylesheet>
              <sf:text-storage><sf:attachments><sf:attachment sfa:ID="attachment" sf:kind="image-attachment"><sf:media sfa:ID="image">
                <sf:geometry><sf:size sfa:w="72" sfa:h="36"/><sf:position sfa:x="0" sfa:y="0"/></sf:geometry>
                <sf:content><sf:image-media><sf:filtered-image><sf:unfiltered><sf:data sf:path="photo.png"/></sf:unfiltered></sf:filtered-image></sf:image-media></sf:content>
              </sf:media></sf:attachment></sf:attachments><sf:text-body>
                <sf:p sf:style="selected"><sf:span sf:style="emphasis">Emphasis</sf:span><sf:link href="https://example.com/resource">Link</sf:link></sf:p>
                <sf:p sf:style="listed" sf:list-level="1">Bullet</sf:p>
                <sf:p>Before<sf:attachment-ref sfa:IDREF="attachment"/>After</sf:p>
              </sf:text-body></sf:text-storage>
            </sl:document>
            """;
        var package = new MemoryStream();
        using (var zip = new ZipArchive(package, ZipArchiveMode.Create, leaveOpen: true)) {
            using (var writer = new StreamWriter(zip.CreateEntry("index.xml").Open())) writer.Write(xml);
            using Stream image = zip.CreateEntry("photo.png").Open();
            image.Write(Image, 0, Image.Length);
        }
        package.Position = 0;
        return package;
    }
}
