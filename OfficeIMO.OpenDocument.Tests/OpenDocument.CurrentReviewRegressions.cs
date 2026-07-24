using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.OpenDocument.Testing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentCurrentReviewRegressionTests {
    [Fact]
    public void TextAndPresentationTablesIncludeHeaderRowsInSourceOrder() {
        OdtDocument text = OdtDocument.Create();
        OdtTable textTable = text.AddTable(2, 1, "TextTable");
        textTable.Cell(0, 0).Text = "Text header";
        textTable.Cell(1, 0).Text = "Text body";
        WrapFirstRowAsHeader(textTable.Element);

        OdpPresentation presentation = OdpPresentation.Create();
        OdpTable presentationTable = presentation.AddSlide("Table").AddTable(
            OdfRect.FromCentimeters(1, 1, 8, 4), 2, 1, "PresentationTable");
        presentationTable.Cell(0, 0).Text = "Slide header";
        presentationTable.Cell(1, 0).Text = "Slide body";
        WrapFirstRowAsHeader(presentationTable.Element.Element(OdfNamespaces.Table + "table")!);

        Assert.Equal(new[] { "Text header", "Text body" }, textTable.Rows.Select(row => row.Cells[0].Text));
        Assert.Equal(new[] { "Slide header", "Slide body" }, presentationTable.Rows.Select(row => row.Cells[0].Text));
    }

    [Fact]
    public void NewOdsColumnsStayAtTableScopeWhenHeaderRowsComeFirst() {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        sheet.Cell(0, 0).SetString("Header");
        sheet.Cell(1, 0).SetString("Body");
        XElement table = sheet.Element;
        WrapFirstRowAsHeader(table);

        _ = sheet.Column(2);

        XElement[] columns = table.Elements(OdfNamespaces.Table + "table-column").ToArray();
        Assert.Equal(3, columns.Length);
        Assert.All(columns, column => Assert.Same(table, column.Parent));
        Assert.Empty(table.Element(OdfNamespaces.Table + "table-header-rows")!
            .Elements(OdfNamespaces.Table + "table-column"));
        Assert.True(document.Validate().IsValid);
    }

    [Fact]
    public void FlatExportTreatsMissingStylesPartAsEmptyStyles() {
        OdtDocument document = OdtDocument.Create();
        document.AddParagraph("Minimal flat document");
        document.Package.RemoveEntry("styles.xml");
        using var stream = new MemoryStream();

        OdfSaveResult save = document.SaveFlatXml(stream);

        stream.Position = 0;
        OdtDocument reopened = OdtDocument.LoadFlatXml(stream);
        Assert.Equal("Minimal flat document", reopened.Paragraphs.Single().Text);
        Assert.DoesNotContain("styles.xml", save.Report.RewrittenEntries);
    }

    [Fact]
    public void FlatImageExtractionPreservesSupportedMimeType() {
        byte[] webp = { 0x52, 0x49, 0x46, 0x46, 0x04, 0x00, 0x00, 0x00, 0x57, 0x45, 0x42, 0x50 };
        OdtDocument document = OdtDocument.Create();
        document.AddParagraph().AddImage(webp, "pixel.webp",
            OdfLength.Centimeters(1), OdfLength.Centimeters(1));
        using var stream = new MemoryStream();

        document.SaveFlatXml(stream);
        stream.Position = 0;
        OdtDocument reopened = OdtDocument.LoadFlatXml(stream);

        OdtImage image = reopened.Paragraphs.Single().Images.Single();
        Assert.EndsWith(".webp", image.Path, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(webp, image.GetImageBytes());
    }

    [Fact]
    public void FlatSvgBinaryDataIsParsedAndReserializedBeforePackaging() {
        const string safeSvg = "<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'><rect width='10' height='10' fill='red'/></svg>";
        const string activeSvg = "<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'><script>alert(1)</script><rect width='10' height='10' fill='red' onclick='alert(2)'/></svg>";
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph().AddImage(
            Encoding.UTF8.GetBytes(safeSvg),
            "shape.svg",
            OdfLength.Centimeters(1),
            OdfLength.Centimeters(1));
        XDocument flat = source.ToFlatXml();
        XElement imageElement = Assert.Single(flat.Descendants(OdfNamespaces.Draw + "image"));
        imageElement.SetAttributeValue(OdfNamespaces.Draw + "mime-type", "image/svg+xml");
        imageElement.Element(OdfNamespaces.Office + "binary-data")!.Value =
            Convert.ToBase64String(Encoding.UTF8.GetBytes(activeSvg));
        using var stream = new MemoryStream();
        flat.Save(stream);
        stream.Position = 0;

        OdtDocument reopened = OdtDocument.LoadFlatXml(stream);
        string packagedSvg = Encoding.UTF8.GetString(
            reopened.Paragraphs.Single().Images.Single().GetImageBytes());

        Assert.Contains("<rect", packagedSvg, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("<script", packagedSvg, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("onclick", packagedSvg, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void PackageVersionFallsBackToContentPartWhenManifestVersionIsMissing() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph("ODF 1.3");
        byte[] package = RewriteWithoutManifestVersion(source.ToBytes(), "1.3");

        OdtDocument reopened = OdtDocument.Load(new MemoryStream(package));

        Assert.Equal(OdfVersion.V1_3, reopened.Version);
        Assert.DoesNotContain(reopened.Diagnostics, diagnostic => diagnostic.Id == "ODF003");
    }

    [Fact]
    public void GraphicColorsAreHiddenWhenFillOrStrokeModeIsNone() {
        OdpPresentation presentation = OdpPresentation.Create();
        OdpRectangle rectangle = presentation.AddSlide("Shape").AddRectangle(
            OdfRect.FromCentimeters(1, 1, 4, 2));
        rectangle.FillColor = OdfColor.Parse("#112233");
        rectangle.StrokeColor = OdfColor.Parse("#445566");
        string styleName = (string)rectangle.Element.Attribute(OdfNamespaces.Draw + "style-name")!;
        XElement properties = presentation.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Style + "style")
            .Single(style => (string?)style.Attribute(OdfNamespaces.Style + "name") == styleName)
            .Element(OdfNamespaces.Style + "graphic-properties")!;
        properties.SetAttributeValue(OdfNamespaces.Draw + "fill", "none");
        properties.SetAttributeValue(OdfNamespaces.Draw + "stroke", "none");

        Assert.Null(rectangle.FillColor);
        Assert.Null(rectangle.StrokeColor);
    }

    [Fact]
    public void PresentationBackgroundsHonorFillModeAndReferencedMasterPageLayout() {
        OdpPresentation presentation = OdpPresentation.Create();
        OdpMasterPage unused = presentation.AddMasterPage("Unused");
        OdpMasterPage selected = presentation.AddMasterPage("Selected");
        OdpSlide slide = presentation.AddSlide("Slide");
        slide.MasterPageName = selected.Name;
        slide.BackgroundColor = OdfColor.Parse("#112233");

        XElement styles = presentation.Package.GetXml("styles.xml").Root!;
        XElement automatic = styles.Element(OdfNamespaces.Office + "automatic-styles")!;
        XElement[] layouts = automatic.Elements(OdfNamespaces.Style + "page-layout").ToArray();
        XElement selectedLayout = new XElement(layouts[0]);
        selectedLayout.SetAttributeValue(OdfNamespaces.Style + "name", "SelectedLayout");
        selectedLayout.Element(OdfNamespaces.Style + "page-layout-properties")!
            .SetAttributeValue(OdfNamespaces.Fo + "page-width", "40cm");
        automatic.AddFirst(selectedLayout);
        styles.Descendants(OdfNamespaces.Style + "master-page")
            .Single(element => (string?)element.Attribute(OdfNamespaces.Style + "name") == selected.Name)
            .SetAttributeValue(OdfNamespaces.Style + "page-layout-name", "SelectedLayout");
        XElement slideProperties = presentation.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Style + "drawing-page-properties").Single();
        slideProperties.SetAttributeValue(OdfNamespaces.Draw + "fill", "none");

        Assert.Null(slide.BackgroundColor);
        Assert.Equal(40D, presentation.PageWidth.ToCentimeters(), 6);
    }

    [Fact]
    public void SpreadsheetDataStylesTolerateMissingStylesPart() {
        OdsDocument document = OdsDocument.Create();
        document.AddNumberStyle("Amount", 2);
        document.Package.RemoveEntry("styles.xml");

        Assert.Contains(document.DataStyles, style => style.Name == "Amount");
    }

    [Fact]
    public void SpreadsheetMergeValidationIncludesHeaderRows() {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        sheet.Cell(0, 0).SetString("Anchor");
        sheet.Cell(0, 1).SetString("Not covered");
        XElement row = sheet.Element.Elements(OdfNamespaces.Table + "table-row").Single();
        row.Elements(OdfNamespaces.Table + "table-cell").First()
            .SetAttributeValue(OdfNamespaces.Table + "number-columns-spanned", 2);
        row.Remove();
        sheet.Element.Add(new XElement(OdfNamespaces.Table + "table-header-rows", row));

        Assert.Contains(document.Validate().Diagnostics, diagnostic => diagnostic.Id == "ODS104");
    }

    [Fact]
    public void DrawableNamesDoNotSatisfyMissingStyleReferences() {
        OdtDocument document = OdtDocument.Create();
        document.AddParagraph("Styled");
        XDocument content = document.Package.GetXml("content.xml");
        content.Descendants(OdfNamespaces.Text + "p").Single()
            .SetAttributeValue(OdfNamespaces.Text + "style-name", "Missing");
        content.Descendants(OdfNamespaces.Office + "text").Single().Add(
            new XElement(OdfNamespaces.Draw + "frame", new XAttribute(OdfNamespaces.Draw + "name", "Missing")));

        Assert.Contains(document.Validate().Diagnostics, diagnostic => diagnostic.Id == "ODF200" &&
            diagnostic.Message.IndexOf("Missing", StringComparison.Ordinal) >= 0);
    }

    private static void WrapFirstRowAsHeader(XElement table) {
        XElement firstRow = table.Elements(OdfNamespaces.Table + "table-row").First();
        XElement secondRow = firstRow.ElementsAfterSelf(OdfNamespaces.Table + "table-row").First();
        firstRow.Remove();
        secondRow.AddBeforeSelf(new XElement(OdfNamespaces.Table + "table-header-rows", firstRow));
    }

    private static byte[] RewriteWithoutManifestVersion(byte[] sourceBytes, string partVersion) {
        return OdfTestPackageRewriter.Rewrite(sourceBytes, (name, bytes) => {
            if (name == "META-INF/manifest.xml") {
                XDocument manifest = XDocument.Parse(Encoding.UTF8.GetString(bytes));
                manifest.Root!.Attribute(OdfNamespaces.Manifest + "version")?.Remove();
                manifest.Root.Elements(OdfNamespaces.Manifest + "file-entry")
                    .First(element => (string?)element.Attribute(OdfNamespaces.Manifest + "full-path") == "/")
                    .Attribute(OdfNamespaces.Manifest + "version")?.Remove();
                bytes = Encoding.UTF8.GetBytes(manifest.ToString(SaveOptions.DisableFormatting));
            } else if (name == "content.xml") {
                XDocument content = XDocument.Parse(Encoding.UTF8.GetString(bytes));
                content.Root!.SetAttributeValue(OdfNamespaces.Office + "version", partVersion);
                bytes = Encoding.UTF8.GetBytes(content.ToString(SaveOptions.DisableFormatting));
            }
            return bytes;
        });
    }
}
