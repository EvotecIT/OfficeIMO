using System.Xml.Linq;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkLegacyLossTests {
    private const string PlaceholderLoss = "IWORK_LEGACY_KEYNOTE_PLACEHOLDER_UNRESOLVED";

    [Theory]
    [InlineData("title", "missing-target")]
    [InlineData("body", "missing-target")]
    [InlineData("title", "missing-target-id")]
    [InlineData("body", "missing-target-id")]
    [InlineData("title", "missing-reference-id")]
    [InlineData("body", "missing-reference-id")]
    [InlineData("title", "mismatched-id")]
    [InlineData("body", "mismatched-id")]
    public void Keynote09UnresolvedPlaceholderReferencesReportLossWithoutDroppingRecoveredPeers(
        string placeholder, string failure) {
        using MemoryStream package = KeynotePackage(placeholder, failure);
        using IWorkSourceDocument source = IWorkSourceDocument.Load(package);
        IWorkKeynoteProjection projection = source.ReadKeynote();
        IWorkKeynoteSlide slide = Assert.Single(projection.Slides);
        Assert.Single(projection.Diagnostics, diagnostic => diagnostic.Code == PlaceholderLoss
            && diagnostic.LossKind == OfficeConversionLossKind.Omission && diagnostic.EntryPath == "index.apxl");
        Assert.Single(slide.Drawables);
        string retained = placeholder == "title" ? "Recovered body" : "Recovered title";
        var converted = Convert(source, "keynote");
        Assert.Single(converted.Report.Diagnostics, diagnostic => diagnostic.Code == PlaceholderLoss
            && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        Assert.Equal(new[] { retained }, TextValues(converted.Document, "keynote"));
    }

    [Fact]
    public void Keynote09RepeatedRecoveredPlaceholdersRemainSingleEditableObjectsWithoutReferenceLoss() {
        using MemoryStream package = KeynotePackage(null, null);
        using IWorkSourceDocument source = IWorkSourceDocument.Load(package);
        IWorkKeynoteProjection projection = source.ReadKeynote();
        IWorkKeynoteSlide slide = Assert.Single(projection.Slides);
        Assert.Equal(2, slide.Drawables.Count);
        Assert.Equal("Recovered title", slide.Title);
        Assert.Single(slide.TextBoxes);
        Assert.DoesNotContain(projection.Diagnostics, diagnostic => diagnostic.Code == PlaceholderLoss);
        var converted = Convert(source, "keynote");
        Assert.DoesNotContain(converted.Report.Diagnostics, diagnostic => diagnostic.Code == PlaceholderLoss);
        Assert.Equal(new[] { "Recovered title", "Recovered body" }, TextValues(converted.Document, "keynote"));
    }

    [Theory]
    [InlineData("pages", "shear")]
    [InlineData("numbers", "shear")]
    [InlineData("keynote", "shear")]
    [InlineData("pages", "flip")]
    [InlineData("numbers", "flip")]
    [InlineData("keynote", "flip")]
    [InlineData("pages", "invalid-angle")]
    [InlineData("numbers", "invalid-angle")]
    [InlineData("keynote", "invalid-angle")]
    public void Legacy09TableTransformOmissionsReachEachPublicConversionReport(string application, string transform) {
        string attribute = transform switch {
            "shear" => "sf:shearXAngle=\"0.5\"",
            "flip" => "sf:horizontalFlip=\"true\"",
            _ => "sf:angle=\"invalid\""
        };
        using MemoryStream package = TablePackage(application, Table(attribute, "<sf:n sf:ct=\"0\" sf:v=\"42\"/>"));
        using IWorkSourceDocument source = IWorkSourceDocument.Load(package);
        var projected = ReadTable(source, application);
        Assert.True(projected.Table.CellStorageIsComplete);
        Assert.Null(projected.Table.Geometry);
        Assert.Equal(42d, projected.Table.GetCell(1, 1)!.Value);
        Assert.Single(projected.Diagnostics, diagnostic => diagnostic.Code == "IWORK_LEGACY_GEOMETRY_UNSUPPORTED"
            && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        var converted = Convert(source, application);
        Assert.Single(converted.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_LEGACY_GEOMETRY_UNSUPPORTED"
            && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        Assert.Contains("42", TextValues(converted.Document, application));
    }

    [Theory]
    [InlineData("pages", "unknown-record")]
    [InlineData("numbers", "unknown-record")]
    [InlineData("keynote", "unknown-record")]
    [InlineData("pages", "invalid-value")]
    [InlineData("numbers", "invalid-value")]
    [InlineData("keynote", "invalid-value")]
    [InlineData("pages", "unsupported-value")]
    [InlineData("numbers", "unsupported-value")]
    [InlineData("keynote", "unsupported-value")]
    public void Legacy09IncompleteTablesRetainRecoveredCellsAndReportTheirLoss(string application, string failure) {
        string rejected = failure switch {
            "unknown-record" => "<sf:unknown-cell/>",
            "invalid-value" => "<sf:n sf:v=\"invalid\"/>",
            _ => "<sf:grouping/>"
        };
        using MemoryStream package = TablePackage(application,
            Table(string.Empty, "<sf:n sf:ct=\"0\" sf:v=\"42\"/>" + rejected + "<sf:n sf:v=\"99\"/>"));
        using IWorkSourceDocument source = IWorkSourceDocument.Load(package);
        var projected = ReadTable(source, application);
        Assert.False(projected.Table.CellStorageIsComplete);
        Assert.Equal(42d, projected.Table.GetCell(1, 1)!.Value);
        string expectedCode = application == "numbers"
            ? "IWORK_LEGACY_NUMBERS_GRID_PARTIAL" : "IWORK_LEGACY_TABLE_GRID_PARTIAL";
        Assert.Single(projected.Diagnostics, diagnostic => diagnostic.Code == expectedCode
            && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        var converted = Convert(source, application);
        Assert.Single(converted.Report.Diagnostics, diagnostic => diagnostic.Code == expectedCode
            && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        string[] values = TextValues(converted.Document, application);
        Assert.Contains("42", values);
        if (failure == "unknown-record") {
            Assert.Null(projected.Table.GetCell(2, 1));
            Assert.DoesNotContain("99", values);
        } else {
            Assert.Equal(99d, projected.Table.GetCell(2, 1)!.Value);
            Assert.Contains("99", values);
        }
    }

    [Theory]
    [InlineData("pages")]
    [InlineData("numbers")]
    [InlineData("keynote")]
    public void Legacy09NullTablesKeepExistingOmissionDiagnosticsAndRecoveredSiblingTables(string application) {
        using MemoryStream package = TablePackage(application,
            Table(string.Empty, "<sf:n sf:ct=\"0\" sf:v=\"42\"/>") + "<sf:tabular-info/>");
        using IWorkSourceDocument source = IWorkSourceDocument.Load(package);
        var projected = ReadTable(source, application);
        Assert.True(projected.Table.CellStorageIsComplete);
        Assert.Equal(42d, projected.Table.GetCell(1, 1)!.Value);
        string expectedCode = application == "numbers"
            ? "IWORK_LEGACY_NUMBERS_GRID_PARTIAL" : "IWORK_LEGACY_DRAWABLE_OMITTED";
        Assert.Single(projected.Diagnostics, diagnostic => diagnostic.Code == expectedCode);
        var converted = Convert(source, application);
        Assert.Single(converted.Report.Diagnostics, diagnostic => diagnostic.Code == expectedCode
            && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        Assert.Contains("42", TextValues(converted.Document, application));
    }

    private static (IWorkTable Table, IReadOnlyList<IWorkDiagnostic> Diagnostics) ReadTable(
        IWorkSourceDocument source, string application) {
        if (application == "pages") {
            IWorkPagesProjection projection = source.ReadPages();
            return (Assert.Single(projection.Tables), projection.Diagnostics);
        }
        if (application == "numbers") {
            IWorkNumbersProjection projection = source.ReadNumbers();
            return (Assert.Single(Assert.Single(projection.Sheets).Tables), projection.Diagnostics);
        }
        IWorkKeynoteProjection keynote = source.ReadKeynote();
        return (Assert.Single(Assert.Single(keynote.Slides).Tables), keynote.Diagnostics);
    }

    private static (IWorkConversionReport Report, XDocument Document) Convert(IWorkSourceDocument source, string application) {
        var options = new IWorkConversionOptions {
            Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true
        };
        IWorkConversionReport report;
        using var output = new MemoryStream();
        if (application == "pages") {
            using var converted = source.ToWordResult(options);
            report = converted.Report;
            converted.Value.Save(output);
        } else if (application == "numbers") {
            using var converted = source.ToExcelResult(options);
            report = converted.Report;
            converted.Value.Save(output);
        } else {
            using var converted = source.ToPowerPointResult(options);
            report = converted.Report;
            converted.Value.Save(output);
        }
        output.Position = 0;
        using var zip = new ZipArchive(output, ZipArchiveMode.Read);
        string path = application switch {
            "pages" => "word/document.xml", "numbers" => "xl/worksheets/sheet1.xml", _ => "ppt/slides/slide1.xml"
        };
        using Stream xml = zip.GetEntry(path)!.Open();
        return (report, XDocument.Load(xml));
    }

    private static string[] TextValues(XDocument document, string application) {
        XNamespace ns = application switch {
            "pages" => "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "numbers" => "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            _ => "http://schemas.openxmlformats.org/drawingml/2006/main"
        };
        return document.Descendants(ns + (application == "numbers" ? "v" : "t")).Select(element => element.Value).ToArray();
    }

    private static MemoryStream KeynotePackage(string? brokenKind, string? failure) {
        string Definition(string kind) => brokenKind == kind && failure == "missing-target" ? string.Empty
            : $"<key:{kind}-placeholder {(brokenKind == kind && failure == "missing-target-id" ? string.Empty : $"sfa:ID=\"{kind}\"")}><key:text><sf:text-storage><sf:text-body><sf:p>Recovered {kind}</sf:p></sf:text-body></sf:text-storage></key:text></key:{kind}-placeholder>";
        string Reference(string kind) => $"<sf:{kind}-placeholder-ref {(brokenKind == kind && failure == "missing-reference-id" ? string.Empty : $"sfa:IDREF=\"{(brokenKind == kind && failure == "mismatched-id" ? "missing" : kind)}\"")}/>";
        return Package($"""
            <key:presentation xmlns:key="http://developer.apple.com/namespaces/keynote2"
              xmlns:sf="http://developer.apple.com/namespaces/sf" xmlns:sfa="http://developer.apple.com/namespaces/sfa"
              key:version="92008102400"><key:slide-list><key:slide>{Definition("title")}{Definition("body")}
              <key:page><sf:drawables>{Reference("title")}{Reference("title")}{Reference("body")}{Reference("body")}</sf:drawables></key:page>
              </key:slide></key:slide-list></key:presentation>
            """, "index.apxl");
    }

    private static string Table(string transform, string cells) => $"""
        <sf:tabular-info><sf:geometry {transform}><sf:position sfa:x="12" sfa:y="20"/><sf:size sfa:w="160" sfa:h="40"/></sf:geometry>
          <sf:tabular-model sf:name="Recovered cells"><sf:grid>
            <sf:columns sf:count="2"><sf:grid-column sf:width="80"/><sf:grid-column sf:width="80"/></sf:columns>
            <sf:rows sf:count="2"><sf:grid-row sf:height="20"/><sf:grid-row sf:height="20"/></sf:rows>
            <sf:datasource>{cells}</sf:datasource>
          </sf:grid></sf:tabular-model>
        </sf:tabular-info>
        """;

    private static MemoryStream TablePackage(string application, string table) {
        const string ns = "xmlns:sf=\"http://developer.apple.com/namespaces/sf\" xmlns:sfa=\"http://developer.apple.com/namespaces/sfa\"";
        return application switch {
            "pages" => Package($"""
                <sl:document xmlns:sl="http://developer.apple.com/namespaces/sl" {ns} sl:version="92008102400">
                  <sf:text-storage><sf:text-body><sf:p>Body</sf:p></sf:text-body></sf:text-storage>
                  <sl:drawables><sl:page-group sl:page="1">{table}</sl:page-group></sl:drawables>
                </sl:document>
                """, "index.xml"),
            "numbers" => Package($"""
                <ls:document xmlns:ls="http://developer.apple.com/namespaces/ls" {ns} ls:version="92008102400">
                  <ls:workspace-array><ls:workspace ls:workspace-name="Recovered"><ls:page-info><sf:layers><sf:layer><sf:drawables>{table}</sf:drawables></sf:layer></sf:layers></ls:page-info></ls:workspace></ls:workspace-array>
                </ls:document>
                """, "index.xml"),
            _ => Package($"""
                <key:presentation xmlns:key="http://developer.apple.com/namespaces/keynote2" {ns} key:version="92008102400">
                  <key:slide-list><key:slide><key:page><sf:drawables>{table}</sf:drawables></key:page></key:slide></key:slide-list>
                </key:presentation>
                """, "index.apxl")
        };
    }

    private static MemoryStream Package(string xml, string name) {
        var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        using (var writer = new StreamWriter(zip.CreateEntry(name).Open())) writer.Write(xml);
        output.Position = 0;
        return output;
    }
}
