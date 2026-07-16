using OfficeIMO.OpenDocument;
using OfficeIMO.OneNote;
using OfficeIMO.Reader.All;
using OfficeIMO.Reader.Zip;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderAllPresetTests {
    private static readonly string[] ExpectedModularHandlerIds = {
        "officeimo.reader.asciidoc",
        "officeimo.reader.csv",
        "officeimo.reader.email-address-book",
        "officeimo.reader.emailstore",
        "officeimo.reader.epub",
        "officeimo.reader.html",
        "officeimo.reader.image",
        "officeimo.reader.json",
        "officeimo.reader.latex",
        "officeimo.reader.notebook",
        "officeimo.reader.onenote",
        "officeimo.reader.opendocument",
        "officeimo.reader.pdf",
        "officeimo.reader.rtf",
        "officeimo.reader.subtitles",
        "officeimo.reader.visio",
        "officeimo.reader.xml",
        "officeimo.reader.yaml",
        "officeimo.reader.zip"
    };

    [Fact]
    public void PresetRegistersEveryLocalAdapterAndExcludesProviders() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddAllOfficeIMOHandlers()
            .Build();

        ReaderHandlerCapability[] modular = reader.GetCapabilities()
            .Where(capability => !capability.IsBuiltIn)
            .ToArray();

        Assert.Equal(ExpectedModularHandlerIds, modular.Select(capability => capability.Id).ToArray());
        ReaderHandlerCapability oneNote = Assert.Single(
            modular,
            capability => capability.Id == "officeimo.reader.onenote");
        Assert.Equal(
            new[] { ".one", ".onepkg", ".onetoc2" },
            oneNote.Extensions.OrderBy(extension => extension, StringComparer.Ordinal).ToArray());
        Assert.DoesNotContain(modular, capability => capability.Id.Contains("ocr", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(modular, capability => capability.Id.Contains("provider", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void PresetReadsOneNoteSectionNotebookAndPackage() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddAllOfficeIMOHandlers()
            .Build();

        OfficeDocumentReadResult sectionResult = reader.ReadDocument(
            Path.Combine(AppContext.BaseDirectory, "OneNoteFixtures", "testOneNote2016.one"));
        AssertOneNotePresetResult(sectionResult);

        OneNoteNotebook notebook = CreateOneNoteNotebook();
        using (var package = new MemoryStream(OneNotePackageWriter.Write(notebook))) {
            OfficeDocumentReadResult packageResult = reader.ReadDocument(package, "preset.onepkg");
            AssertOneNotePresetResult(packageResult);
        }

        string root = Path.Combine(Path.GetTempPath(), "OfficeIMO-Reader-All-OneNote-" + Guid.NewGuid().ToString("N"));
        try {
            OneNoteNotebookWriter.Write(notebook, root);
            OfficeDocumentReadResult notebookResult = reader.ReadDocument(
                Path.Combine(root, "Open Notebook.onetoc2"));
            AssertOneNotePresetResult(notebookResult);
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, true);
        }
    }

    [Fact]
    public void PresetConfigurationRemainsInstanceScoped() {
        OfficeDocumentReader configured = new OfficeDocumentReaderBuilder()
            .AddAllOfficeIMOHandlers(new ReaderAllOptions {
                Csv = new Csv.CsvReadOptions { ChunkRows = 1 }
            })
            .Build();
        OfficeDocumentReader unconfigured = new OfficeDocumentReaderBuilder().Build();

        Assert.Contains(configured.GetCapabilities(), capability => capability.Id == "officeimo.reader.csv" && !capability.IsBuiltIn);
        Assert.DoesNotContain(unconfigured.GetCapabilities(), capability => capability.Id == "officeimo.reader.csv" && !capability.IsBuiltIn);

        OfficeDocumentReadResult document = configured.ReadDocument(
            Encoding.UTF8.GetBytes("name,value\nalpha,1\nbeta,2"),
            "data.csv");
        Assert.Equal(2, document.Chunks.Count);
    }
    [Fact]
    public void PresetPreservesOpenDocumentRoutingWhenContentIsPreferred() {
        OdtDocument text = OdtDocument.Create();
        text.AddParagraph("ODT semantic marker");
        OdsDocument spreadsheet = OdsDocument.Create();
        spreadsheet.AddSheet("Data").Cell(0, 0).SetString("ODS semantic marker");
        OdpPresentation presentation = OdpPresentation.Create();
        presentation.AddSlide("Summary")
            .AddTextBox(OdfRect.FromCentimeters(1, 1, 12, 3), "ODP semantic marker");
        var cases = new[] {
            (
                Name: "document.odt",
                CrossLabeledName: "document.ods",
                Bytes: text.ToBytes(),
                Marker: "ODT semantic marker",
                MediaType: "application/vnd.oasis.opendocument.text"),
            (
                Name: "document.ods",
                CrossLabeledName: "document.odp",
                Bytes: spreadsheet.ToBytes(),
                Marker: "ODS semantic marker",
                MediaType: "application/vnd.oasis.opendocument.spreadsheet"),
            (
                Name: "document.odp",
                CrossLabeledName: "document.odt",
                Bytes: presentation.ToBytes(),
                Marker: "ODP semantic marker",
                MediaType: "application/vnd.oasis.opendocument.presentation")
        };
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddAllOfficeIMOHandlers()
            .Build();

        foreach ((string name, string crossLabeledName, byte[] bytes, string marker, string mediaType) in cases) {
            ReaderDetectionResult detection = reader.Detect(bytes, name);

            Assert.Equal(ReaderInputKind.OpenDocument, detection.ExtensionKind);
            Assert.Equal(ReaderInputKind.OpenDocument, detection.ContentKind);
            Assert.Equal(ReaderInputKind.OpenDocument, detection.Kind);
            Assert.Equal(mediaType, detection.MediaType);
            Assert.Equal(ReaderDetectionConfidence.High, detection.Confidence);

            ReaderDetectionResult crossLabeled = reader.Detect(
                bytes,
                crossLabeledName,
                new ReaderDetectionOptions { Mode = ReaderDetectionMode.PreferContent });
            Assert.Equal(ReaderInputKind.OpenDocument, crossLabeled.ExtensionKind);
            Assert.Equal(ReaderInputKind.OpenDocument, crossLabeled.ContentKind);
            Assert.Equal(ReaderInputKind.OpenDocument, crossLabeled.Kind);
            Assert.Equal(mediaType, crossLabeled.MediaType);
            Assert.Contains("container:opendocument-mimetype", crossLabeled.Evidence);

            OfficeDocumentReadResult result = reader.ReadDocument(
                bytes,
                name,
                new ReaderOptions { DetectionMode = ReaderDetectionMode.PreferContent });

            Assert.Equal(ReaderInputKind.OpenDocument, result.Kind);
            Assert.Contains("officeimo.reader.opendocument", result.CapabilitiesUsed);
            IEnumerable<string> extractedValues = result.Chunks.Select(chunk => chunk.Text)
                .Concat(result.Chunks
                    .SelectMany(chunk => chunk.Tables ?? Array.Empty<ReaderTable>())
                    .SelectMany(table => table.Columns.Concat(table.Rows.SelectMany(row => row))));
            Assert.Contains(marker, string.Join("\n", extractedValues), StringComparison.Ordinal);
        }
    }

    [Theory]
    [InlineData("document.odt", ReaderDetectionMode.ContentWhenUnknown)]
    [InlineData("document.blob", ReaderDetectionMode.ContentWhenUnknown)]
    [InlineData("document.zip", ReaderDetectionMode.PreferContent)]
    public void ZipOnlyRegistrationFallsBackForOpenDocumentStreams(
        string sourceName,
        ReaderDetectionMode detectionMode) {
        OdtDocument document = OdtDocument.Create();
        document.AddParagraph("ZIP fallback marker");
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddZipHandler()
            .Build();

        OfficeDocumentReadResult result = reader.ReadDocument(
            document.ToBytes(),
            sourceName,
            new ReaderOptions { DetectionMode = detectionMode });

        Assert.NotEmpty(result.Chunks);
        Assert.Contains(result.Chunks, chunk =>
            chunk.Location.Path?.Contains("::", StringComparison.Ordinal) == true);
    }

    [Fact]
    public void ZipOnlyRegistrationFallsBackForOpenDocumentPaths() {
        string path = Path.Combine(Path.GetTempPath(), $"officeimo-zip-fallback-{Guid.NewGuid():N}.odt");
        try {
            OdtDocument document = OdtDocument.Create();
            document.AddParagraph("ZIP path fallback marker");
            File.WriteAllBytes(path, document.ToBytes());
            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
                .AddZipHandler()
                .Build();

            OfficeDocumentReadResult result = reader.ReadDocument(path);

            Assert.NotEmpty(result.Chunks);
            Assert.Contains(result.Chunks, chunk =>
                chunk.Location.Path?.Contains("::", StringComparison.Ordinal) == true);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    private static OneNoteNotebook CreateOneNoteNotebook() {
        var notebook = new OneNoteNotebook { Name = "Reader All" };
        var section = new OneNoteSection { Name = "Preset" };
        var page = new OneNotePage { Title = "OneNote preset" };
        var paragraph = new OneNoteParagraph();
        paragraph.Runs.Add(new OneNoteTextRun { Text = "Reader All OneNote content" });
        page.DirectContent.Add(paragraph);
        section.Pages.Add(page);
        notebook.Sections.Add(section);
        return notebook;
    }

    private static void AssertOneNotePresetResult(OfficeDocumentReadResult result) {
        Assert.Equal(ReaderInputKind.OneNote, result.Kind);
        Assert.Contains("officeimo.reader.onenote", result.CapabilitiesUsed);
        Assert.NotEmpty(result.Chunks);
    }
}
