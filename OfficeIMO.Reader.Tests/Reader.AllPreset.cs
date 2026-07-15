using OfficeIMO.OpenDocument;
using OfficeIMO.Reader.All;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderAllPresetTests {
    private static readonly string[] ExpectedModularHandlerIds = {
        "officeimo.reader.asciidoc",
        "officeimo.reader.csv",
        "officeimo.reader.epub",
        "officeimo.reader.html",
        "officeimo.reader.json",
        "officeimo.reader.latex",
        "officeimo.reader.opendocument",
        "officeimo.reader.pdf",
        "officeimo.reader.rtf",
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
        Assert.DoesNotContain(modular, capability => capability.Id.Contains("ocr", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(modular, capability => capability.Id.Contains("provider", StringComparison.OrdinalIgnoreCase));
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
                Bytes: text.ToBytes(),
                Marker: "ODT semantic marker",
                MediaType: "application/vnd.oasis.opendocument.text"),
            (
                Name: "document.ods",
                Bytes: spreadsheet.ToBytes(),
                Marker: "ODS semantic marker",
                MediaType: "application/vnd.oasis.opendocument.spreadsheet"),
            (
                Name: "document.odp",
                Bytes: presentation.ToBytes(),
                Marker: "ODP semantic marker",
                MediaType: "application/vnd.oasis.opendocument.presentation")
        };
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddAllOfficeIMOHandlers()
            .Build();

        foreach ((string name, byte[] bytes, string marker, string mediaType) in cases) {
            ReaderDetectionResult detection = reader.Detect(bytes, name);

            Assert.Equal(ReaderInputKind.OpenDocument, detection.ExtensionKind);
            Assert.Equal(ReaderInputKind.OpenDocument, detection.ContentKind);
            Assert.Equal(ReaderInputKind.OpenDocument, detection.Kind);
            Assert.Equal(mediaType, detection.MediaType);
            Assert.Equal(ReaderDetectionConfidence.High, detection.Confidence);

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
}
