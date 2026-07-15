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
}