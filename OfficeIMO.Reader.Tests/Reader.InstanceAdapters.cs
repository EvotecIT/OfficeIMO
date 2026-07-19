using OfficeIMO.Reader;
using OfficeIMO.Reader.Csv;
using OfficeIMO.Reader.Epub;
using OfficeIMO.Reader.Html;
using OfficeIMO.Reader.Json;
using OfficeIMO.Reader.OpenDocument;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Reader.Rtf;
using OfficeIMO.Reader.Visio;
using OfficeIMO.Reader.Xml;
using OfficeIMO.Reader.Yaml;
using OfficeIMO.Reader.Zip;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderInstanceAdapterTests {
    [Fact]
    public void ModularAdapters_CanConfigureIsolatedReaders() {
        AssertHandler(
            new OfficeDocumentReaderBuilder().AddCsvHandler().Build(),
            OfficeDocumentReaderBuilderCsvExtensions.HandlerId);
        AssertHandler(
            new OfficeDocumentReaderBuilder().AddEpubHandler().Build(),
            OfficeDocumentReaderBuilderEpubExtensions.HandlerId);
        AssertHandler(
            new OfficeDocumentReaderBuilder().AddHtmlHandler().Build(),
            OfficeDocumentReaderBuilderHtmlExtensions.HandlerId);
        AssertHandler(
            new OfficeDocumentReaderBuilder().AddJsonHandler().Build(),
            OfficeDocumentReaderBuilderJsonExtensions.HandlerId);
        AssertHandler(
            new OfficeDocumentReaderBuilder().AddOpenDocumentHandler().Build(),
            OfficeDocumentReaderBuilderOpenDocumentExtensions.HandlerId);
        AssertHandler(
            new OfficeDocumentReaderBuilder().AddPdfHandler().Build(),
            OfficeDocumentReaderBuilderPdfExtensions.HandlerId);
        AssertHandler(
            new OfficeDocumentReaderBuilder().AddRtfHandler().Build(),
            OfficeDocumentReaderBuilderRtfExtensions.HandlerId);
        AssertHandler(
            new OfficeDocumentReaderBuilder().AddVisioHandler().Build(),
            OfficeDocumentReaderBuilderVisioExtensions.HandlerId);
        AssertHandler(
            new OfficeDocumentReaderBuilder().AddXmlHandler().Build(),
            OfficeDocumentReaderBuilderXmlExtensions.HandlerId);
        AssertHandler(
            new OfficeDocumentReaderBuilder().AddYamlHandler().Build(),
            OfficeDocumentReaderBuilderYamlExtensions.HandlerId);
        AssertHandler(
            new OfficeDocumentReaderBuilder().AddZipHandler().Build(),
            OfficeDocumentReaderBuilderZipExtensions.HandlerId);
    }

    [Fact]
    public void JsonAdapter_RoutesThroughInstanceWithoutStaticRegistration() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddJsonHandler(new JsonReadOptions { IncludeMarkdown = false })
            .Build();
        using var stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes("{\"name\":\"OfficeIMO\"}"));

        ReaderChunk[] chunks = reader.Read(stream, "sample.json").ToArray();

        Assert.NotEmpty(chunks);
        Assert.Contains(chunks, chunk => chunk.Text?.Contains("OfficeIMO", StringComparison.Ordinal) == true);
    }

    [Fact]
    public void InstanceConvenienceApis_UseConfiguredHandler() {
        const string extension = ".surfaceix";
        var table = new ReaderTable { Columns = new[] { "Name" }, Rows = new[] { new[] { "OfficeIMO" } } };
        var visual = new ReaderVisual { Kind = "diagram", Content = "graph TD" };
        var asset = new OfficeDocumentAsset { Id = "asset-1", Kind = "image", FileName = "image.png" };
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddHandler(new ReaderHandlerRegistration {
                Id = "officeimo.tests.instance.surfaces",
                Kind = ReaderInputKind.Text,
                Extensions = new[] { extension },
                ReadStream = (stream, sourceName, options, cancellationToken) => new[] {
                    new ReaderChunk {
                        Id = "surface-1",
                        Kind = ReaderInputKind.Text,
                        Tables = new[] { table },
                        Visuals = new[] { visual }
                    }
                },
                ReadDocumentStream = (stream, sourceName, options, cancellationToken) => new OfficeDocumentReadResult {
                    Kind = ReaderInputKind.Text,
                    Assets = new[] { asset }
                }
            })
            .Build();

        using var tableStream = new MemoryStream(new byte[] { 1 });
        using var visualStream = new MemoryStream(new byte[] { 1 });
        using var assetStream = new MemoryStream(new byte[] { 1 });

        Assert.Equal("OfficeIMO", Assert.Single(reader.ReadTables(tableStream, "sample" + extension)).Rows[0][0]);
        Assert.Equal("graph TD", Assert.Single(reader.ReadVisuals(visualStream, "sample" + extension)).Content);
        Assert.Same(asset, Assert.Single(reader.ReadAssets(assetStream, "sample" + extension)));
    }

    private static void AssertHandler(OfficeDocumentReader reader, string expectedHandlerId) {
        ReaderHandlerCapability capability = Assert.Single(reader.GetCapabilities(), item => item.Origin == ReaderHandlerOrigin.OfficeIMO);
        Assert.Equal(expectedHandlerId, capability.Id);
    }
}
