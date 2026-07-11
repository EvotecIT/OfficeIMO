using OfficeIMO.Reader;
using OfficeIMO.Reader.Csv;
using OfficeIMO.Reader.Epub;
using OfficeIMO.Reader.Html;
using OfficeIMO.Reader.Json;
using OfficeIMO.Reader.OpenDocument;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Reader.Rtf;
using OfficeIMO.Reader.Text;
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
            DocumentReaderCsvRegistrationExtensions.HandlerId);
        AssertHandler(
            new OfficeDocumentReaderBuilder().AddEpubHandler().Build(),
            DocumentReaderEpubRegistrationExtensions.HandlerId);
        AssertHandler(
            new OfficeDocumentReaderBuilder().AddHtmlHandler().Build(),
            DocumentReaderHtmlRegistrationExtensions.HandlerId);
        AssertHandler(
            new OfficeDocumentReaderBuilder().AddJsonHandler().Build(),
            DocumentReaderJsonRegistrationExtensions.HandlerId);
        AssertHandler(
            new OfficeDocumentReaderBuilder().AddOpenDocumentHandler().Build(),
            DocumentReaderOpenDocumentRegistrationExtensions.HandlerId);
        AssertHandler(
            new OfficeDocumentReaderBuilder().AddPdfHandler().Build(),
            DocumentReaderPdfRegistrationExtensions.HandlerId);
        AssertHandler(
            new OfficeDocumentReaderBuilder().AddRtfHandler().Build(),
            DocumentReaderRtfRegistrationExtensions.HandlerId);
        AssertHandler(
            new OfficeDocumentReaderBuilder().AddStructuredTextHandler().Build(),
            DocumentReaderTextRegistrationExtensions.HandlerId);
        AssertHandler(
            new OfficeDocumentReaderBuilder().AddVisioHandler().Build(),
            DocumentReaderVisioRegistrationExtensions.HandlerId);
        AssertHandler(
            new OfficeDocumentReaderBuilder().AddXmlHandler().Build(),
            DocumentReaderXmlRegistrationExtensions.HandlerId);
        AssertHandler(
            new OfficeDocumentReaderBuilder().AddYamlHandler().Build(),
            DocumentReaderYamlRegistrationExtensions.HandlerId);
        AssertHandler(
            new OfficeDocumentReaderBuilder().AddZipHandler().Build(),
            DocumentReaderZipRegistrationExtensions.HandlerId);
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

    private static void AssertHandler(OfficeDocumentReader reader, string expectedHandlerId) {
        ReaderHandlerCapability capability = Assert.Single(reader.GetCapabilities(includeBuiltIn: false, includeCustom: true));
        Assert.Equal(expectedHandlerId, capability.Id);
    }
}
