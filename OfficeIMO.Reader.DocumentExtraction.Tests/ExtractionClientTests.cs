using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DocumentExtraction;
using Microsoft.Extensions.Documents;
using OfficeIMO.Reader.All;
using Xunit;
using static OfficeIMO.Reader.DocumentExtraction.Tests.ExtractionMappingTests;

namespace OfficeIMO.Reader.DocumentExtraction.Tests;

public sealed class ExtractionClientTests {
    private static readonly byte[] Html = System.Text.Encoding.UTF8.GetBytes("<html><body><h1>Notes</h1><p>Body text.</p></body></html>");

    [Fact]
    public async Task UnsupportedExplicitOptionsAreRejected() {
        using var client = new OfficeIMODocumentExtractionClient(builder => builder.AddAllOfficeIMOHandlers());
        using var stream = new MemoryStream(Html);

        await Assert.ThrowsAsync<NotSupportedException>(() =>
            client.ExtractAsync(stream, "notes", "text/html", new DocumentExtractionOptions { ExtractorName = "layout" }));
        var unknown = new DocumentExtractionOptions();
        unknown.AdditionalProperties["temperature"] = 0.2;
        await Assert.ThrowsAsync<NotSupportedException>(() => client.ExtractAsync(stream, "notes", "text/html", unknown));
    }

    [Fact]
    public async Task PerCallReaderOptionsAreApplied() {
        using var client = new OfficeIMODocumentExtractionClient(builder => builder.AddAllOfficeIMOHandlers());
        var options = new DocumentExtractionOptions();
        options.AdditionalProperties[OfficeIMOExtractionPropertyNames.ReaderOptionsOverride] = new ReaderOptions { ComputeHashes = false };
        using var stream = new MemoryStream(Html);

        ExtractionDocument document = await client.ExtractAsync(stream, "notes", "text/html", options);

        Assert.False(document.Metadata.ContainsKey(OfficeIMOExtractionPropertyNames.SourceHash));
        Assert.Equal("Notes", Text(document.Sections.Single().Elements[0]));
    }

    [Fact]
    public async Task ServiceRegistrationComposesWithMiddleware() {
        var services = new ServiceCollection();
        services.AddOfficeIMODocumentExtraction(builder => builder.AddAllOfficeIMOHandlers())
            .Use(inner => new CountingClient(inner));
        using ServiceProvider provider = services.BuildServiceProvider();

        DocumentExtractionClient client = provider.GetRequiredService<DocumentExtractionClient>();
        using var stream = new MemoryStream(Html);
        ExtractionDocument document = await client.ExtractAsync(stream, "notes", "text/html");

        Assert.Equal(1, Assert.IsType<CountingClient>(client).Calls);
        Assert.IsType<OfficeIMODocumentExtractionClient>(client.GetService(typeof(OfficeIMODocumentExtractionClient)));
        Assert.Equal("Notes", Text(document.Sections.Single().Elements[0]));
    }

    [Fact]
    public void KeyedRegistrationUsesTheReaderFromTheContainerWhenNoneIsConfigured() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddAllOfficeIMOHandlers().Build();
        var services = new ServiceCollection();
        services.AddSingleton(reader);
        services.AddKeyedOfficeIMODocumentExtraction("office");
        using ServiceProvider provider = services.BuildServiceProvider();

        DocumentExtractionClient client = provider.GetRequiredKeyedService<DocumentExtractionClient>("office");

        Assert.Same(reader, client.GetService(typeof(OfficeDocumentReader)));
        Assert.Null(provider.GetService<DocumentExtractionClient>());
    }

    private sealed class CountingClient : DelegatingDocumentExtractionClient {
        public CountingClient(DocumentExtractionClient inner) : base(inner) { }

        public int Calls { get; private set; }

        public override Task<ExtractionDocument> ExtractAsync(Stream document, string identifier, string mediaType,
            DocumentExtractionOptions? options = null, CancellationToken cancellationToken = default) {
            Calls++;
            return base.ExtractAsync(document, identifier, mediaType, options, cancellationToken);
        }
    }
}
