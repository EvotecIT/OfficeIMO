using OfficeIMO.Reader;
using OfficeIMO.Reader.Visio;
using OfficeIMO.Visio;
using System.Text;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderVisioModularTests {
    [Fact]
    public void DocumentReaderVisio_ReadVisio_EmitsPageChunkWithShapeDataTable() {
        using MemoryStream stream = BuildSampleVisio();

        List<ReaderChunk> chunks = DocumentReaderVisioExtensions.ReadVisio(
            stream,
            sourceName: " topology.vsdx ",
            readerOptions: new ReaderOptions { ComputeHashes = true }).ToList();

        ReaderChunk chunk = Assert.Single(chunks);
        Assert.Equal(ReaderInputKind.Visio, chunk.Kind);
        Assert.Equal("topology.vsdx", chunk.Location.Path);
        Assert.Equal(1, chunk.Location.Page);
        Assert.Contains("Gateway", chunk.Markdown, StringComparison.Ordinal);
        Assert.Contains("Gateway -> Service", chunk.Markdown, StringComparison.Ordinal);
        Assert.False(string.IsNullOrWhiteSpace(chunk.SourceId));
        Assert.False(string.IsNullOrWhiteSpace(chunk.SourceHash));
        Assert.False(string.IsNullOrWhiteSpace(chunk.ChunkHash));
        ReaderTable table = Assert.Single(chunk.Tables!);
        Assert.Equal("visio-shape-data", table.Kind);
        Assert.Contains(table.Rows, row => row[0] == "shape" && row[3] == "Owner" && row[5] == "Platform");
        Assert.Contains(table.Rows, row => row[0] == "connector" && row[3] == "Protocol" && row[5] == "TLS");
    }

    [Fact]
    public void DocumentReaderVisio_ReadVisioDocument_MapsBlocksLinksTablesAndPreviewAssets() {
        using MemoryStream stream = BuildSampleVisio();

        OfficeDocumentReadResult result = DocumentReaderVisioExtensions.ReadVisioDocument(
            stream,
            sourceName: "topology.vsdx",
            readerOptions: new ReaderOptions { ComputeHashes = true },
            visioOptions: new ReaderVisioOptions { IncludeSvgPreviewAssets = true });

        Assert.Equal(ReaderInputKind.Visio, result.Kind);
        Assert.Equal("topology.vsdx", result.Source.Path);
        Assert.Equal("Topology", result.Source.Title);
        Assert.Contains("officeimo.reader.visio", result.CapabilitiesUsed);
        Assert.Contains("officeimo.visio.inspection-snapshot", result.CapabilitiesUsed);
        Assert.Contains("officeimo.visio.svg-preview", result.CapabilitiesUsed);
        OfficeDocumentPage page = Assert.Single(result.Pages);
        Assert.Equal("Topology", page.Name);
        Assert.Contains(result.Blocks, block => block.Kind == "shape" && block.Text.Contains("Gateway", StringComparison.Ordinal));
        Assert.Contains(result.Blocks, block => block.Kind == "connector" && block.Text.Contains("TLS", StringComparison.Ordinal));
        Assert.Same(Assert.Single(result.Tables), Assert.Single(page.Tables));
        Assert.Contains(result.Links, link => link.Uri == "https://example.test/gateway" && link.Text == "Gateway details");
        Assert.Contains(result.Links, link => link.Uri == "https://example.test/flow" && link.Text == "Flow details");
        OfficeDocumentAsset asset = Assert.Single(result.Assets);
        Assert.Equal("preview-svg", asset.Kind);
        Assert.Equal("image/svg+xml", asset.MediaType);
        Assert.Equal("visio-page-0001-preview-svg.svg", asset.FileName);
        Assert.True(asset.LengthBytes > 0);
        Assert.False(string.IsNullOrWhiteSpace(asset.PayloadHash));
        Assert.NotNull(asset.PayloadBytes);
        Assert.Same(asset, Assert.Single(page.Assets));
    }

    [Fact]
    public void DocumentReaderVisio_ReadVisioTables_ReturnsShapeDataTablesOnly() {
        using MemoryStream stream = BuildSampleVisio();

        IReadOnlyList<ReaderTable> tables = DocumentReaderVisioExtensions.ReadVisioTables(
            stream,
            sourceName: "tables-only.vsdx");

        ReaderTable table = Assert.Single(tables);
        Assert.Equal("visio-shape-data", table.Kind);
        Assert.Equal("tables-only.vsdx", table.Location?.Path);
        Assert.Equal(1, table.Location?.Page);
        Assert.Contains(table.Rows, row => row[0] == "shape" && row[3] == "Owner" && row[5] == "Platform");
        Assert.Contains(table.Rows, row => row[0] == "connector" && row[3] == "Protocol" && row[5] == "TLS");

        using MemoryStream exportStream = BuildSampleVisio();
        ReaderTableExportBundle export = Assert.Single(DocumentReaderVisioExtensions.ReadVisioTableExports(
            exportStream,
            sourceName: "tables-only.vsdx"));
        Assert.Equal("tables-only-page-0001-table-0000", export.Id);
        Assert.Contains("Owner,Platform", export.Csv, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderVisio_ReadVisioDocumentJson_EmitsStableTransportShape() {
        using MemoryStream stream = BuildSampleVisio();

        string json = DocumentReaderVisioExtensions.ReadVisioDocumentJson(
            stream,
            sourceName: "topology.vsdx",
            visioOptions: new ReaderVisioOptions { IncludeSvgPreviewAssets = true });

        using JsonDocument document = JsonDocument.Parse(json);
        JsonElement root = document.RootElement;
        Assert.Equal(OfficeDocumentReadResultSchema.Id, root.GetProperty("schemaId").GetString());
        Assert.Equal(OfficeDocumentReadResultSchema.Version, root.GetProperty("schemaVersion").GetInt32());
        Assert.Equal("Visio", root.GetProperty("kind").GetString());
        Assert.Equal("Topology", root.GetProperty("source").GetProperty("title").GetString());
        Assert.Equal(1, root.GetProperty("pages").GetArrayLength());
        Assert.True(root.GetProperty("blocks").GetArrayLength() >= 3);
        Assert.Equal("preview-svg", root.GetProperty("assets")[0].GetProperty("kind").GetString());
        Assert.Equal("visio-page-0001-preview-svg.svg", root.GetProperty("assets")[0].GetProperty("fileName").GetString());
        Assert.Contains("Gateway", root.GetProperty("markdown").GetString(), StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderVisio_ReadVisioDocument_AlignsLinksAndAssetsWithSnapshotPageOrder() {
        using MemoryStream stream = BuildOutOfOrderVisioPages();

        OfficeDocumentReadResult result = DocumentReaderVisioExtensions.ReadVisioDocument(
            stream,
            sourceName: "out-of-order.vsdx",
            visioOptions: new ReaderVisioOptions { IncludeSvgPreviewAssets = true });

        Assert.Equal(new[] { "Earlier", "Later" }, result.Pages.Select(page => page.Name).ToArray());

        OfficeDocumentLink earlierLink = Assert.Single(result.Links, link => link.Text == "Earlier details");
        OfficeDocumentLink laterLink = Assert.Single(result.Links, link => link.Text == "Later details");
        Assert.Equal(1, earlierLink.Location.Page);
        Assert.Equal(2, laterLink.Location.Page);

        OfficeDocumentAsset earlierAsset = Assert.Single(result.Pages[0].Assets);
        OfficeDocumentAsset laterAsset = Assert.Single(result.Pages[1].Assets);
        Assert.Equal("visio-page-0001-preview-svg", earlierAsset.Id);
        Assert.Equal("visio-page-0002-preview-svg", laterAsset.Id);
        Assert.Contains("Earlier Shape", Encoding.UTF8.GetString(earlierAsset.PayloadBytes!), StringComparison.Ordinal);
        Assert.Contains("Later Shape", Encoding.UTF8.GetString(laterAsset.PayloadBytes!), StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderVisio_Registration_DispatchesVisioStream() {
        try {
            DocumentReaderVisioRegistrationExtensions.RegisterVisioHandler();
            using MemoryStream stream = BuildSampleVisio();

            List<ReaderChunk> chunks = DocumentReader.Read(stream, "registered.vsdx").ToList();

            ReaderChunk chunk = Assert.Single(chunks);
            Assert.Equal(ReaderInputKind.Visio, chunk.Kind);
            Assert.Contains("Gateway", chunk.Markdown, StringComparison.Ordinal);
        } finally {
            DocumentReaderVisioRegistrationExtensions.UnregisterVisioHandler();
        }
    }

    private static MemoryStream BuildSampleVisio() {
        var stream = new MemoryStream();
        VisioDocument document = VisioDocument.Create(stream);
        document.Title = "Topology";
        document.Author = "OfficeIMO";
        VisioPage page = document.AddPage("Topology", 8, 5);
        VisioShape gateway = page.AddRectangle(1.5, 3.5, 1.4, 0.7, "Gateway");
        gateway.SetShapeData("Owner", "Platform", "Owner", VisioShapeDataType.String);
        gateway.AddHyperlink("https://example.test/gateway", "Gateway details");
        VisioShape service = page.AddRectangle(5.5, 3.5, 1.4, 0.7, "Service");
        service.SetShapeData("Tier", "Backend", "Tier", VisioShapeDataType.String);
        VisioConnector connector = page.AddConnector(gateway, service, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);
        connector.Label = "Gateway -> Service";
        connector.SetShapeData("Protocol", "TLS", "Protocol", VisioShapeDataType.String);
        connector.AddHyperlink("https://example.test/flow", "Flow details");
        document.Save();
        stream.Position = 0;
        return stream;
    }

    private static MemoryStream BuildOutOfOrderVisioPages() {
        var stream = new MemoryStream();
        VisioDocument document = VisioDocument.Create(stream);

        VisioPage later = document.AddPage("Later", id: 10);
        VisioShape laterShape = later.AddRectangle(1.5, 3.5, 1.4, 0.7, "Later Shape");
        laterShape.AddHyperlink("https://example.test/later", "Later details");

        VisioPage earlier = document.AddPage("Earlier", id: 5);
        VisioShape earlierShape = earlier.AddRectangle(1.5, 3.5, 1.4, 0.7, "Earlier Shape");
        earlierShape.AddHyperlink("https://example.test/earlier", "Earlier details");

        document.Save();
        stream.Position = 0;
        return stream;
    }
}
