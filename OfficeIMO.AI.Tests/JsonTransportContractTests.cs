using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeIMO.AI;
using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed class JsonTransportContractTests {
    [Fact]
    public void GeneratedTransportPreservesSnapshotIdentityAndReviewArtifactShape() {
        byte[] image = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aX1sAAAAASUVORK5CYII=");
        var document = OfficeAiDocument.FromReadResult([1, 2, 3], new OfficeDocumentReadResult {
            Blocks = [new OfficeDocumentBlock { Id = "paragraph", Text = "Stable source evidence" }]
        }, [new OfficeAiImage("page", 1, "image/png", image, 1, 1)]);
        // Persisted citations bind this exact byte shape, including the original property order and casing.
        byte[] previousSnapshotShape = JsonSerializer.SerializeToUtf8Bytes(new {
            sourceHash = document.SourceHash, evidence = document.Evidence, pages = document.Pages,
            pageProvenance = document.PageProvenance, hasSourceDiagnostics = document.HasSourceDiagnostics,
            images = document.Images.Select(item => new { item.Id, item.Page, item.MediaType, item.Width, item.Height, item.ContentHash })
        });
        Assert.Equal(Convert.ToHexString(SHA256.HashData(previousSnapshotShape)).ToLowerInvariant(), document.SnapshotHash);
        var result = new OfficeAiResult { Profile = new OfficeAiExecutionProfile { Id = "test", Provider = "fixture", Model = "synthetic" }, RequestId = "transport-contract", SourceHash = document.SourceHash,
            SnapshotHash = document.SnapshotHash, Operation = OfficeAiOperation.Ask, Status = OfficeAiResultStatus.Completed };
        string previousReportShape = JsonSerializer.Serialize(new {
            schema = "officeimo.ai.report.v1",
            source = new { document.SourceHash, document.SnapshotHash, document.SourceByteLength, document.PageProvenance, document.Pages },
            evidence = document.Evidence,
            images = document.Images.Select(item => new { item.Id, item.Page, item.MediaType, item.Width, item.Height, item.ByteLength, item.ContentHash }),
            result
        }, new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase, WriteIndented = true,
            Converters = { new JsonStringEnumConverter() } });
        Assert.Equal(previousReportShape, OfficeAiArtifacts.SerializeReport(document, result));
    }
}
