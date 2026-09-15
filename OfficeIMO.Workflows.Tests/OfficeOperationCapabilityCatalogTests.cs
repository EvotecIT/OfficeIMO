using System.Text.Json;
using OfficeIMO;
using OfficeIMO.Workflows;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class OfficeOperationCapabilityCatalogTests {
    [Fact]
    public void CatalogProjectsEveryRequiredOperationAndOutcomeClass() {
        OfficeOperationKind[] requiredOperations = {
            OfficeOperationKind.Create,
            OfficeOperationKind.Read,
            OfficeOperationKind.Edit,
            OfficeOperationKind.Preserve,
            OfficeOperationKind.Inspect,
            OfficeOperationKind.Convert,
            OfficeOperationKind.Export
        };
        OfficeOperationSupportState[] requiredStates = {
            OfficeOperationSupportState.Supported,
            OfficeOperationSupportState.Partial,
            OfficeOperationSupportState.Preserved,
            OfficeOperationSupportState.Rejected,
            OfficeOperationSupportState.Unsupported
        };

        Assert.All(requiredOperations, operation =>
            Assert.Contains(OfficeOperationCapabilityCatalog.All, row => row.Operation == operation));
        Assert.All(requiredStates, state =>
            Assert.Contains(OfficeOperationCapabilityCatalog.All, row => row.State == state));
    }

    [Fact]
    public void CatalogKeepsDetailedOwnersAndEvidence() {
        Assert.Contains(OfficeOperationCapabilityCatalog.All, row =>
            row.SourceCatalog == "OfficeIMO.Word.LegacyDoc" &&
            row.PackageId == "OfficeIMO.Word" &&
            row.Operation == OfficeOperationKind.Read);
        Assert.Contains(OfficeOperationCapabilityCatalog.All, row =>
            row.SourceCatalog == "OfficeConversionCapabilityCatalog" &&
            row.CapabilityId == "docx-pdf" &&
            row.PackageId == "OfficeIMO.Word.Pdf");
        Assert.Contains(OfficeOperationCapabilityCatalog.All, row =>
            row.SourceCatalog == "OfficeIMO.ProtectedContent" &&
            row.State == OfficeOperationSupportState.Rejected);
        Assert.Contains(OfficeOperationCapabilityCatalog.All, row =>
            row.SourceCatalog == OfficeProvenanceWorkflowCatalog.Id &&
            row.Operation == OfficeOperationKind.Inspect);

        Assert.All(OfficeOperationCapabilityCatalog.All, row => {
            Assert.False(string.IsNullOrWhiteSpace(row.PublicApi));
            Assert.False(string.IsNullOrWhiteSpace(row.Evidence));
            Assert.False(string.IsNullOrWhiteSpace(row.SourceCatalog));
        });
    }

    [Theory]
    [InlineData(".docx", "OfficeIMO.Word")]
    [InlineData(".xlsx", "OfficeIMO.Excel")]
    [InlineData(".pptx", "OfficeIMO.PowerPoint")]
    [InlineData(".vsdx", "OfficeIMO.Visio")]
    public void ProtectionRowsExposeOnlyExtensionsOwnedByTheirPackage(string extension, string packageId) {
        OfficeOperationCapability[] rows = OfficeOperationCapabilityCatalog.FindByExtension(extension)
            .Where(row => row.Id.StartsWith("protection:", StringComparison.Ordinal))
            .ToArray();

        Assert.NotEmpty(rows);
        Assert.All(rows, row => Assert.Equal(packageId, row.PackageId));
    }

    [Theory]
    [InlineData(".docx", "OfficeIMO.Word")]
    [InlineData(".xlsx", "OfficeIMO.Excel")]
    [InlineData(".pptx", "OfficeIMO.PowerPoint")]
    [InlineData(".vsdx", "OfficeIMO.Visio")]
    [InlineData(".one", "OfficeIMO.OneNote")]
    public void NativeFormatsPublishCreateReadAndEditLifecycle(string extension, string packageId) {
        IReadOnlyList<OfficeOperationCapability> rows = OfficeOperationCapabilityCatalog.FindByExtension(extension);

        Assert.All(new[] { OfficeOperationKind.Create, OfficeOperationKind.Read, OfficeOperationKind.Edit }, operation =>
            Assert.Contains(rows, row =>
                row.SourceCatalog == "OfficeIMO.NativeLifecycle" &&
                row.PackageId == packageId &&
                row.Operation == operation &&
                row.State == OfficeOperationSupportState.Supported));
    }

    [Fact]
    public void EpubLifecycleDoesNotOverstateAuthoringSupport() {
        IReadOnlyList<OfficeOperationCapability> rows = OfficeOperationCapabilityCatalog.FindByExtension(".epub");

        Assert.Contains(rows, row => row.SourceCatalog == "OfficeIMO.NativeLifecycle" &&
            row.Operation == OfficeOperationKind.Read && row.State == OfficeOperationSupportState.Supported);
        Assert.All(new[] { OfficeOperationKind.Create, OfficeOperationKind.Edit }, operation =>
            Assert.Contains(rows, row => row.SourceCatalog == "OfficeIMO.NativeLifecycle" &&
                row.Operation == operation && row.State == OfficeOperationSupportState.Unsupported));
    }

    [Fact]
    public void CatalogSerializesDeterministicallyAndFiltersByExtension() {
        string first = OfficeOperationCapabilityCatalog.ToJson();
        string second = OfficeOperationCapabilityCatalog.ToJson();

        Assert.Equal(first, second);
        using JsonDocument parsed = JsonDocument.Parse(first);
        Assert.Equal(OfficeOperationCapabilityCatalog.Id, parsed.RootElement.GetProperty("id").GetString());
        Assert.Equal(OfficeOperationCapabilityCatalog.SchemaVersion, parsed.RootElement.GetProperty("schemaVersion").GetInt32());
        Assert.Equal(OfficeOperationCapabilityCatalog.All.Count,
            parsed.RootElement.GetProperty("capabilities").GetArrayLength());

        IReadOnlyList<OfficeOperationCapability> docx = OfficeOperationCapabilityCatalog.FindByExtension("docx");
        Assert.NotEmpty(docx);
        Assert.Contains(docx, row => row.CapabilityId == "docx-pdf");
        Assert.All(docx, row => Assert.Contains(".docx", row.Extensions, StringComparer.OrdinalIgnoreCase));
    }

    [Fact]
    public void CatalogIdsAreUniqueAndMarkdownRetainsOwnership() {
        Assert.Equal(
            OfficeOperationCapabilityCatalog.All.Count,
            OfficeOperationCapabilityCatalog.All.Select(row => row.Id).Distinct(StringComparer.Ordinal).Count());

        string markdown = OfficeOperationCapabilityCatalog.ToMarkdown();
        Assert.Contains("| Package | Format | Target | Operation | State |", markdown, StringComparison.Ordinal);
        Assert.Contains("`OfficeIMO.Word.Pdf`", markdown, StringComparison.Ordinal);
        Assert.Contains("OfficeConversionCapabilityCatalog", markdown, StringComparison.Ordinal);
    }
}
