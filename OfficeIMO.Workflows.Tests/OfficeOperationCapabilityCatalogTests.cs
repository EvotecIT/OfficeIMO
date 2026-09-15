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
