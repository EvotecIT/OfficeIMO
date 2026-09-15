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

    [Fact]
    public void MimeAliasPublishesTheSameSmimeProtectionRowsAsEml() {
        string[] eml = OfficeOperationCapabilityCatalog.FindByExtension(".eml")
            .Where(row => row.Id.StartsWith("protection:", StringComparison.Ordinal))
            .Select(row => row.Id)
            .OrderBy(id => id, StringComparer.Ordinal)
            .ToArray();
        string[] mime = OfficeOperationCapabilityCatalog.FindByExtension(".mime")
            .Where(row => row.Id.StartsWith("protection:", StringComparison.Ordinal))
            .Select(row => row.Id)
            .OrderBy(id => id, StringComparer.Ordinal)
            .ToArray();

        Assert.NotEmpty(eml);
        Assert.Equal(eml, mime);
    }

    [Theory]
    [InlineData(".pst", "Email.Store.Pst", OfficeOperationSupportState.Supported, OfficeOperationSupportState.Supported)]
    [InlineData(".ost", "Email.Store.Ost", OfficeOperationSupportState.Unsupported, OfficeOperationSupportState.Unsupported)]
    [InlineData(".olm", "Email.Store.Olm", OfficeOperationSupportState.Unsupported, OfficeOperationSupportState.Unsupported)]
    [InlineData(".mbox", "Email.Store.Mbox", OfficeOperationSupportState.Supported, OfficeOperationSupportState.Supported)]
    [InlineData(".emlx", "Email.Store.Emlx", OfficeOperationSupportState.Supported, OfficeOperationSupportState.Supported)]
    public void EmailStoresPublishFormatSpecificLifecycleBoundaries(
        string extension,
        string formatId,
        OfficeOperationSupportState createState,
        OfficeOperationSupportState editState) {
        OfficeOperationCapability[] rows = OfficeOperationCapabilityCatalog.FindByExtension(extension)
            .Where(row => row.SourceCatalog == "OfficeIMO.NativeLifecycle" && row.FormatId == formatId)
            .ToArray();

        Assert.NotEmpty(rows);
        Assert.Contains(rows, row => row.Operation == OfficeOperationKind.Create && row.State == createState);
        Assert.Contains(rows, row => row.Operation == OfficeOperationKind.Read && row.State == OfficeOperationSupportState.Supported);
        Assert.Contains(rows, row => row.Operation == OfficeOperationKind.Edit && row.State == editState);
        Assert.Contains(rows, row => row.Operation == OfficeOperationKind.Inspect && row.State == OfficeOperationSupportState.Supported);
        Assert.All(rows, row => Assert.False(string.IsNullOrWhiteSpace(row.Limitation)));
    }

    [Fact]
    public void PstLifecycleNamesOnlyPublicStoreEntryPoints() {
        OfficeOperationCapability[] rows = OfficeOperationCapabilityCatalog.FindByExtension(".pst")
            .Where(row => row.SourceCatalog == "OfficeIMO.NativeLifecycle" && row.FormatId == "Email.Store.Pst")
            .ToArray();

        Assert.NotEmpty(rows);
        Assert.All(rows, row => {
            Assert.Contains("EmailStorePstWriter", row.PublicApi, StringComparison.Ordinal);
            Assert.Contains("EmailStoreConverter.ConvertToPst", row.PublicApi, StringComparison.Ordinal);
            Assert.Contains("EmailStoreConverter.MergeToPst", row.PublicApi, StringComparison.Ordinal);
            Assert.Contains("EmailStorePstMutationTransaction", row.PublicApi, StringComparison.Ordinal);
            Assert.DoesNotContain("EmailStorePstMerger", row.PublicApi, StringComparison.Ordinal);
            Assert.DoesNotContain("RewriteUnicodePst", row.PublicApi, StringComparison.Ordinal);
        });
    }

    [Theory]
    [InlineData(".docx", "OfficeIMO.Word")]
    [InlineData(".xlsx", "OfficeIMO.Excel")]
    [InlineData(".pptx", "OfficeIMO.PowerPoint")]
    [InlineData(".vsdx", "OfficeIMO.Visio")]
    [InlineData(".one", "OfficeIMO.OneNote")]
    [InlineData(".mpp", "OfficeIMO.Project")]
    [InlineData(".mpt", "OfficeIMO.Project")]
    [InlineData(".mpx", "OfficeIMO.Project")]
    public void NativeFormatsPublishCreateReadAndEditLifecycle(string extension, string packageId) {
        IReadOnlyList<OfficeOperationCapability> rows = OfficeOperationCapabilityCatalog.FindByExtension(extension);

        Assert.All(new[] { OfficeOperationKind.Create, OfficeOperationKind.Read, OfficeOperationKind.Edit }, operation =>
            Assert.Contains(rows, row =>
                row.SourceCatalog == "OfficeIMO.NativeLifecycle" &&
                row.PackageId == packageId &&
                row.Operation == operation &&
                row.State == OfficeOperationSupportState.Supported));
    }

    [Theory]
    [InlineData(".html", OfficeOperationKind.Edit)]
    [InlineData(".pdf", OfficeOperationKind.Edit)]
    [InlineData(".eml", OfficeOperationKind.Edit)]
    [InlineData(".mpp", OfficeOperationKind.Create)]
    [InlineData(".mpp", OfficeOperationKind.Read)]
    [InlineData(".mpp", OfficeOperationKind.Edit)]
    [InlineData(".mpp", OfficeOperationKind.Preserve)]
    [InlineData(".mpp", OfficeOperationKind.Inspect)]
    [InlineData(".mpx", OfficeOperationKind.Create)]
    [InlineData(".mpx", OfficeOperationKind.Read)]
    [InlineData(".mpx", OfficeOperationKind.Edit)]
    [InlineData(".mpx", OfficeOperationKind.Preserve)]
    [InlineData(".mpx", OfficeOperationKind.Inspect)]
    public void NativeLifecycleRetainsBoundariesOnSupportedOperations(string extension, OfficeOperationKind operation) {
        OfficeOperationCapability row = Assert.Single(OfficeOperationCapabilityCatalog.FindByExtension(extension), row =>
            row.SourceCatalog == "OfficeIMO.NativeLifecycle" && row.Operation == operation);

        Assert.Equal(OfficeOperationSupportState.Supported, row.State);
        Assert.False(string.IsNullOrWhiteSpace(row.Limitation));
    }

    [Fact]
    public void ProjectConversionPublishesDirectionalAssessedLossBoundaries() {
        OfficeOperationCapability[] binaryRows = OfficeOperationCapabilityCatalog.FindByExtension(".mpp")
            .Where(row => row.SourceCatalog == "OfficeIMO.NativeLifecycle" && row.Operation == OfficeOperationKind.Convert)
            .ToArray();
        OfficeOperationCapability[] mpxRows = OfficeOperationCapabilityCatalog.FindByExtension(".mpx")
            .Where(row => row.SourceCatalog == "OfficeIMO.NativeLifecycle" && row.Operation == OfficeOperationKind.Convert)
            .ToArray();

        Assert.Equal(new[] { "Project.Mpx", "Project.Xml" }, binaryRows.Select(row => row.TargetFormatId).OrderBy(value => value).ToArray());
        Assert.Equal(new[] { "Project.MppMpt", "Project.Xml" }, mpxRows.Select(row => row.TargetFormatId).OrderBy(value => value).ToArray());
        Assert.All(binaryRows.Concat(mpxRows), row => {
            Assert.Equal(OfficeOperationSupportState.Partial, row.State);
            Assert.Contains("explicit caller permission", row.Limitation, StringComparison.Ordinal);
        });
    }

    [Fact]
    public void ProjectBinaryAndMpxLifecycleRowsKeepDistinctFormatContracts() {
        OfficeOperationCapability[] binaryRows = OfficeOperationCapabilityCatalog.FindByExtension(".mpp")
            .Where(row => row.SourceCatalog == "OfficeIMO.NativeLifecycle")
            .ToArray();
        OfficeOperationCapability[] mpxRows = OfficeOperationCapabilityCatalog.FindByExtension(".mpx")
            .Where(row => row.SourceCatalog == "OfficeIMO.NativeLifecycle")
            .ToArray();

        Assert.NotEmpty(binaryRows);
        Assert.NotEmpty(mpxRows);
        Assert.All(binaryRows, row => Assert.Equal("Project.MppMpt", row.FormatId));
        Assert.All(mpxRows, row => Assert.Equal("Project.Mpx", row.FormatId));
        Assert.Contains(binaryRows, row => row.Operation == OfficeOperationKind.Create && row.Limitation.Contains("Global.mpt", StringComparison.Ordinal));
        Assert.Contains(mpxRows, row => row.Operation == OfficeOperationKind.Read && row.Limitation.Contains("4.0/4.1", StringComparison.Ordinal));
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
