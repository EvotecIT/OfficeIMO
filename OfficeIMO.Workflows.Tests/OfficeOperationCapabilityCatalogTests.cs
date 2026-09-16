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
    [InlineData(".docm", "docx-pdf", OfficeOperationKind.Convert)]
    [InlineData(".dotm", "docx-markdown", OfficeOperationKind.Convert)]
    [InlineData(".xlsm", "xlsx-pdf", OfficeOperationKind.Convert)]
    [InlineData(".xlam", "xlsx-csv", OfficeOperationKind.Convert)]
    [InlineData(".pptm", "pptx-pdf", OfficeOperationKind.Convert)]
    [InlineData(".ppam", "pptx-svg", OfficeOperationKind.Export)]
    public void ModernOfficeFamilyVariantsRetainCanonicalConversionRoutes(
        string extension,
        string capabilityId,
        OfficeOperationKind operation) {
        Assert.Contains(
            OfficeConversionCapabilityCatalog.FindBySourceExtension(extension),
            route => route.Id == capabilityId);
        Assert.Contains(
            OfficeOperationCapabilityCatalog.FindByExtension(extension),
            row => row.Operation == operation && row.CapabilityId == capabilityId);
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
    [InlineData(".docm", ".dotm")]
    [InlineData(".xlsm", ".xltm")]
    [InlineData(".xlsm", ".xlam")]
    [InlineData(".pptm", ".potm")]
    [InlineData(".pptm", ".ppsm")]
    [InlineData(".pptm", ".ppam")]
    public void MacroEnabledFamilyVariantsPublishTheSameVbaProtectionRows(string canonical, string variant) {
        string[] expected = OfficeOperationCapabilityCatalog.FindByExtension(canonical)
            .Where(row => row.Id.StartsWith("protection:", StringComparison.Ordinal) &&
                row.CapabilityId.Contains("Vba", StringComparison.OrdinalIgnoreCase))
            .Select(row => row.Id)
            .OrderBy(id => id, StringComparer.Ordinal)
            .ToArray();
        string[] actual = OfficeOperationCapabilityCatalog.FindByExtension(variant)
            .Where(row => row.Id.StartsWith("protection:", StringComparison.Ordinal) &&
                row.CapabilityId.Contains("Vba", StringComparison.OrdinalIgnoreCase))
            .Select(row => row.Id)
            .OrderBy(id => id, StringComparer.Ordinal)
            .ToArray();

        Assert.NotEmpty(expected);
        Assert.Equal(expected, actual);
    }

    [Fact]
    public void EveryModernMacroEnabledFormatPublishesVbaProtectionRows() {
        OfficeFormatDescriptor[] formats = OfficeIMO.Word.WordFormatCatalog.All
            .Concat(OfficeIMO.Excel.ExcelFormatCatalog.All)
            .Concat(OfficeIMO.PowerPoint.PowerPointFormatCatalog.All)
            .Where(format => format.Generation == OfficeFormatGeneration.Modern && format.IsMacroEnabled)
            .ToArray();

        Assert.NotEmpty(formats);
        Assert.All(formats, format =>
            Assert.Contains(OfficeOperationCapabilityCatalog.FindByExtension(format.Extension), row =>
                row.Id.StartsWith("protection:", StringComparison.Ordinal) &&
                row.CapabilityId.Contains("Vba", StringComparison.OrdinalIgnoreCase)));
    }

    [Theory]
    [InlineData(".docx", ".docm")]
    [InlineData(".docx", ".dotm")]
    [InlineData(".xlsx", ".xlsm")]
    [InlineData(".xlsx", ".xltm")]
    [InlineData(".xlsx", ".xlam")]
    [InlineData(".pptx", ".pptm")]
    [InlineData(".pptx", ".potm")]
    [InlineData(".pptx", ".ppsm")]
    [InlineData(".pptx", ".ppam")]
    public void MacroEnabledOpenXmlVariantsPublishTheSamePackageSignatureRows(string canonical, string variant) {
        string[] expected = OfficeOperationCapabilityCatalog.FindByExtension(canonical)
            .Where(row => row.Id.StartsWith("protection:", StringComparison.Ordinal) &&
                row.CapabilityId == "opc-package-signature")
            .Select(row => row.Id)
            .OrderBy(id => id, StringComparer.Ordinal)
            .ToArray();
        string[] actual = OfficeOperationCapabilityCatalog.FindByExtension(variant)
            .Where(row => row.Id.StartsWith("protection:", StringComparison.Ordinal) &&
                row.CapabilityId == "opc-package-signature")
            .Select(row => row.Id)
            .OrderBy(id => id, StringComparer.Ordinal)
            .ToArray();

        Assert.NotEmpty(expected);
        Assert.Equal(expected, actual);
    }

    [Theory]
    [InlineData(".docx", ".docm")]
    [InlineData(".docx", ".dotm")]
    [InlineData(".xlsx", ".xlsm")]
    [InlineData(".xlsx", ".xltm")]
    [InlineData(".xlsx", ".xlam")]
    [InlineData(".pptx", ".pptm")]
    [InlineData(".pptx", ".potm")]
    [InlineData(".pptx", ".ppsm")]
    [InlineData(".pptx", ".ppam")]
    public void MacroEnabledOpenXmlVariantsPublishTheSamePasswordProtectionRows(string canonical, string variant) {
        string[] expected = OfficeOperationCapabilityCatalog.FindByExtension(canonical)
            .Where(row => row.Id.StartsWith("protection:", StringComparison.Ordinal) &&
                row.CapabilityId == "ooxml-password")
            .Select(row => row.Id)
            .OrderBy(id => id, StringComparer.Ordinal)
            .ToArray();
        string[] actual = OfficeOperationCapabilityCatalog.FindByExtension(variant)
            .Where(row => row.Id.StartsWith("protection:", StringComparison.Ordinal) &&
                row.CapabilityId == "ooxml-password")
            .Select(row => row.Id)
            .OrderBy(id => id, StringComparer.Ordinal)
            .ToArray();

        Assert.NotEmpty(expected);
        Assert.Equal(expected, actual);
    }

    [Theory]
    [InlineData(".dot", "OfficeIMO.Word")]
    [InlineData(".xlt", "OfficeIMO.Excel")]
    [InlineData(".xla", "OfficeIMO.Excel")]
    [InlineData(".xlm", "OfficeIMO.Excel")]
    [InlineData(".xlw", "OfficeIMO.Excel")]
    [InlineData(".pot", "OfficeIMO.PowerPoint")]
    [InlineData(".pps", "OfficeIMO.PowerPoint")]
    [InlineData(".ppa", "OfficeIMO.PowerPoint")]
    public void LegacyVariantsPublishReadAndModernizationRows(string extension, string packageId) {
        OfficeOperationCapability[] rows = OfficeOperationCapabilityCatalog.FindByExtension(extension)
            .Where(row => row.Id.StartsWith("legacy:", StringComparison.Ordinal) && row.PackageId == packageId)
            .ToArray();

        Assert.Contains(rows, row => row.Operation == OfficeOperationKind.Read);
        Assert.Contains(rows, row => row.Operation == OfficeOperationKind.Convert && row.TargetFormatId != null);
    }

    [Theory]
    [InlineData(".xls", true)]
    [InlineData(".xlt", false)]
    [InlineData(".xla", false)]
    [InlineData(".xlm", false)]
    [InlineData(".xlw", false)]
    [InlineData(".ppt", true)]
    [InlineData(".pot", true)]
    [InlineData(".pps", true)]
    [InlineData(".ppa", false)]
    public void LegacyEditRowsExposeOnlyProvenWriterExtensions(string extension, bool expected) {
        OfficeOperationCapability[] rows = OfficeOperationCapabilityCatalog.FindByExtension(extension)
            .Where(row => row.Id.StartsWith("legacy:", StringComparison.Ordinal))
            .ToArray();

        Assert.Equal(expected, rows.Any(row => row.Operation == OfficeOperationKind.Edit));
    }

    [Theory]
    [InlineData(".xls", true)]
    [InlineData(".xlt", false)]
    [InlineData(".xla", false)]
    [InlineData(".xlm", false)]
    [InlineData(".xlw", false)]
    [InlineData(".ppt", true)]
    [InlineData(".pot", true)]
    [InlineData(".pps", true)]
    [InlineData(".ppa", true)]
    public void LegacyPreserveRowsExposeOnlySourceRoundTripExtensions(string extension, bool expected) {
        OfficeOperationCapability[] rows = OfficeOperationCapabilityCatalog.FindByExtension(extension)
            .Where(row => row.Id.StartsWith("legacy:", StringComparison.Ordinal))
            .ToArray();

        Assert.Equal(expected, rows.Any(row => row.Operation == OfficeOperationKind.Preserve));
    }

    [Theory]
    [InlineData(".xls", true)]
    [InlineData(".xlsb", true)]
    [InlineData(".xlt", false)]
    [InlineData(".xla", false)]
    [InlineData(".xlm", false)]
    [InlineData(".xlw", false)]
    [InlineData(".ppt", true)]
    [InlineData(".pot", true)]
    [InlineData(".pps", true)]
    [InlineData(".ppa", false)]
    public void LegacyCreationRowsExposeOnlyProvenWriterExtensions(string extension, bool expected) {
        bool actual = OfficeOperationCapabilityCatalog.FindByExtension(extension)
            .Any(row => row.Id.StartsWith("legacy:", StringComparison.Ordinal) &&
                row.Operation == OfficeOperationKind.Create);

        Assert.Equal(expected, actual);
    }

    [Theory]
    [InlineData(".eml", "email-eml-msg", "OfficeIMO.Email", ".msg")]
    [InlineData(".msg", "email-msg-eml", "OfficeIMO.Email", ".eml")]
    [InlineData(".bib", "bibliography-bibtex-csl-json", "OfficeIMO.Bibliography", ".json")]
    [InlineData(".medline", "bibliography-nbib-ris", "OfficeIMO.Bibliography", ".ris")]
    [InlineData(".pages", "pages-docx", "OfficeIMO.Word.IWork", ".docx")]
    [InlineData(".numbers", "numbers-xlsx", "OfficeIMO.Excel.IWork", ".xlsx")]
    [InlineData(".key", "keynote-pptx", "OfficeIMO.PowerPoint.IWork", ".pptx")]
    public void CrossFormatCodecAndAdapterRoutesRemainDiscoverable(
        string sourceExtension,
        string routeId,
        string packageId,
        string targetExtension) {
        OfficeConversionCapability route = Assert.Single(
            OfficeConversionCapabilityCatalog.FindBySourceExtension(sourceExtension),
            candidate => candidate.Id == routeId);
        Assert.Equal(packageId, route.PackageId);
        Assert.Equal(targetExtension, route.TargetExtension);
        if (routeId.StartsWith("bibliography-", StringComparison.Ordinal)) {
            Assert.Contains(".Document.Save(", route.Api, StringComparison.Ordinal);
        }
        Assert.Contains(
            OfficeOperationCapabilityCatalog.FindByExtension(sourceExtension),
            row => row.Operation == OfficeOperationKind.Convert &&
                row.CapabilityId == routeId && row.PackageId == packageId);
    }

    [Fact]
    public void XlsbConversionRowsExcludeTheXlsbTargetFromItsOwnSourceSet() {
        OfficeOperationCapability[] xlsbRows = OfficeOperationCapabilityCatalog.FindByExtension(".xlsb")
            .Where(row => row.Id.StartsWith("legacy:", StringComparison.Ordinal) &&
                row.Operation == OfficeOperationKind.Convert)
            .ToArray();
        OfficeOperationCapability[] xlsxRows = OfficeOperationCapabilityCatalog.FindByExtension(".xlsx")
            .Where(row => row.Id.StartsWith("legacy:", StringComparison.Ordinal) &&
                row.Operation == OfficeOperationKind.Convert)
            .ToArray();

        Assert.DoesNotContain(xlsbRows, row => row.TargetFormatId == "Excel.Xlsb");
        Assert.Contains(xlsxRows, row => row.TargetFormatId == "Excel.Xlsb");
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
        Assert.Contains(rows, row => row.Operation == OfficeOperationKind.Export && row.State == OfficeOperationSupportState.Supported);
        Assert.All(rows, row => Assert.False(string.IsNullOrWhiteSpace(row.Limitation)));
    }

    [Fact]
    public void OfflineAddressBookPublishesBoundedReadOnlyLifecycle() {
        OfficeOperationCapability[] rows = OfficeOperationCapabilityCatalog.FindByExtension(".oab")
            .Where(row => row.SourceCatalog == "OfficeIMO.NativeLifecycle" &&
                row.FormatId == "Email.AddressBook.Oab")
            .ToArray();

        Assert.All(new[] { OfficeOperationKind.Read, OfficeOperationKind.Inspect, OfficeOperationKind.Validate }, operation =>
            Assert.Contains(rows, row => row.Operation == operation &&
                row.State == OfficeOperationSupportState.Supported));
        Assert.All(new[] { OfficeOperationKind.Create, OfficeOperationKind.Edit, OfficeOperationKind.Preserve }, operation =>
            Assert.Contains(rows, row => row.Operation == operation &&
                row.State == OfficeOperationSupportState.Unsupported));
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
    [InlineData(".xml", "OfficeIMO.Project")]
    [InlineData(".ics", "OfficeIMO.Email")]
    [InlineData(".vcf", "OfficeIMO.Email")]
    [InlineData(".csv", "OfficeIMO.CSV")]
    [InlineData(".opml", "OfficeIMO.Opml")]
    [InlineData(".mhtml", "OfficeIMO.Mhtml")]
    [InlineData(".adoc", "OfficeIMO.AsciiDoc")]
    [InlineData(".tex", "OfficeIMO.Latex")]
    [InlineData(".bib", "OfficeIMO.Bibliography")]
    [InlineData(".dbk", "OfficeIMO.DocBook")]
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
    [InlineData(".pages")]
    [InlineData(".numbers")]
    [InlineData(".key")]
    public void IWorkLifecycleDoesNotOverstateAuthoringSupport(string extension) {
        IReadOnlyList<OfficeOperationCapability> rows = OfficeOperationCapabilityCatalog.FindByExtension(extension);

        Assert.Contains(rows, row => row.SourceCatalog == "OfficeIMO.NativeLifecycle" &&
            row.Operation == OfficeOperationKind.Read && row.State == OfficeOperationSupportState.Supported);
        Assert.Contains(rows, row => row.SourceCatalog == "OfficeIMO.NativeLifecycle" &&
            row.Operation == OfficeOperationKind.Inspect && row.State == OfficeOperationSupportState.Supported);
        Assert.All(new[] { OfficeOperationKind.Create, OfficeOperationKind.Edit, OfficeOperationKind.Preserve }, operation =>
            Assert.Contains(rows, row => row.SourceCatalog == "OfficeIMO.NativeLifecycle" &&
                row.Operation == operation && row.State == OfficeOperationSupportState.Unsupported));
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
    [InlineData(".xml", OfficeOperationKind.Create)]
    [InlineData(".xml", OfficeOperationKind.Read)]
    [InlineData(".xml", OfficeOperationKind.Edit)]
    [InlineData(".xml", OfficeOperationKind.Preserve)]
    [InlineData(".xml", OfficeOperationKind.Inspect)]
    [InlineData(".xml", OfficeOperationKind.Validate)]
    [InlineData(".mpp", OfficeOperationKind.Validate)]
    [InlineData(".mpx", OfficeOperationKind.Validate)]
    public void NativeLifecycleRetainsBoundariesOnSupportedOperations(string extension, OfficeOperationKind operation) {
        OfficeOperationCapability row = Assert.Single(OfficeOperationCapabilityCatalog.FindByExtension(extension), row =>
            row.SourceCatalog == "OfficeIMO.NativeLifecycle" &&
            row.Operation == operation &&
            (extension != ".xml" || row.PackageId == "OfficeIMO.Project"));

        Assert.Equal(OfficeOperationSupportState.Supported, row.State);
        Assert.False(string.IsNullOrWhiteSpace(row.Limitation));
    }

    [Theory]
    [InlineData(".ics", OfficeOperationKind.Validate)]
    [InlineData(".opml", OfficeOperationKind.Preserve)]
    [InlineData(".adoc", OfficeOperationKind.Read)]
    [InlineData(".tex", OfficeOperationKind.Inspect)]
    [InlineData(".bib", OfficeOperationKind.Edit)]
    [InlineData(".dbk", OfficeOperationKind.Validate)]
    [InlineData(".pages", OfficeOperationKind.Read)]
    public void SupportedNativeRowsRetainTheirDeclaredFormatBoundaries(string extension, OfficeOperationKind operation) {
        OfficeOperationCapability row = Assert.Single(OfficeOperationCapabilityCatalog.FindByExtension(extension), row =>
            row.SourceCatalog == "OfficeIMO.NativeLifecycle" &&
            row.Operation == operation);

        Assert.Equal(OfficeOperationSupportState.Supported, row.State);
        Assert.False(string.IsNullOrWhiteSpace(row.Limitation));
    }

    [Theory]
    [InlineData(".gz")]
    [InlineData(".gzip")]
    [InlineData(".deflate")]
    [InlineData(".br")]
    [InlineData(".brotli")]
    [InlineData(".zlib")]
    public void CompressedCsvExtensionsPublishTheCsvLifecycle(string extension) {
        OfficeOperationCapability[] rows = OfficeOperationCapabilityCatalog.FindByExtension(extension)
            .Where(row => row.SourceCatalog == "OfficeIMO.NativeLifecycle" && row.PackageId == "OfficeIMO.CSV")
            .ToArray();

        Assert.Contains(rows, row => row.Operation == OfficeOperationKind.Create && row.State == OfficeOperationSupportState.Supported);
        Assert.Contains(rows, row => row.Operation == OfficeOperationKind.Read && row.State == OfficeOperationSupportState.Supported);
        Assert.Contains(rows, row => row.Operation == OfficeOperationKind.Validate && row.State == OfficeOperationSupportState.Supported);
    }

    [Fact]
    public void ProjectConversionPublishesDirectionalAssessedLossBoundaries() {
        OfficeOperationCapability[] binaryRows = OfficeOperationCapabilityCatalog.FindByExtension(".mpp")
            .Where(row => row.SourceCatalog == "OfficeIMO.NativeLifecycle" && row.Operation == OfficeOperationKind.Convert)
            .ToArray();
        OfficeOperationCapability[] mpxRows = OfficeOperationCapabilityCatalog.FindByExtension(".mpx")
            .Where(row => row.SourceCatalog == "OfficeIMO.NativeLifecycle" && row.Operation == OfficeOperationKind.Convert)
            .ToArray();
        OfficeOperationCapability[] xmlRows = OfficeOperationCapabilityCatalog.FindByExtension(".xml")
            .Where(row => row.SourceCatalog == "OfficeIMO.NativeLifecycle" &&
                row.PackageId == "OfficeIMO.Project" &&
                row.Operation == OfficeOperationKind.Convert)
            .ToArray();

        Assert.Equal(new[] { "Project.Mpx", "Project.Xml" }, binaryRows.Select(row => row.TargetFormatId).OrderBy(value => value).ToArray());
        Assert.Equal(new[] { "Project.MppMpt", "Project.Xml" }, mpxRows.Select(row => row.TargetFormatId).OrderBy(value => value).ToArray());
        Assert.Equal(new[] { "Project.MppMpt", "Project.Mpx" }, xmlRows.Select(row => row.TargetFormatId).OrderBy(value => value).ToArray());
        Assert.All(binaryRows.Concat(mpxRows).Concat(xmlRows), row => {
            Assert.Equal(OfficeOperationSupportState.Partial, row.State);
            Assert.Contains("explicit caller permission", row.Limitation, StringComparison.Ordinal);
        });
    }

    [Fact]
    public void ProjectXmlBinaryAndMpxLifecycleRowsKeepDistinctFormatContracts() {
        OfficeOperationCapability[] xmlRows = OfficeOperationCapabilityCatalog.FindByExtension(".xml")
            .Where(row => row.SourceCatalog == "OfficeIMO.NativeLifecycle" && row.PackageId == "OfficeIMO.Project")
            .ToArray();
        OfficeOperationCapability[] binaryRows = OfficeOperationCapabilityCatalog.FindByExtension(".mpp")
            .Where(row => row.SourceCatalog == "OfficeIMO.NativeLifecycle")
            .ToArray();
        OfficeOperationCapability[] mpxRows = OfficeOperationCapabilityCatalog.FindByExtension(".mpx")
            .Where(row => row.SourceCatalog == "OfficeIMO.NativeLifecycle")
            .ToArray();

        Assert.NotEmpty(xmlRows);
        Assert.NotEmpty(binaryRows);
        Assert.NotEmpty(mpxRows);
        Assert.All(xmlRows, row => Assert.Equal("Project.Xml", row.FormatId));
        Assert.All(binaryRows, row => Assert.Equal("Project.MppMpt", row.FormatId));
        Assert.All(mpxRows, row => Assert.Equal("Project.Mpx", row.FormatId));
        Assert.Contains(xmlRows, row => row.Operation == OfficeOperationKind.Validate && row.Limitation.Contains("output limits", StringComparison.Ordinal));
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
        OfficeOperationCapability multiExtensionRow = OfficeOperationCapabilityCatalog.All.First(row => row.Extensions.Count > 1);
        Assert.Contains("| Package | Format | Extensions | Target | Operation | State |", markdown, StringComparison.Ordinal);
        Assert.Contains(string.Join(", ", multiExtensionRow.Extensions), markdown, StringComparison.Ordinal);
        Assert.Contains("`OfficeIMO.Word.Pdf`", markdown, StringComparison.Ordinal);
        Assert.Contains("OfficeConversionCapabilityCatalog", markdown, StringComparison.Ordinal);
    }
}
