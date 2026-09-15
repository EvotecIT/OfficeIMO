using System.Text.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Security;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class OfficeCompatibilityCatalogContractTests {
    [Fact]
    public void ProtectedContentCapabilityPreservesTheOriginalConstructorSignature() {
        Type[] legacySignature = {
            typeof(string), typeof(string), typeof(string), typeof(OfficeProtectionKind),
            typeof(OfficeProtectionCoverageState), typeof(OfficeProtectionCoverageState),
            typeof(OfficeProtectionCoverageState), typeof(OfficeProtectionCoverageState),
            typeof(OfficeProtectionCoverageState), typeof(OfficeProtectionCoverageState),
            typeof(string), typeof(string)
        };

        Assert.NotNull(typeof(OfficeProtectionCapability).GetConstructor(legacySignature));
    }

    [Fact]
    public void BinaryFormatCatalogsExposeUniqueStableRowsAndValidFormatReferences() {
        OfficeCapabilityCatalog[] catalogs = {
            WordCompatibilityCatalog.Current,
            ExcelCompatibilityCatalog.Xls,
            ExcelCompatibilityCatalog.Xlsb,
            PowerPointCompatibilityCatalog.Current
        };
        var knownFormats = WordFormatCatalog.All
            .Concat(ExcelFormatCatalog.All)
            .Concat(PowerPointFormatCatalog.All)
            .Select(format => format.Id)
            .ToHashSet(StringComparer.Ordinal);

        foreach (OfficeCapabilityCatalog catalog in catalogs) {
            Assert.NotEmpty(catalog.Capabilities);
            Assert.Equal(
                catalog.Capabilities.Count,
                catalog.Capabilities.Select(capability => capability.Id).Distinct(StringComparer.Ordinal).Count());
            Assert.All(catalog.Capabilities, capability => {
                Assert.Contains(capability.FormatId, knownFormats);
                if (capability.GetState(OfficeCapabilityLane.LegacyToModern) == OfficeCapabilityCoverageState.Dropped
                    || capability.GetState(OfficeCapabilityLane.ModernToLegacy) == OfficeCapabilityCoverageState.Dropped) {
                    Assert.NotEqual(OfficeCompatibilityImpact.None, capability.AffectedFidelity);
                    Assert.False(string.IsNullOrWhiteSpace(capability.Note));
                }
            });
        }
    }

    [Fact]
    public void CapabilitySerializationIsDeterministicAndMachineReadable() {
        OfficeCapabilityCatalog catalog = ExcelCompatibilityCatalog.Xlsb;

        string first = catalog.ToJson();
        string second = catalog.ToJson();

        Assert.Equal(first, second);
        using JsonDocument parsed = JsonDocument.Parse(first);
        Assert.Equal(catalog.Id, parsed.RootElement.GetProperty("id").GetString());
        Assert.Equal(catalog.SchemaVersion, parsed.RootElement.GetProperty("schemaVersion").GetInt32());
        Assert.Equal(catalog.Capabilities.Count, parsed.RootElement.GetProperty("capabilities").GetArrayLength());
        Assert.Contains("| Legacy import |", catalog.ToMarkdown(), StringComparison.Ordinal);
    }

    [Fact]
    public void PowerPointSharedCatalogRetainsStaticVisualAndOpaqueDistinctions() {
        OfficeCapability chart = PowerPointCompatibilityCatalog.Current.Get("PowerPoint.Ppt.Charts");
        OfficeCapability unknown = PowerPointCompatibilityCatalog.Current.Get("PowerPoint.Ppt.UnknownRecordsAndStreams");

        Assert.Equal(OfficeCapabilityCoverageState.Rasterized, chart.ModernToLegacy);
        Assert.True(chart.AffectedFidelity.HasFlag(OfficeCompatibilityImpact.Editability));
        Assert.Equal(OfficeCapabilityCoverageState.PreservedOpaque, unknown.LegacyRoundTrip);
        Assert.True(unknown.AffectedFidelity.HasFlag(OfficeCompatibilityImpact.Carrier));
    }

    [Fact]
    public void ProtectedContentCatalogIsDeterministicAndKeepsNonCryptographicProtectionDistinct() {
        OfficeProtectionCapabilityCatalog catalog = OfficeProtectionCapabilityCatalog.Current;

        string first = catalog.ToJson();
        string second = catalog.ToJson();
        using JsonDocument parsed = JsonDocument.Parse(first);

        Assert.Equal(first, second);
        Assert.Equal(catalog.Capabilities.Count, parsed.RootElement.GetProperty("capabilities").GetArrayLength());
        Assert.Equal(OfficeProtectionKind.AccessDeterrence, catalog.Get("pst-password").Kind);
        Assert.Equal(OfficeProtectionKind.EditingRestriction, catalog.Get("rtf-editing-restrictions").Kind);
        Assert.Equal(OfficeProtectionCoverageState.Blocked, catalog.Get("onenote-encrypted-revision").Mutate);
        Assert.Equal(OfficeProtectionCoverageState.Supported, catalog.Get("odf-password").Create);
        Assert.Equal(OfficeProtectionCoverageState.NotApplicable, catalog.Get("epub-font-obfuscation").Mutate);
        Assert.Equal(OfficeProtectionCoverageState.NotSupported, catalog.Get("smime-signature-msg-tnef").Create);
        Assert.Equal(2, catalog.SchemaVersion);
        Assert.Equal(OfficeProtectionUnsupportedDisposition.RoadmapTracked,
            catalog.Get("xls-password").UnsupportedOperations.Single(item =>
                item.Operation == OfficeProtectionOperation.Create).Disposition);
        Assert.Equal(OfficeProtectionUnsupportedDisposition.IntentionalBoundary,
            catalog.Get("smime-signature-msg-tnef").UnsupportedOperations.Single().Disposition);
        Assert.Contains("| Inspect | Open | Create |", catalog.ToMarkdown(), StringComparison.Ordinal);
        Assert.Contains("[roadmap](../../ROADMAP.md#security-and-protected-content)", catalog.ToMarkdown(), StringComparison.Ordinal);
    }

    [Fact]
    public void ProtectedContentCatalogRejectsMissingOrExtraneousUnsupportedDispositions() {
        OfficeProtectionCapability MissingDisposition() => new(
            "missing", "DOC", "OfficeIMO.Word", OfficeProtectionKind.PasswordEncryption,
            OfficeProtectionCoverageState.Detected, OfficeProtectionCoverageState.NotSupported,
            OfficeProtectionCoverageState.NotApplicable, OfficeProtectionCoverageState.NotApplicable,
            OfficeProtectionCoverageState.Blocked, OfficeProtectionCoverageState.NotApplicable,
            "WordDocument.Load", "Missing disposition");
        OfficeProtectionCapability ExtraDisposition() => new(
            "extra", "PDF", "OfficeIMO.Pdf", OfficeProtectionKind.PasswordEncryption,
            OfficeProtectionCoverageState.Supported, OfficeProtectionCoverageState.Supported,
            OfficeProtectionCoverageState.Supported, OfficeProtectionCoverageState.NotApplicable,
            OfficeProtectionCoverageState.Supported, OfficeProtectionCoverageState.Supported,
            "PdfDocument.Security", "Extraneous disposition", new[] {
                new OfficeProtectionUnsupportedOperation(
                    OfficeProtectionOperation.Open,
                    OfficeProtectionUnsupportedDisposition.IntentionalBoundary,
                    "This operation is actually supported.")
            });

        Assert.Throws<ArgumentException>(() =>
            new OfficeProtectionCapabilityCatalog("missing-catalog", 1, new[] { MissingDisposition() }));
        Assert.Throws<ArgumentException>(() =>
            new OfficeProtectionCapabilityCatalog("extra-catalog", 1, new[] { ExtraDisposition() }));
    }

    [Fact]
    public void ProtectedContentCatalogEscapesEveryJsonControlCharacter() {
        var row = new OfficeProtectionCapability(
            "control-row", "EML\tformat", "OfficeIMO.Email", OfficeProtectionKind.DigitalSignature,
            OfficeProtectionCoverageState.Supported, OfficeProtectionCoverageState.Supported,
            OfficeProtectionCoverageState.NotSupported, OfficeProtectionCoverageState.NotApplicable,
            OfficeProtectionCoverageState.Preserved, OfficeProtectionCoverageState.NotApplicable,
            "Verify\u0001Api", "line\bfeed\f", new[] {
                new OfficeProtectionUnsupportedOperation(
                    OfficeProtectionOperation.Create,
                    OfficeProtectionUnsupportedDisposition.IntentionalBoundary,
                    "Control disposition\u0003reason")
            });
        var catalog = new OfficeProtectionCapabilityCatalog("control\u0002catalog", 1, new[] { row });

        using JsonDocument parsed = JsonDocument.Parse(catalog.ToJson());

        Assert.Equal("EML\tformat", parsed.RootElement.GetProperty("capabilities")[0].GetProperty("formatId").GetString());
        Assert.Equal("Control disposition\u0003reason", parsed.RootElement.GetProperty("capabilities")[0]
            .GetProperty("unsupportedOperations")[0].GetProperty("rationale").GetString());
    }

    [Fact]
    public void ProtectedContentCatalogEscapesRoadmapReferencesInMarkdownLinks() {
        var row = new OfficeProtectionCapability(
            "roadmap-row", "DOC", "OfficeIMO.Word", OfficeProtectionKind.PasswordEncryption,
            OfficeProtectionCoverageState.Detected, OfficeProtectionCoverageState.NotSupported,
            OfficeProtectionCoverageState.NotApplicable, OfficeProtectionCoverageState.NotApplicable,
            OfficeProtectionCoverageState.Blocked, OfficeProtectionCoverageState.NotApplicable,
            "WordDocument.Load", "Roadmap escaping", new[] {
                new OfficeProtectionUnsupportedOperation(
                    OfficeProtectionOperation.Open,
                    OfficeProtectionUnsupportedDisposition.RoadmapTracked,
                    "Open support is tracked.",
                    "../../ROAD|MAP(1).md#open\r\nnext")
            });
        var catalog = new OfficeProtectionCapabilityCatalog("roadmap-catalog", 2, new[] { row });

        string markdown = catalog.ToMarkdown();

        Assert.Contains(
            "[roadmap](../../ROAD%7CMAP%281%29.md#open%0D%0Anext)",
            markdown,
            StringComparison.Ordinal);
        Assert.DoesNotContain("ROAD|MAP", markdown, StringComparison.Ordinal);
    }
}
